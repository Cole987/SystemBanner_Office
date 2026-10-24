//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //the text of the document split in portions: those that start with a portion mark, and how many have none
    public class PortionScan
    {
        public List<Marking> Marks = new List<Marking>();               //IC portion marks, "(S//NF)"
        public List<FrenchMarking> FrenchMarks = new List<FrenchMarking>(); //IGI 1300 paragraph marks, "[S SF]"
        public List<GermanMarking> GermanMarks = new List<GermanMarking>(); //VS-Anweisung part marks, "[Geh.]"
        public List<EuMarking> EuMarks = new List<EuMarking>();             //EU paragraph abbreviations, "[R-UE/EU-R]"

        //how many marking systems the document's marks belong to
        public int Systems
        {
            get { return (Marks.Count > 0 ? 1 : 0) + (FrenchMarks.Count > 0 ? 1 : 0) + (GermanMarks.Count > 0 ? 1 : 0) + (EuMarks.Count > 0 ? 1 : 0); }
        }
        public List<string> Errors = new List<string>();
        public int Unmarked;
    }

    //one line of a stamp at the top or bottom of the pages
    public class StampLine
    {
        public string Text;
        public int Color;        //BGR text color
        public int Size;         //points, 0 to keep the document's size
        public float Frame;      //frame width in points, 0 for none
        public StampLine(string text, int color, int size, float frame)
        {
            Text = text;
            Color = color;
            Size = size;
            Frame = frame;
        }
    }

    //what goes at the top and bottom of every page: the IC banner line in the color of its level,
    //or the IGI 1300 stamps, framed text in red, Spécial France in blue
    public class StampSet
    {
        public List<StampLine> Top = new List<StampLine>();
        public List<StampLine> Bottom = new List<StampLine>();
        public List<StampLine> Cover;    //bound documents: bottom of the first page instead of Bottom, null otherwise
        public int Background = -1;      //BGR background of the lines, -1 for none (IGI stamps are on the page itself)
        public bool PageNumbers;         //"page X / Y" under the bottom stamp
        public string PageFormat = "page {P} / {N}"; //{P} the page, {N} the page count
        public string Reference;         //TRÈS SECRET: the document reference on every page, before the page number

        public static readonly int Red = Late.Rgb(200, 16, 46);
        public static readonly int Blue = Late.Rgb(0, 51, 160);
        public const float PageFrame = 2.5f;   //annexe 37 § 2
        public const float CoverFrame = 3f;    //annexe 37 § 1

        //OfficeBannerBackground = 1: the banner line on the colored background of SystemBanner, as on screen;
        //by default its text in the color of the level, with no background: a background a printer or a PDF converter
        //leaves out would leave white text, an invisible banner, and black and white printing turns it into a pale grey
        public static bool BannerBackground;

        public static StampSet Banner(string text, int background, int foreground)
        {
            StampSet stamps = new StampSet();
            int color = BannerBackground ? foreground : TextColor(background);
            stamps.Top.Add(new StampLine(text, color, 0, 0));
            stamps.Bottom.Add(new StampLine(text, color, 0, 0));
            stamps.Background = BannerBackground ? background : -1;
            return stamps;
        }

        //the color of the level as a text color on white paper; a light one (the yellow of TOP SECRET//SCI) is written in black
        public static int TextColor(int bgr)
        {
            int red = bgr & 0xFF, green = (bgr >> 8) & 0xFF, blue = (bgr >> 16) & 0xFF;
            return red * 299 + green * 587 + blue * 114 > 160000 ? 0 : bgr;
        }

        //IGI 1300, 7.1.2.3 and annexes 1 and 37
        public static StampSet French(FrenchMarking m, FrenchIdentification id, bool bound)
        {
            StampSet stamps = new StampSet();
            if (m.Level == FrenchLevel.NonProtege)
            {
                return stamps; //no stamp: whatever stamps were there are removed
            }
            stamps.Top.Add(new StampLine(FrenchFormatter.Stamp(m), Red, 18, PageFrame));
            if (m.SpecialFrance)
            {
                stamps.Top.Add(new StampLine(FrenchFormatter.SpecialFranceStamp, Blue, 18, PageFrame)); //at the top only, under the stamp
            }
            string bottom = FrenchFormatter.BottomStamp(m);
            if (bottom.Length > 0)
            {
                stamps.Bottom.Add(new StampLine(bottom, Red, 18, PageFrame));
                stamps.PageNumbers = true; //7.1.2.3 c): every page of a classified document is numbered
            }
            if (m.Level == FrenchLevel.TresSecret && id != null && !string.IsNullOrEmpty(id.Reference))
            {
                stamps.Reference = id.Reference;
            }
            if (bound && m.IsClassified)
            {
                stamps.Cover = new List<StampLine>();
                stamps.Cover.Add(new StampLine(FrenchFormatter.Stamp(m), Red, 18, CoverFrame));
                stamps.Cover.Add(new StampLine(FrenchFormatter.CoverWarning, Red, 6, 0));
            }
            return stamps;
        }

        //VS-Anweisung, Anlage IV § 1.1: STRENG GEHEIM and GEHEIM in red at the top and bottom, with "amtlich geheim gehalten";
        //VS-VERTRAULICH (with it) and VS-NfD (without) at the top only; Seite X von Y on every page (§ 20 Abs. 2 Nr. 8)
        public static StampSet German(GermanMarking m)
        {
            StampSet stamps = new StampSet();
            if (!m.IsClassified)
            {
                return stamps;
            }
            int[] rgb = GermanFormatter.Color(m.Level);
            int color = Late.Rgb(rgb[0], rgb[1], rgb[2]);
            List<StampLine> stamp = new List<StampLine>();
            stamp.Add(new StampLine(GermanFormatter.Stamp(m.Level), color, 14, 0));
            if (m.AmtlichGeheimGehalten)
            {
                stamp.Add(new StampLine(GermanFormatter.AmtlichGeheimGehalten, color, 10, 0));
            }
            stamps.Top.AddRange(stamp);
            foreach (string vermerk in m.Vermerke)
            {
                stamps.Top.Add(new StampLine(vermerk, color, 10, 0));
            }
            if (GermanFormatter.AtBottom(m.Level))
            {
                stamps.Bottom.AddRange(stamp);
            }
            stamps.PageNumbers = true;
            stamps.PageFormat = "Seite {P} von {N}";
            return stamps;
        }

        //GSCP: the classification and its additional markings at the top and bottom, in capitals, with no colour coding
        public static StampSet British(UkMarking m)
        {
            StampSet stamps = new StampSet();
            string text = UkFormatter.Marking(m);
            stamps.Top.Add(new StampLine(text, 0, 12, 0));
            if (m.HandlingInstruction == UkRegister.Embargoed && !string.IsNullOrEmpty(m.Embargo))
            {
                stamps.Top.Add(new StampLine("EMBARGOED: " + m.Embargo, 0, 10, 0)); //as close as possible to the classification
            }
            stamps.Bottom.Add(new StampLine(text, 0, 12, 0));
            return stamps;
        }

        //EUCI: each page marked with the level (CSR Annex III pt 13(a)) and numbered (13(b)); the texts set no place, colour or size:
        //top and bottom in black like the GSCP, the additional markings under the top one, the reference and copy number by the page number
        public static StampSet Eu(EuMarking m)
        {
            StampSet stamps = new StampSet();
            string text = EuFormatter.Marking(m.Level);
            stamps.Top.Add(new StampLine(text, 0, 12, 0));
            foreach (string line in m.Additional)
            {
                stamps.Top.Add(new StampLine(line, 0, 10, 0));
            }
            stamps.Bottom.Add(new StampLine(text, 0, 12, 0));
            stamps.PageNumbers = true;
            stamps.PageFormat = "{P} / {N}";
            string footer = EuFormatter.Footer(m);
            if (footer.Length > 0) stamps.Reference = footer;
            return stamps;
        }
    }

    //what the add-in does in a document, the same way in Word, Excel and PowerPoint; every call is late bound
    public abstract class DocumentHost
    {
        //document properties, the classification metadata other tools (mail guards, file servers) can read
        public const string ClassificationProperty = "Classification";
        public const string SystemProperty = "SystemBanner.System";   //"IC" or "France"
        public const string BannerProperty = "SystemBanner.Banner";
        public const string ClassifiedByProperty = "SystemBanner.ClassifiedBy";
        public const string AgencyProperty = "SystemBanner.Agency";
        public const string ReasonProperty = "SystemBanner.Reason";
        public const string DerivedFromProperty = "SystemBanner.DerivedFrom";
        public const string DeclassifyOnProperty = "SystemBanner.DeclassifyOn";
        public const string DateOfOriginProperty = "SystemBanner.DateOfOrigin";
        public const string OriginalProperty = "SystemBanner.Original";
        public const string SourcesProperty = "SystemBanner.Sources";
        private const string Fr = "SystemBanner.FR.";

        public const string IcSystem = "IC";
        public const string FrenchSystem = "France";
        public const string GermanSystem = GermanSettings.Germany;
        public const string BritishSystem = UkSettings.UnitedKingdom;
        public const string EuSystem = EuSettings.EuropeanUnion;

        //a portion mark at the start of a paragraph: "(S//NF) text" or "[S SF] text"
        private static readonly Regex PortionStart = new Regex(@"^\s*\(([^()\r\n]{1,300})\)");
        private static readonly Regex FrenchStart = new Regex(@"^\s*\[([A-Z ]{1,12})\]");
        private static readonly Regex GermanStart = new Regex(@"^\s*\[(Str\. Geh\.|Geh\.|VS-Vertr\.|VS-NfD|offen)\]");

        public object Application;
        public object Document;

        //the host of the active document, null when no document is open
        public static DocumentHost Active(object application)
        {
            DocumentHost host = Create(application);
            if (host == null)
            {
                return null;
            }
            host.Document = Late.TryGet(application, host is WordHost ? "ActiveDocument" : host is ExcelHost ? "ActiveWorkbook" : "ActivePresentation");
            return host.Document == null ? null : host;
        }

        //the host of a given document, from the events that pass it
        public static DocumentHost For(object application, object document)
        {
            DocumentHost host = Create(application);
            if (host != null)
            {
                host.Document = document;
            }
            return host;
        }

        private static DocumentHost Create(object application)
        {
            string name = Convert.ToString(Late.Get(application, "Name"));
            DocumentHost host;
            if (name.IndexOf("Word") >= 0) host = new WordHost();
            else if (name.IndexOf("Excel") >= 0) host = new ExcelHost();
            else if (name.IndexOf("PowerPoint") >= 0) host = new PowerPointHost();
            else return null;
            host.Application = application;
            return host;
        }

        //major version of Office: 9 for Office 2000, 12 for 2007, 14 for 2010
        public int Version
        {
            get
            {
                string version = Convert.ToString(Late.Get(Application, "Version"));
                int dot = version.IndexOf('.');
                int major;
                return int.TryParse(dot > 0 ? version.Substring(0, dot) : version, out major) ? major : 0;
            }
        }

        public string GetProperty(string name)
        {
            object properties = Late.Get(Document, "CustomDocumentProperties");
            try
            {
                object property = Late.Item(properties, name);
                return Convert.ToString(Late.Get(property, "Value"));
            }
            catch (Exception)
            {
                return null; //no such property
            }
        }

        public void SetProperty(string name, string value)
        {
            object properties = Late.Get(Document, "CustomDocumentProperties");
            try
            {
                object property = Late.Item(properties, name);
                if (string.IsNullOrEmpty(value))
                {
                    Late.Call(property, "Delete");
                }
                else
                {
                    Late.Set(property, "Value", value);
                }
            }
            catch (Exception)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    //msoPropertyTypeString = 4, not linked to content
                    Late.Call(properties, "Add", name, false, 4, value);
                }
            }
        }

        //"IC" or "France" once the document is classified, null before
        public string StoredSystem()
        {
            string system = GetProperty(SystemProperty);
            if (string.IsNullOrEmpty(system) && !string.IsNullOrEmpty(GetProperty(BannerProperty)))
            {
                return IcSystem; //documents marked before French markings existed
            }
            return string.IsNullOrEmpty(system) ? null : system;
        }

        //the banner stored in the document, null when it is not classified yet
        public Marking StoredBanner()
        {
            string text = GetProperty(BannerProperty);
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }
            ParseResult result = MarkingParser.Parse(text);
            return result.Success ? result.Marking : null;
        }

        public AuthorityBlock StoredAuthorityBlock()
        {
            AuthorityBlock block = new AuthorityBlock();
            block.ClassifiedBy = GetProperty(ClassifiedByProperty);
            block.AgencyAndOffice = GetProperty(AgencyProperty);
            block.Reason = GetProperty(ReasonProperty);
            block.DerivedFrom = GetProperty(DerivedFromProperty);
            block.DeclassifyOn = GetProperty(DeclassifyOnProperty);
            block.Original = GetProperty(OriginalProperty) == "1";
            DateTime origin;
            if (DateTime.TryParseExact(GetProperty(DateOfOriginProperty) ?? "", "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out origin))
            {
                block.DateOfOrigin = origin;
            }
            string sources = GetProperty(SourcesProperty);
            if (!string.IsNullOrEmpty(sources))
            {
                block.Sources.AddRange(sources.Split('\n'));
            }
            return block;
        }

        public void StoreAuthorityBlock(AuthorityBlock block)
        {
            SetProperty(ClassifiedByProperty, block.ClassifiedBy);
            SetProperty(AgencyProperty, block.AgencyAndOffice);
            SetProperty(ReasonProperty, block.Original ? block.Reason : null);
            SetProperty(DerivedFromProperty, block.Original ? null : block.DerivedFrom);
            SetProperty(DeclassifyOnProperty, block.DeclassifyOn);
            SetProperty(OriginalProperty, block.Original ? "1" : "0");
            SetProperty(DateOfOriginProperty, block.DateOfOrigin.ToString("yyyyMMdd"));
            SetProperty(SourcesProperty, block.Sources.Count > 0 ? string.Join("\n", block.Sources.ToArray()) : null);
        }

        //writes the banner line at the top and bottom of every page and stores it in the document properties
        public void Mark(Marking banner)
        {
            string text = MarkingFormatter.Banner(banner);
            SetProperty(SystemProperty, IcSystem);
            SetProperty(BannerProperty, text);
            SetProperty(ClassificationProperty, MarkingFormatter.LevelTitle(banner.Level == Level.Restricted ? Level.Confidential : banner.Level));
            int[] colors = Presets.ColorsFor(banner);
            WriteStamps(StampSet.Banner(text, Late.Rgb(colors[0], colors[1], colors[2]), Late.Rgb(colors[3], colors[4], colors[5])));
        }

        //a national banner of SystemBanner, on computers set to the German, British, NATO or UN systems
        public void Mark(Preset preset)
        {
            SetProperty(SystemProperty, preset.ClassificationSystem);
            SetProperty(BannerProperty, null);
            SetProperty(ClassificationProperty, preset.Text);
            WriteStamps(StampSet.Banner(preset.Text, Late.Rgb(preset.Colors[0], preset.Colors[1], preset.Colors[2]), Late.Rgb(preset.Colors[3], preset.Colors[4], preset.Colors[5])));
        }

        //the IGI 1300 stamps, and the French marking in the document properties
        public void Mark(FrenchMarking marking, FrenchIdentification id, bool bound)
        {
            SetProperty(SystemProperty, FrenchSystem);
            SetProperty(BannerProperty, null);
            SetProperty(ClassificationProperty, marking.Level == FrenchLevel.NonProtege ? null : marking.ToString());
            SetProperty(Fr + "Niveau", FrenchFormatter.Abbreviation(marking.Level));
            SetProperty(Fr + "SpecialFrance", marking.SpecialFrance ? "1" : null);
            SetProperty(Fr + "Mention", marking.MentionParticuliere);
            SetProperty(Fr + "Relie", bound ? "1" : null);
            WriteStamps(StampSet.French(marking, id, bound && SupportsCover));
        }

        private const string De = "SystemBanner.DE.";
        private const string Uk = "SystemBanner.UK.";

        //the VS-Anweisung stamps, and the German marking in the document properties
        public void Mark(GermanMarking marking)
        {
            SetProperty(SystemProperty, GermanSystem);
            SetProperty(BannerProperty, null);
            SetProperty(ClassificationProperty, marking.IsClassified ? GermanFormatter.Stamp(marking.Level) : null);
            SetProperty(De + "Grad", GermanFormatter.Abbreviation(marking.Level));
            SetProperty(De + "Vermerke", marking.Vermerke.Count > 0 ? string.Join("\n", marking.Vermerke.ToArray()) : null);
            WriteStamps(StampSet.German(marking));
        }

        public GermanMarking StoredGermanMarking()
        {
            string level = GetProperty(De + "Grad");
            if (string.IsNullOrEmpty(level)) return null;
            GermanMarking m = GermanParser.ParsePortion("[" + level + "]") ?? new GermanMarking();
            string vermerke = GetProperty(De + "Vermerke");
            if (!string.IsNullOrEmpty(vermerke)) m.Vermerke.AddRange(vermerke.Split('\n'));
            return m;
        }

        public GermanIdentification StoredGermanIdentification()
        {
            GermanIdentification id = new GermanIdentification();
            id.Herausgeber = GetProperty(De + "Herausgeber");
            id.Kennzeichen = GetProperty(De + "Kennzeichen");
            id.Empfaenger = GetProperty(De + "Empfaenger");
            DateTime date;
            if (TryDate(GetProperty(De + "Datum"), out date)) id.Datum = date;
            int number;
            if (int.TryParse(GetProperty(De + "EinstufungEndet"), out number)) id.EinstufungEndet = number;
            if (int.TryParse(GetProperty(De + "Ausfertigung"), out number)) id.Ausfertigung = number;
            return id;
        }

        public void StoreGermanIdentification(GermanIdentification id)
        {
            SetProperty(De + "Herausgeber", id.Herausgeber);
            SetProperty(De + "Kennzeichen", id.Kennzeichen);
            SetProperty(De + "Empfaenger", id.Empfaenger);
            SetProperty(De + "Datum", id.Datum.ToString("yyyyMMdd"));
            SetProperty(De + "EinstufungEndet", id.EinstufungEndet > 0 ? id.EinstufungEndet.ToString() : null);
            SetProperty(De + "Ausfertigung", id.Ausfertigung > 0 ? id.Ausfertigung.ToString() : null);
        }

        //the GSCP markings, and the British marking in the document properties
        public void Mark(UkMarking marking)
        {
            SetProperty(SystemProperty, BritishSystem);
            SetProperty(BannerProperty, null);
            SetProperty(ClassificationProperty, UkFormatter.Classification(marking));
            SetProperty(Uk + "Level", ((int)marking.Level).ToString());
            SetProperty(Uk + "Sensitive", marking.Sensitive ? "1" : null);
            SetProperty(Uk + "Prefix", marking.UkPrefix ? "1" : null);
            SetProperty(Uk + "RelEu", marking.RelEu ? "1" : null);
            SetProperty(Uk + "Caveats", marking.Caveats.Count > 0 ? string.Join("\n", marking.Caveats.ToArray()) : null);
            SetProperty(Uk + "Handling", marking.HandlingInstruction);
            SetProperty(Uk + "Organisation", marking.Organisation);
            SetProperty(Uk + "Embargo", marking.Embargo);
            SetProperty(Uk + "Codeword", marking.Codeword);
            SetProperty(Uk + "Descriptors", marking.Descriptors.Count > 0 ? string.Join("\n", marking.Descriptors.ToArray()) : null);
            SetProperty(Uk + "Marking", UkFormatter.Marking(marking));
            WriteStamps(StampSet.British(marking));
        }

        private const string Eu = "SystemBanner.EU.";

        //the EU markings, and the EU marking in the document properties
        public void Mark(EuMarking marking)
        {
            SetProperty(SystemProperty, EuSystem);
            SetProperty(BannerProperty, null);
            SetProperty(ClassificationProperty, EuFormatter.Marking(marking.Level));
            SetProperty(Eu + "Level", EuFormatter.Abbreviation(marking.Level));
            SetProperty(Eu + "Additional", marking.Additional.Count > 0 ? string.Join("\n", marking.Additional.ToArray()) : null);
            SetProperty(Eu + "Reference", marking.Reference);
            SetProperty(Eu + "Copy", marking.Copy > 0 ? marking.Copy.ToString() : null);
            WriteStamps(StampSet.Eu(marking));
        }

        public EuMarking StoredEuMarking()
        {
            string level = GetProperty(Eu + "Level");
            if (string.IsNullOrEmpty(level)) return null;
            EuMarking m = EuParser.ParsePortion("[" + level + "]") ?? new EuMarking();
            string additional = GetProperty(Eu + "Additional");
            if (!string.IsNullOrEmpty(additional)) m.Additional.AddRange(additional.Split('\n'));
            m.Reference = GetProperty(Eu + "Reference");
            int copy;
            if (int.TryParse(GetProperty(Eu + "Copy"), out copy)) m.Copy = copy;
            return m;
        }

        public UkMarking StoredUkMarking()
        {
            string level = GetProperty(Uk + "Level");
            if (string.IsNullOrEmpty(level)) return null;
            UkMarking m = new UkMarking();
            int number;
            if (int.TryParse(level, out number) && number >= 0 && number <= 2) m.Level = (UkLevel)number;
            m.Sensitive = GetProperty(Uk + "Sensitive") == "1";
            m.UkPrefix = GetProperty(Uk + "Prefix") == "1";
            m.RelEu = GetProperty(Uk + "RelEu") == "1";
            string caveats = GetProperty(Uk + "Caveats");
            if (!string.IsNullOrEmpty(caveats)) m.Caveats.AddRange(caveats.Split('\n'));
            m.HandlingInstruction = GetProperty(Uk + "Handling");
            m.Organisation = GetProperty(Uk + "Organisation");
            m.Embargo = GetProperty(Uk + "Embargo");
            m.Codeword = GetProperty(Uk + "Codeword");
            string descriptors = GetProperty(Uk + "Descriptors");
            if (!string.IsNullOrEmpty(descriptors)) m.Descriptors.AddRange(descriptors.Split('\n'));
            return m;
        }

        public FrenchMarking StoredFrenchMarking()
        {
            string level = GetProperty(Fr + "Niveau");
            if (string.IsNullOrEmpty(level))
            {
                return null;
            }
            FrenchMarking m = FrenchParser.ParsePortion("[" + level + "]") ?? new FrenchMarking();
            m.SpecialFrance = GetProperty(Fr + "SpecialFrance") == "1";
            m.MentionParticuliere = GetProperty(Fr + "Mention");
            return m;
        }

        public bool StoredBound()
        {
            return GetProperty(Fr + "Relie") == "1";
        }

        //the date the document was declassified (annexe 38), null when it is not
        public string StoredDeclassification()
        {
            return GetProperty(Fr + "Declassifie");
        }

        public void StoreDeclassification(string text)
        {
            SetProperty(Fr + "Declassifie", text);
        }

        public FrenchIdentification StoredFrenchIdentification()
        {
            FrenchIdentification id = new FrenchIdentification();
            id.Lieu = GetProperty(Fr + "Lieu");
            id.NumeroEnregistrement = GetProperty(Fr + "Enregistrement");
            id.AutoriteEmettrice = GetProperty(Fr + "Autorite");
            id.Auteur = GetProperty(Fr + "Auteur");
            id.Reference = GetProperty(Fr + "Reference");
            id.Annexes = GetProperty(Fr + "Annexes");
            DateTime date;
            if (TryDate(GetProperty(Fr + "DateEmission"), out date)) id.DateEmission = date;
            if (TryDate(GetProperty(Fr + "Echeance"), out date)) id.Echeance = date;
            if (TryDate(GetProperty(Fr + "Reevaluation"), out date)) id.Reevaluation = date;
            id.SansEcheance = GetProperty(Fr + "SansEcheance") == "1";
            int number;
            if (int.TryParse(GetProperty(Fr + "Exemplaire"), out number)) id.Exemplaire = number;
            if (int.TryParse(GetProperty(Fr + "NombreExemplaires"), out number)) id.NombreExemplaires = number;
            return id;
        }

        public void StoreFrenchIdentification(FrenchIdentification id)
        {
            SetProperty(Fr + "Lieu", id.Lieu);
            SetProperty(Fr + "Enregistrement", id.NumeroEnregistrement);
            SetProperty(Fr + "Autorite", id.AutoriteEmettrice);
            SetProperty(Fr + "Auteur", id.Auteur);
            SetProperty(Fr + "Reference", id.Reference);
            SetProperty(Fr + "Annexes", id.Annexes);
            SetProperty(Fr + "DateEmission", id.DateEmission.ToString("yyyyMMdd"));
            SetProperty(Fr + "Echeance", id.Echeance.HasValue ? id.Echeance.Value.ToString("yyyyMMdd") : null);
            SetProperty(Fr + "Reevaluation", id.Reevaluation.HasValue ? id.Reevaluation.Value.ToString("yyyyMMdd") : null);
            SetProperty(Fr + "SansEcheance", id.SansEcheance ? "1" : null);
            SetProperty(Fr + "Exemplaire", id.Exemplaire > 0 ? id.Exemplaire.ToString() : null);
            SetProperty(Fr + "NombreExemplaires", id.NombreExemplaires > 0 ? id.NombreExemplaires.ToString() : null);
        }

        private static bool TryDate(string text, out DateTime date)
        {
            return DateTime.TryParseExact(text ?? "", "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        public string StoredClassification()
        {
            return GetProperty(ClassificationProperty);
        }

        protected abstract void WriteStamps(StampSet stamps);

        //only Word has a first page that can carry the cover stamp of bound documents
        public virtual bool SupportsCover
        {
            get { return false; }
        }

        public virtual bool SupportsPortions
        {
            get { return false; }
        }

        //the paragraphs of the document, for the roll-up
        protected virtual IEnumerable<string> Paragraphs()
        {
            return new string[0];
        }

        public PortionScan ScanPortions()
        {
            PortionScan scan = new PortionScan();
            foreach (string paragraph in Paragraphs())
            {
                string text = paragraph.Trim('\a', '\v', '\f', ' ', '\t', '\r', '\n', '\u0007');
                if (text.Length == 0)
                {
                    continue;
                }
                Match eu = EuParser.PortionStart.Match(text);
                if (eu.Success)
                {
                    EuMarking mark = EuParser.ParsePortion(eu.Value.Trim());
                    if (mark != null)
                    {
                        scan.EuMarks.Add(mark);
                    }
                    else
                    {
                        scan.Errors.Add(eu.Value.Trim() + " " + Text.T(": not a valid EU paragraph mark; the marks are [TS-UE/EU-TS], [S-UE/EU-S], [C-UE/EU-C] and [R-UE/EU-R].", " : marque de paragraphe UE non valide ; les marques sont [TS-UE/EU-TS], [S-UE/EU-S], [C-UE/EU-C] et [R-UE/EU-R].", ": keine gültige EU-Absatzkennzeichnung; die Kennzeichnungen sind [TS-UE/EU-TS], [S-UE/EU-S], [C-UE/EU-C] und [R-UE/EU-R]."));
                    }
                    continue;
                }
                Match german = GermanStart.Match(text);
                if (german.Success)
                {
                    scan.GermanMarks.Add(GermanParser.ParsePortion("[" + german.Groups[1].Value + "]"));
                    continue;
                }
                Match french = FrenchStart.Match(text);
                if (french.Success)
                {
                    FrenchMarking mark = FrenchParser.ParsePortion("[" + french.Groups[1].Value + "]");
                    if (mark != null)
                    {
                        scan.FrenchMarks.Add(mark);
                    }
                    else if (Regex.IsMatch(french.Groups[1].Value, @"^(NP|DR|S|TS|CD|SD|TSD)\b"))
                    {
                        scan.Errors.Add("[" + french.Groups[1].Value + "] " + Text.T(": not a valid paragraph mark; the marks are [NP], [DR], [S] and [TS], followed by SF for Spécial France (Confidentiel Défense no longer exists since 1 July 2021).", ": marque de paragraphe non valide ; les marques sont [NP], [DR], [S] et [TS], suivies de SF pour Spécial France (Confidentiel Défense n'existe plus depuis le 1er juillet 2021).", ": keine gültige Absatzkennzeichnung; die Kennzeichnungen sind [NP], [DR], [S] und [TS], gefolgt von SF für Spécial France (Confidentiel Défense gibt es seit dem 1. Juli 2021 nicht mehr)."));
                    }
                    else
                    {
                        scan.Unmarked++;
                    }
                    continue;
                }
                Match match = PortionStart.Match(text);
                if (!match.Success)
                {
                    scan.Unmarked++;
                    continue;
                }
                ParseResult result = MarkingParser.Parse(match.Groups[1].Value);
                if (result.Success)
                {
                    scan.Marks.Add(result.Marking);
                }
                else if (LooksLikeMarking(match.Groups[1].Value))
                {
                    scan.Errors.Add("(" + match.Groups[1].Value + "): " + string.Join(" ", result.Errors.ToArray()));
                }
                else
                {
                    scan.Unmarked++; //a parenthesis that is not a portion mark, "(see annex)"
                }
            }
            return scan;
        }

        private static bool LooksLikeMarking(string text)
        {
            return text.StartsWith("//") || Regex.IsMatch(text, @"^(TS|S|C|U|TOP SECRET|SECRET|CONFIDENTIAL|UNCLASSIFIED)(//|$)");
        }

        //puts the mark ("(S//NF)" or "[S SF]") at the start of the paragraph of the cursor, replacing the one already there
        public virtual void InsertPortionMark(string mark)
        {
            throw new NotSupportedException();
        }

        //writes lines at the cursor, or in place of the block written before under the same bookmark (Word only);
        //"{NUMPAGES}" in a line becomes the page count field
        public virtual void WriteBlock(string bookmark, string[] lines)
        {
            throw new NotSupportedException();
        }

        //writes the IC classification authority block in a framed box at the bottom right of the first page (Word only),
        //in place of the box or the block written before; the Register allows it anywhere on the first page or cover
        public virtual void WriteAuthorityBox(string[] lines)
        {
            throw new NotSupportedException();
        }

        public const string AuthorityBookmark = "SystemBannerAuthority";
        public const string IdentificationBookmark = "SystemBannerIdentification";
        public const string ChangeBookmark = "SystemBannerDeclassement";

        //length of the portion mark at the start of a paragraph's text, with the space after it, 0 when there is none
        protected static int ExistingMarkLength(string paragraphText)
        {
            Match match = EuParser.PortionStart.Match(paragraphText);
            bool found = match.Success;
            if (!found)
            {
                match = PortionStart.Match(paragraphText);
                found = match.Success && MarkingParser.Parse(match.Groups[1].Value).Success;
            }
            if (!found)
            {
                match = FrenchStart.Match(paragraphText);
                found = match.Success && FrenchParser.ParsePortion("[" + match.Groups[1].Value + "]") != null;
            }
            if (!found)
            {
                match = GermanStart.Match(paragraphText);
                found = match.Success;
            }
            if (!found)
            {
                return 0;
            }
            int length = match.Length;
            if (length < paragraphText.Length && paragraphText[length] == ' ')
            {
                length++;
            }
            return length;
        }
    }

    public class WordHost : DocumentHost
    {
        private const int HeaderFooterPrimary = 1, HeaderFooterFirstPage = 2, HeaderFooterEvenPages = 3;
        private const int FieldPage = 33, FieldNumPages = 26;
        private const int ColorAutomatic = -16777216;

        public override bool SupportsPortions
        {
            get { return true; }
        }

        public override bool SupportsCover
        {
            get { return true; }
        }

        protected override void WriteStamps(StampSet stamps)
        {
            object sections = Late.Get(Document, "Sections");
            int count = Late.Count(sections);
            if (stamps.Cover != null)
            {
                //the cover stamp goes on the first page only: its own footer
                Late.TrySet(Late.Get(Late.Item(sections, 1), "PageSetup"), "DifferentFirstPageHeaderFooter", -1);
            }
            for (int s = 1; s <= count; s++)
            {
                object section = Late.Item(sections, s);
                foreach (int kind in new int[] { HeaderFooterPrimary, HeaderFooterFirstPage, HeaderFooterEvenPages })
                {
                    bool cover = stamps.Cover != null && s == 1 && kind == HeaderFooterFirstPage;
                    object footer = Late.Item(Late.Get(section, "Footers"), kind);
                    WriteHeaderFooter(Late.Item(Late.Get(section, "Headers"), kind), "SystemBannerTop" + s + "_" + kind, true, stamps.Top, stamps.Background);
                    WriteHeaderFooter(footer, "SystemBannerBottom" + s + "_" + kind, false, cover ? stamps.Cover : stamps.Bottom, stamps.Background);
                    WritePageNumbers(footer, "SystemBannerPages" + s + "_" + kind, stamps.PageNumbers, stamps.Reference, stamps.PageFormat);
                }
            }
        }

        //the stamps are their own paragraphs, first in the header and last in the footer, kept in a bookmark so the user's content stays
        private void WriteHeaderFooter(object headerFooter, string bookmark, bool top, List<StampLine> lines, int background)
        {
            if (Convert.ToBoolean(Late.TryGet(headerFooter, "LinkToPrevious") ?? false))
            {
                return; //the previous section's header already shows it
            }
            object bookmarks = Late.Get(Document, "Bookmarks");
            bool exists = Convert.ToBoolean(Late.Call(bookmarks, "Exists", bookmark));
            if (lines.Count == 0)
            {
                if (exists)
                {
                    RemoveBlock(headerFooter, Late.Get(Late.Item(bookmarks, bookmark), "Range"));
                }
                return;
            }
            object range = exists ? Late.Get(Late.Item(bookmarks, bookmark), "Range") : NewParagraph(headerFooter, top);
            List<string> texts = new List<string>();
            foreach (StampLine line in lines)
            {
                texts.Add(line.Text);
            }
            Late.Set(range, "Text", string.Join("\r", texts.ToArray()));
            object paragraphs = Late.Get(range, "Paragraphs");
            int paragraphCount = Late.Count(paragraphs);
            for (int i = 1; i <= paragraphCount && i <= lines.Count; i++)
            {
                object paragraphRange = Late.Get(Late.Item(paragraphs, i), "Range");
                int start = Convert.ToInt32(Late.Get(paragraphRange, "Start"));
                int end = Convert.ToInt32(Late.Get(paragraphRange, "End"));
                if (Convert.ToString(Late.Get(paragraphRange, "Text")).EndsWith("\r")) end--;
                FormatLine(StoryRange(headerFooter, start, end), paragraphRange, lines[i - 1], background);
            }
            Late.Call(bookmarks, "Add", bookmark, range);
        }

        private void FormatLine(object text, object paragraph, StampLine line, int background)
        {
            object font = Late.Get(text, "Font");
            Late.TrySet(font, "Bold", true);
            if (line.Size > 0)
            {
                Late.TrySet(font, "Name", "Arial");
                Late.TrySet(font, "Size", line.Size);
            }
            if (!Late.TrySet(font, "Color", line.Color))
            {
                Late.TrySet(font, "ColorIndex", line.Color == 0 ? 1 : (line.Color == Late.Rgb(255, 255, 255) ? 8 : 6)); //Word 2000: black, white or red
            }
            //the IGI frame is drawn around the text itself, a character border; the IC banner has none
            object borders = Late.TryGet(font, "Borders");
            if (borders != null)
            {
                if (line.Frame > 0)
                {
                    Late.TrySet(borders, "OutsideLineStyle", 1); //wdLineStyleSingle
                    Late.TrySet(borders, "OutsideLineWidth", line.Frame >= 3 ? 24 : 18); //wdLineWidth300pt, wdLineWidth225pt: the nearest widths Word has
                    Late.TrySet(borders, "OutsideColor", line.Color);
                }
                else
                {
                    Late.TrySet(borders, "Enable", 0);
                }
            }
            object format = Late.Get(paragraph, "ParagraphFormat");
            Late.TrySet(format, "Alignment", 1); //wdAlignParagraphCenter
            Late.TrySet(Late.Get(format, "Shading"), "BackgroundPatternColor", background >= 0 ? background : ColorAutomatic);
        }

        //a new empty paragraph first in the header or last in the footer, as a range without its paragraph mark
        private object NewParagraph(object headerFooter, bool top)
        {
            object whole = Late.Get(headerFooter, "Range");
            string existing = Convert.ToString(Late.Get(whole, "Text")).Trim('\r', '\n', ' ');
            if (existing.Length > 0)
            {
                Late.Call(whole, top ? "InsertParagraphBefore" : "InsertParagraphAfter");
            }
            object paragraphs = Late.Get(Late.Get(headerFooter, "Range"), "Paragraphs");
            object paragraph = top ? Late.Get(paragraphs, "First") : Late.Get(paragraphs, "Last");
            object range = Late.Get(paragraph, "Range");
            int start = Convert.ToInt32(Late.Get(range, "Start"));
            int end = Convert.ToInt32(Late.Get(range, "End"));
            if (Convert.ToString(Late.Get(range, "Text")).EndsWith("\r"))
            {
                end--; //keep the paragraph mark out of the bookmark
            }
            object result = Late.Get(Late.Get(headerFooter, "Range"), "Duplicate");
            Late.Call(result, "SetRange", start, end);
            return result;
        }

        //removes a block written before, with its paragraph unless it is the only one of the header or footer
        private void RemoveBlock(object headerFooter, object range)
        {
            Late.Set(range, "Text", "");
            object paragraph = Late.Get(Late.Item(Late.Get(range, "Paragraphs"), 1), "Range");
            if (Late.Count(Late.Get(Late.Get(headerFooter, "Range"), "Paragraphs")) > 1 && Convert.ToString(Late.Get(paragraph, "Text")) == "\r")
            {
                Late.Call(paragraph, "Delete");
            }
            object borders = Late.TryGet(Late.Get(paragraph, "Font"), "Borders");
            if (borders != null) Late.TrySet(borders, "Enable", 0);
            Late.TrySet(Late.Get(Late.Get(paragraph, "ParagraphFormat"), "Shading"), "BackgroundPatternColor", ColorAutomatic);
        }

        //"page X / Y", with the reference before it at the TRÈS SECRET level, as page and page count fields
        private void WritePageNumbers(object footer, string bookmark, bool show, string reference, string format)
        {
            if (Convert.ToBoolean(Late.TryGet(footer, "LinkToPrevious") ?? false))
            {
                return;
            }
            object bookmarks = Late.Get(Document, "Bookmarks");
            bool exists = Convert.ToBoolean(Late.Call(bookmarks, "Exists", bookmark));
            if (!show)
            {
                if (exists) RemoveBlock(footer, Late.Get(Late.Item(bookmarks, bookmark), "Range"));
                return;
            }
            object range = exists ? Late.Get(Late.Item(bookmarks, bookmark), "Range") : NewParagraph(footer, false);
            int start = Convert.ToInt32(Late.Get(range, "Start"));
            string text = (string.IsNullOrEmpty(reference) ? "" : reference + " - ") + format.Replace("{P}", "\u0001").Replace("{N}", "\u0002");
            Late.Set(range, "Text", text);
            object borders = Late.TryGet(Late.Get(range, "Font"), "Borders");
            if (borders != null) Late.TrySet(borders, "Enable", 0);
            //the fields replace the placeholders, the last one first so the position of the other stays right
            InsertField(footer, start + text.IndexOf('\u0002'), FieldNumPages);
            InsertField(footer, start + text.IndexOf('\u0001'), FieldPage);
            object paragraph = Late.Get(Late.Item(Late.Get(StoryRange(footer, start, start), "Paragraphs"), 1), "Range");
            int end = Convert.ToInt32(Late.Get(paragraph, "End")) - 1;
            object block = Late.Get(Late.Get(footer, "Range"), "Duplicate");
            Late.Call(block, "SetRange", start, end);
            Late.TrySet(Late.Get(paragraph, "ParagraphFormat"), "Alignment", 1);
            Late.Call(bookmarks, "Add", bookmark, block);
        }

        //a range of a header or footer: Document.Range would point at the same positions in the body of the document
        private static object StoryRange(object storyItem, int start, int end)
        {
            object range = Late.Get(Late.Get(storyItem, "Range"), "Duplicate");
            Late.Call(range, "SetRange", start, end);
            return range;
        }

        private void InsertField(object storyItem, int position, int type)
        {
            object placeholder = Late.Get(Late.Get(storyItem, "Range"), "Duplicate");
            Late.Call(placeholder, "SetRange", position, position + 1);
            Late.Call(Late.Get(Document, "Fields"), "Add", placeholder, type);
        }

        protected override IEnumerable<string> Paragraphs()
        {
            //one string for the whole body is much faster than a COM call per paragraph; table cells end with "\r\a"
            string text = Convert.ToString(Late.Get(Late.Get(Document, "Content"), "Text"));
            return text.Split('\r');
        }

        public override void InsertPortionMark(string mark)
        {
            object selection = Late.Get(Application, "Selection");
            object paragraph = Late.Get(Late.Item(Late.Get(selection, "Paragraphs"), 1), "Range");
            int start = Convert.ToInt32(Late.Get(paragraph, "Start"));
            string text = Convert.ToString(Late.Get(paragraph, "Text"));
            int leading = text.Length - text.TrimStart(' ', '\t').Length;
            int existing = ExistingMarkLength(text.Substring(leading));
            object range = Late.Call(Document, "Range", start + leading, start + leading + existing);
            Late.Set(range, "Text", mark + " ");
        }

        public override void WriteBlock(string bookmark, string[] lines)
        {
            object bookmarks = Late.Get(Document, "Bookmarks");
            object range;
            if (Convert.ToBoolean(Late.Call(bookmarks, "Exists", bookmark)))
            {
                range = Late.Get(Late.Item(bookmarks, bookmark), "Range");
            }
            else
            {
                //a new block is its own paragraph, so the portion mark of the paragraph at the cursor stays at its start
                range = Late.Get(Late.Get(Application, "Selection"), "Range");
                Late.Call(range, "Collapse", 1); //wdCollapseStart
                Late.Call(range, "InsertParagraphBefore");
                Late.Call(range, "Collapse", 1);
            }
            int start = Convert.ToInt32(Late.Get(range, "Start"));
            string text = string.Join("\v", lines); //line breaks keep the block in one paragraph
            int field = text.IndexOf("{NUMPAGES}");
            if (field >= 0)
            {
                text = text.Replace("{NUMPAGES}", "\u0002");
            }
            Late.Set(range, "Text", text);
            if (field >= 0)
            {
                object placeholder = Late.Call(Document, "Range", start + field, start + field + 1);
                Late.Call(Late.Get(Document, "Fields"), "Add", placeholder, FieldNumPages);
                range = Late.Call(Document, "Range", start, Convert.ToInt32(Late.Get(range, "End")));
            }
            Late.Call(bookmarks, "Add", bookmark, range);
        }

        public override void WriteAuthorityBox(string[] lines)
        {
            //the box written before, and the block of earlier versions written at the cursor
            object shapes = Late.Get(Document, "Shapes");
            for (int i = Late.Count(shapes); i >= 1; i--)
            {
                object shape = Late.Item(shapes, i);
                if (Convert.ToString(Late.TryGet(shape, "Name")) == AuthorityBookmark)
                {
                    Late.Call(shape, "Delete");
                }
            }
            object bookmarks = Late.Get(Document, "Bookmarks");
            if (Convert.ToBoolean(Late.Call(bookmarks, "Exists", AuthorityBookmark)))
            {
                object old = Late.Get(Late.Item(bookmarks, AuthorityBookmark), "Range");
                Late.Call(Late.Get(Late.Item(Late.Get(old, "Paragraphs"), 1), "Range"), "Delete");
            }
            //anchored to the first paragraph, so it stays on the first page
            object anchor = Late.Get(Late.Item(Late.Get(Document, "Paragraphs"), 1), "Range");
            float height = lines.Length * 10f + 10f;
            object box = Late.Call(shapes, "AddTextbox", 1, 0, 0, 216, height, anchor); //msoTextOrientationHorizontal, 7.6 cm wide
            Late.Set(box, "Name", AuthorityBookmark);
            Late.TrySet(box, "LockAnchor", true);
            Late.TrySet(Late.Get(box, "WrapFormat"), "Type", 0); //wdWrapSquare: the text flows around the box
            object line = Late.Get(box, "Line");
            Late.TrySet(line, "Weight", 0.75f);
            Late.TrySet(Late.Get(line, "ForeColor"), "RGB", 0);
            object frame = Late.Get(box, "TextFrame");
            foreach (string margin in new string[] { "MarginLeft", "MarginRight", "MarginTop", "MarginBottom" })
            {
                Late.TrySet(frame, margin, 4f);
            }
            object text = Late.Get(frame, "TextRange");
            Late.Set(text, "Text", string.Join("\r", lines).Replace("{NUMPAGES}", ""));
            object font = Late.Get(text, "Font");
            Late.Set(font, "Name", "Arial");
            Late.Set(font, "Size", 8);
            Late.TrySet(font, "Bold", 0);
            Late.TrySet(Late.Get(text, "ParagraphFormat"), "SpaceAfter", 0);
            Late.TrySet(Late.Get(text, "ParagraphFormat"), "SpaceBefore", 0);
            Late.TrySet(frame, "AutoSize", -1); //Word 2007 and later: the box fits its lines
            //in the corner of the margins, from the page edges: the wdShapeRight and wdShapeBottom positions follow the anchor
            //paragraph in some Word versions, the box then sits in the middle of the page
            object setup = Late.Get(Late.Item(Late.Get(Document, "Sections"), 1), "PageSetup");
            float pageWidth = Convert.ToSingle(Late.Get(setup, "PageWidth"));
            float pageHeight = Convert.ToSingle(Late.Get(setup, "PageHeight"));
            float right = Convert.ToSingle(Late.Get(setup, "RightMargin"));
            float bottom = Convert.ToSingle(Late.Get(setup, "BottomMargin"));
            Late.TrySet(box, "RelativeHorizontalPosition", 1); //wdRelativeHorizontalPositionPage
            Late.TrySet(box, "RelativeVerticalPosition", 1); //wdRelativeVerticalPositionPage
            Late.Set(box, "Left", pageWidth - right - Convert.ToSingle(Late.Get(box, "Width")));
            Late.Set(box, "Top", pageHeight - bottom - Convert.ToSingle(Late.Get(box, "Height")));
        }
    }

    public class ExcelHost : DocumentHost
    {
        protected override void WriteStamps(StampSet stamps)
        {
            //Excel headers and footers only take text, no frame; colors exist from Excel 2007 (&K), an escaped "&" is "&&"
            string header = Lines(stamps.Top, stamps.Background);
            string footer = Lines(stamps.Bottom, stamps.Background);
            if (stamps.PageNumbers)
            {
                footer += (footer.Length > 0 ? "\n" : "") + "&\"Arial,Regular\"&10" + Color(0) + (string.IsNullOrEmpty(stamps.Reference) ? "" : Escape(stamps.Reference) + " - ") + stamps.PageFormat.Replace("{P}", "&P").Replace("{N}", "&N");
            }
            object sheets = Late.Get(Document, "Worksheets");
            int count = Late.Count(sheets);
            for (int i = 1; i <= count; i++)
            {
                object setup = Late.Get(Late.Item(sheets, i), "PageSetup");
                Late.Set(setup, "CenterHeader", header);
                Late.Set(setup, "CenterFooter", footer);
            }
        }

        private string Lines(List<StampLine> lines, int background)
        {
            List<string> parts = new List<string>();
            foreach (StampLine line in lines)
            {
                //the IC banner shows the color of its background, IGI stamps their own color
                int color = background >= 0 ? background : line.Color;
                parts.Add("&\"Arial,Bold\"&" + (line.Size > 0 ? line.Size : 12) + Color(color) + Escape(line.Text));
            }
            return string.Join("\n", parts.ToArray());
        }

        private string Color(int bgr)
        {
            if (Version < 12)
            {
                return "";
            }
            int red = bgr & 0xFF, green = (bgr >> 8) & 0xFF, blue = (bgr >> 16) & 0xFF;
            return "&K" + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
        }

        private static string Escape(string text)
        {
            return text.Replace("&", "&&");
        }
    }

    public class PowerPointHost : DocumentHost
    {
        private const string ShapePrefix = "SystemBanner";
        private const float LineHeight = 20;

        public override bool SupportsPortions
        {
            get { return true; }
        }

        protected override void WriteStamps(StampSet stamps)
        {
            List<object> masters = new List<object>();
            masters.Add(Late.Get(Document, "SlideMaster"));
            if (Convert.ToInt32(Late.TryGet(Document, "HasTitleMaster") ?? 0) != 0)
            {
                masters.Add(Late.Get(Document, "TitleMaster")); //PowerPoint 2000 to 2003 title slides use their own master
            }
            object setup = Late.Get(Document, "PageSetup");
            float width = Convert.ToSingle(Late.Get(setup, "SlideWidth"));
            float height = Convert.ToSingle(Late.Get(setup, "SlideHeight"));
            foreach (object master in masters)
            {
                object shapes = Late.Get(master, "Shapes");
                RemoveShapes(shapes);
                float top = 0;
                for (int i = 0; i < stamps.Top.Count; i++)
                {
                    top += AddLine(shapes, ShapePrefix + "Top" + (i == 0 ? "" : (i + 1).ToString()), top, width, stamps.Top[i], stamps.Background);
                }
                float bottom = height;
                for (int i = stamps.Bottom.Count - 1; i >= 0; i--)
                {
                    bottom -= Height(stamps.Bottom[i]);
                    AddLine(shapes, ShapePrefix + "Bottom" + (i == 0 ? "" : (i + 1).ToString()), bottom, width, stamps.Bottom[i], stamps.Background);
                }
            }
        }

        private static float Height(StampLine line)
        {
            return line.Size > 12 ? line.Size + 12 : LineHeight;
        }

        private static void RemoveShapes(object shapes)
        {
            for (int i = Late.Count(shapes); i >= 1; i--)
            {
                object shape = Late.Item(shapes, i);
                if (Convert.ToString(Late.Get(shape, "Name")).StartsWith(ShapePrefix))
                {
                    Late.Call(shape, "Delete");
                }
            }
        }

        //an IC banner line across the slide, or an IGI stamp: framed, sized to its text and centered
        private static float AddLine(object shapes, string name, float top, float width, StampLine line, int background)
        {
            float height = Height(line);
            object shape = Late.Call(shapes, "AddTextbox", 1, 0f, top, width, height); //msoTextOrientationHorizontal
            Late.Set(shape, "Name", name);
            object fill = Late.Get(shape, "Fill");
            Late.Set(fill, "Visible", -1);
            Late.Set(Late.Get(fill, "ForeColor"), "RGB", background >= 0 ? background : Late.Rgb(255, 255, 255));
            object frame = Late.Get(shape, "TextFrame");
            object range = Late.Get(frame, "TextRange");
            Late.Set(range, "Text", line.Text);
            object font = Late.Get(range, "Font");
            Late.Set(font, "Bold", -1);
            Late.Set(font, "Size", line.Size > 0 ? line.Size : 12);
            Late.TrySet(font, "Name", "Arial");
            Late.Set(Late.Get(font, "Color"), "RGB", line.Color);
            Late.Set(Late.Get(range, "ParagraphFormat"), "Alignment", 2); //ppAlignCenter
            if (line.Frame > 0)
            {
                Late.TrySet(frame, "WordWrap", 0);
                Late.TrySet(frame, "AutoSize", 1); //ppAutoSizeShapeToFitText
                object border = Late.Get(shape, "Line");
                Late.Set(border, "Visible", -1);
                Late.Set(Late.Get(border, "ForeColor"), "RGB", line.Color);
                Late.Set(border, "Weight", line.Frame);
                float shapeWidth = Convert.ToSingle(Late.Get(shape, "Width"));
                Late.Set(shape, "Left", (width - shapeWidth) / 2);
            }
            else
            {
                Late.TrySet(frame, "WordWrap", -1);
                Late.TrySet(frame, "AutoSize", 0);
            }
            return height;
        }

        protected override IEnumerable<string> Paragraphs()
        {
            List<string> paragraphs = new List<string>();
            object slides = Late.Get(Document, "Slides");
            int count = Late.Count(slides);
            for (int i = 1; i <= count; i++)
            {
                object shapes = Late.Get(Late.Item(slides, i), "Shapes");
                int shapeCount = Late.Count(shapes);
                for (int j = 1; j <= shapeCount; j++)
                {
                    object shape = Late.Item(shapes, j);
                    if (Convert.ToInt32(Late.TryGet(shape, "HasTextFrame") ?? 0) == 0)
                    {
                        continue;
                    }
                    object frame = Late.Get(shape, "TextFrame");
                    if (Convert.ToInt32(Late.TryGet(frame, "HasText") ?? 0) == 0)
                    {
                        continue;
                    }
                    paragraphs.AddRange(Convert.ToString(Late.Get(Late.Get(frame, "TextRange"), "Text")).Split('\r'));
                }
            }
            return paragraphs;
        }

        public override void InsertPortionMark(string mark)
        {
            object selection = Late.Get(Late.Get(Application, "ActiveWindow"), "Selection");
            if (Convert.ToInt32(Late.Get(selection, "Type")) != 3) //ppSelectionText
            {
                throw new InvalidOperationException(Text.T("Put the cursor in the text that the portion mark applies to.", "Placer le curseur dans le texte auquel la marque s'applique.", "Setzen Sie den Cursor in den Text, für den die Teilkennzeichnung gilt."));
            }
            object paragraph = Late.Call(Late.Get(selection, "TextRange"), "Paragraphs", 1, 1);
            string text = Convert.ToString(Late.Get(paragraph, "Text"));
            int existing = ExistingMarkLength(text);
            if (existing > 0)
            {
                Late.Call(Late.Call(paragraph, "Characters", 1, existing), "Delete");
            }
            Late.Call(paragraph, "InsertBefore", mark + " ");
        }
    }
}
