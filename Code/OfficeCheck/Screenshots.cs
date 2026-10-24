//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice.Check
{
    //OfficeCheck.exe screenshots [folder]: the pictures of the README, made again at each build so they follow the add-in:
    //each dialog filled with an example in English, French and German, and a marked Word page and PowerPoint slide of each system.
    //Example content only: no real document goes through it.
    public static class Screenshots
    {
        private static string folder;
        private static int written;

        public static int Run(string target)
        {
            folder = Path.GetFullPath(target);
            Directory.CreateDirectory(folder);
            Language saved = Text.Current;
            try
            {
                foreach (Language language in new Language[] { Language.English, Language.French, Language.German })
                {
                    Text.Current = language;
                    Dialogs(language == Language.English ? "en" : (language == Language.French ? "fr" : "de"));
                }
            }
            finally
            {
                Text.Current = saved;
            }
            Text.Current = Language.English;
            Office("Word.Application", WordPages);
            Office("PowerPoint.Application", PowerPointSlides);
            Console.WriteLine(written + " pictures written to " + folder);
            return 0;
        }

        private delegate Form NewDialog();

        private static void Dialogs(string language)
        {
            SystemSettings settings = new SystemSettings();
            Dialog("Dialog_System_" + language, delegate { return new SystemForm(DocumentHost.IcSystem, new string[] { DocumentHost.FrenchSystem, DocumentHost.GermanSystem, DocumentHost.BritishSystem, DocumentHost.EuSystem, DocumentHost.IcSystem }); });
            Dialog("Dialog_IC_Classify_" + language, delegate { return new MarkingForm(false, MarkingParser.Parse("TOP SECRET//SI//NOFORN").Marking, null, settings); });
            Dialog("Dialog_IC_PortionMark_" + language, delegate { return new MarkingForm(true, MarkingParser.Parse("S//REL TO USA, GBR").Marking, MarkingParser.Parse("TOP SECRET//SI//NOFORN").Marking, settings); });
            Dialog("Dialog_IC_AuthorityBlock_" + language, delegate { return new AuthorityForm(Authority(), MarkingParser.Parse("TOP SECRET//SI//NOFORN").Marking); });
            Dialog("Dialog_France_Classify_" + language, delegate
            {
                FrenchMarking m = FrenchParser.ParsePortion("[S SF]");
                return new FrenchForm(false, m, false, true, settings);
            });
            Dialog("Dialog_France_Identification_" + language, delegate { return new IdentificationForm(FrenchIdentificationExample(), FrenchParser.ParsePortion("[TS]")); });
            Dialog("Dialog_France_Downgrade_" + language, delegate { return new ChangeForm(FrenchParser.ParsePortion("[TS]"), settings); });
            Dialog("Dialog_Germany_Classify_" + language, delegate
            {
                GermanMarking m = GermanParser.ParsePortion("[Geh.]");
                m.Vermerke.Add(GermanVermerke.Sperrvermerke[0]);
                return new GermanForm(false, m, settings);
            });
            Dialog("Dialog_Germany_FirstPage_" + language, delegate { return new GermanIdentificationForm(GermanIdentificationExample(), GermanParser.ParsePortion("[Geh.]")); });
            Dialog("Dialog_Germany_Downgrade_" + language, delegate { return new GermanChangeForm(GermanParser.ParsePortion("[Geh.]")); });
            Dialog("Dialog_UK_Classify_" + language, delegate { return new BritishForm(UkExample(), settings); });
            Dialog("Dialog_EU_Classify_" + language, delegate { return new EuForm(false, EuExample(), settings); });
        }

        private static void Dialog(string name, NewDialog create)
        {
            try
            {
                using (Form form = create())
                {
                    form.StartPosition = FormStartPosition.Manual;
                    form.Location = new Point(-2000, -2000); //drawn off screen
                    form.Show();
                    Application.DoEvents();
                    using (Bitmap image = new Bitmap(form.Width, form.Height))
                    {
                        form.DrawToBitmap(image, new Rectangle(0, 0, form.Width, form.Height));
                        Save(image, name);
                    }
                    form.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name + ": " + (ex.InnerException ?? ex).Message);
            }
        }

        private static void Save(Image image, string name)
        {
            image.Save(Path.Combine(folder, name + ".png"), ImageFormat.Png);
            written++;
            Console.WriteLine("ok   " + name + ".png");
        }

        private delegate void Scenario(object application, string version);

        private static void Office(string progId, Scenario scenario)
        {
            Type type = Type.GetTypeFromProgID(progId);
            if (type == null)
            {
                Console.WriteLine("skip " + progId + ": not installed");
                return;
            }
            object application = OfficeCheck.Start(type);
            try
            {
                string version = Convert.ToString(Late.Get(application, "Version"));
                int dot = version.IndexOf('.');
                scenario(application, Year(dot > 0 ? version.Substring(0, dot) : version));
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + progId + ": " + (ex.InnerException ?? ex).Message);
            }
            finally
            {
                try { Late.Call(application, "Quit"); } catch (Exception) { }
                Marshal.ReleaseComObject(application);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        //"Word2016": the name of the Office version in the file names
        private static string Year(string major)
        {
            switch (major)
            {
                case "9": return "2000";
                case "10": return "XP";
                case "11": return "2003";
                case "12": return "2007";
                case "14": return "2010";
                case "15": return "2013";
                case "16": return "2016";
                default: return major;
            }
        }

        private delegate void Marker(object document, DocumentHost host);

        private static void WordPages(object word, string version)
        {
            Late.TrySet(word, "DisplayAlerts", 0);
            string prefix = "Word" + version + "_";
            WordPage(word, prefix + "IC_Secret",
                "(U) Example document: the banner lines and portion marks of the IC Markings System.\r(S//REL TO USA, GBR) A paragraph releasable to the United Kingdom.\r(S//NF) A SECRET paragraph, not releasable to foreign nationals.\r(U) The banner line is rolled up from the portion marks; the classification authority block is at the bottom right of the first page.",
                delegate(object document, DocumentHost host)
                {
                    RollUpResult result = RollUp.Compute(host.ScanPortions().Marks, null);
                    host.Mark(result.Banner);
                    host.WriteAuthorityBox(Authority().Lines());
                });
            WordPage(word, prefix + "France_SecretSpecialFrance",
                "[NP] Objet : exemple de document classifié selon l'IGI 1300.\r[S SF] Paragraphe classifié Secret, Spécial France.\r[DR] Paragraphe Diffusion Restreinte.",
                delegate(object document, DocumentHost host)
                {
                    FrenchIdentification id = FrenchIdentificationExample();
                    FrenchMarking marking = FrenchRollUp.Compute(host.ScanPortions().FrenchMarks).Document;
                    host.Mark(marking, id, false);
                    Late.Call(Late.Get(word, "Selection"), "HomeKey", 6); //wdStory
                    host.WriteBlock(DocumentHost.IdentificationBookmark, id.Lines(marking));
                });
            WordPage(word, prefix + "Germany_Geheim",
                "[offen] Betreff: Beispiel eines Dokuments nach der VS-Anweisung.\r[Geh.] Ein GEHEIM eingestufter Teil.\r[VS-NfD] Ein Teil VS-NUR FÜR DEN DIENSTGEBRAUCH.",
                delegate(object document, DocumentHost host)
                {
                    GermanMarking marking = GermanRollUp.Compute(host.ScanPortions().GermanMarks);
                    marking.Vermerke.Add(GermanVermerke.Sperrvermerke[0]);
                    host.Mark(marking);
                    Late.Call(Late.Get(word, "Selection"), "HomeKey", 6);
                    host.WriteBlock(DocumentHost.IdentificationBookmark, GermanIdentificationExample().Lines(marking));
                });
            WordPage(word, prefix + "UK_OfficialSensitive",
                "Example document marked under the Government Security Classifications Policy.\rThe GSCP has no portion marking: the whole marking is at the top and bottom of every page, in black.",
                delegate(object document, DocumentHost host) { host.Mark(UkExample()); });
            WordPage(word, prefix + "EU_Secret",
                "[R-UE/EU-R] Subject: example of an EU classified document.\r[S-UE/EU-S] A paragraph classified SECRET UE/EU SECRET.\r[R-UE/EU-R] A paragraph classified RESTREINT UE/EU RESTRICTED.",
                delegate(object document, DocumentHost host)
                {
                    EuMarking marking = EuRollUp.Compute(host.ScanPortions().EuMarks);
                    EuMarking example = EuExample();
                    marking.Reference = example.Reference;
                    marking.Copy = example.Copy;
                    marking.Additional = example.Additional;
                    host.Mark(marking);
                });
        }

        private static void WordPage(object word, string name, string text, Marker mark)
        {
            object document = Late.Call(Late.Get(word, "Documents"), "Add");
            try
            {
                Late.Set(Late.Get(document, "Content"), "Text", text);
                DocumentHost host = DocumentHost.For(word, document);
                mark(document, host);
                object window = Late.Get(document, "ActiveWindow");
                Late.Set(Late.Get(window, "View"), "Type", 3); //wdPrintView
                object page = Late.Item(Late.Get(Late.Item(Late.Get(window, "Panes"), 1), "Pages"), 1);
                byte[] bits = (byte[])Late.Get(page, "EnhMetaFileBits");
                //drawn at three times the size, then reduced: the text of the page keeps its spacing
                using (MemoryStream stream = new MemoryStream(bits))
                using (Metafile metafile = new Metafile(stream))
                using (Bitmap large = new Bitmap(2550, 3300))
                using (Bitmap image = new Bitmap(850, 1100))
                using (Graphics g = Graphics.FromImage(image))
                {
                    using (Graphics l = Graphics.FromImage(large))
                    {
                        l.Clear(Color.White);
                        l.DrawImage(metafile, 0, 0, 2550, 3300);
                    }
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(large, 0, 0, 850, 1100);
                    using (Pen edge = new Pen(Color.FromArgb(180, 180, 180)))
                    {
                        g.DrawRectangle(edge, 0, 0, 849, 1099);
                    }
                    Save(image, name);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL " + name + ": " + (ex.InnerException ?? ex).Message);
            }
            finally
            {
                Late.Call(document, "Close", 0);
            }
        }

        private static void PowerPointSlides(object powerPoint, string version)
        {
            object presentation = Late.Call(Late.Get(powerPoint, "Presentations"), "Add", 0); //no window
            try
            {
                object slide = Late.Call(Late.Get(presentation, "Slides"), "Add", 1, 2); //ppLayoutText
                object shapes = Late.Get(slide, "Shapes");
                Late.Set(Late.Get(Late.Get(Late.Item(shapes, 1), "TextFrame"), "TextRange"), "Text", "(S) Example briefing");
                Late.Set(Late.Get(Late.Get(Late.Item(shapes, 2), "TextFrame"), "TextRange"), "Text", "(U) Banner lines on the slide master, at the top and bottom of every slide.\r(S//NF) A SECRET bullet, not releasable to foreign nationals.\r(U) The banner is rolled up from the portion marks.");
                DocumentHost host = DocumentHost.For(powerPoint, presentation);
                host.Mark(RollUp.Compute(host.ScanPortions().Marks, null).Banner);
                string name = "PowerPoint" + version + "_IC_Secret";
                Late.Call(slide, "Export", Path.Combine(folder, name + ".png"), "PNG", 1280, 720);
                written++;
                Console.WriteLine("ok   " + name + ".png");
            }
            finally
            {
                Late.Call(presentation, "Close");
            }
        }

        private static AuthorityBlock Authority()
        {
            AuthorityBlock block = new AuthorityBlock();
            block.ClassifiedBy = "J. Doe, Senior Analyst";
            block.DerivedFrom = "Example Classification Guide, 20260101";
            block.DeclassifyOn = "20510101";
            return block;
        }

        private static FrenchIdentification FrenchIdentificationExample()
        {
            FrenchIdentification id = new FrenchIdentification();
            id.Lieu = "Paris";
            id.AutoriteEmettrice = "Ministère X";
            id.Auteur = "Bureau Y";
            id.NumeroEnregistrement = "2026-042";
            id.Echeance = DateTime.Today.AddYears(10);
            id.Exemplaire = 1;
            id.NombreExemplaires = 3;
            id.Reference = "TS-2026-042";
            return id;
        }

        private static GermanIdentification GermanIdentificationExample()
        {
            GermanIdentification id = new GermanIdentification();
            id.Herausgeber = "Bundesministerium X";
            id.Kennzeichen = "Tgb.-Nr. 42/26";
            id.Ausfertigung = 2;
            id.Empfaenger = "Referat Y";
            id.EinstufungEndet = DateTime.Today.Year + 10;
            return id;
        }

        private static UkMarking UkExample()
        {
            UkMarking m = new UkMarking();
            m.Sensitive = true;
            m.HandlingInstruction = UkRegister.RecipientsOnly;
            m.Descriptors.Add("PERSONAL DATA");
            return m;
        }

        private static EuMarking EuExample()
        {
            EuMarking m = new EuMarking();
            m.Level = EuLevel.Secret;
            m.Reference = "EX(2026) 42";
            m.Copy = 2;
            m.Additional.Add("RELEASABLE TO NATO");
            return m;
        }
    }
}
