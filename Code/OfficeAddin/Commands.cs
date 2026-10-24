//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //the buttons of the add-in, the same in the Office 2007-2010 ribbon and the Office 2000-2003 toolbar
    public class Commands
    {
        public const string Title = "SystemBanner";
        private readonly object application;
        private Marking lastPortion;
        private FrenchMarking lastFrenchPortion;
        private GermanMarking lastGermanPortion;

        public Commands(object application)
        {
            this.application = application;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, IntPtr processId);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
        private delegate bool WindowCallback(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(uint thread, WindowCallback callback, IntPtr parameter);

        //dialogs are owned by the main Office window, so they open in front of it. Office 2000 has no window handle property.
        //The foreground window alone is not enough: in Office 2007 it can be a ribbon or screen tip window, which hides
        //or closes the dialog with it. The main window is the top level window of Office's thread with Office's frame class.
        private class OfficeWindow : IWin32Window
        {
            private readonly IntPtr handle = Find();
            public IntPtr Handle
            {
                get { return handle; }
            }

            private static bool IsFrame(IntPtr window)
            {
                StringBuilder name = new StringBuilder(64);
                GetClassName(window, name, name.Capacity);
                string frame = name.ToString();
                //Word: OpusApp; Excel: XLMAIN; PowerPoint: PP9FrameClass to PP12FrameClass, then PPTFrameClass
                return frame == "OpusApp" || frame == "XLMAIN" || (frame.StartsWith("PP") && frame.EndsWith("FrameClass"));
            }

            private static IntPtr Find()
            {
                uint thread = GetCurrentThreadId();
                IntPtr foreground = GetAncestor(GetForegroundWindow(), 2); //GA_ROOT
                if (foreground != IntPtr.Zero && GetWindowThreadProcessId(foreground, IntPtr.Zero) == thread && IsFrame(foreground))
                {
                    return foreground;
                }
                IntPtr found = IntPtr.Zero;
                EnumThreadWindows(thread, delegate(IntPtr window, IntPtr parameter)
                {
                    if (IsWindowVisible(window) && IsFrame(window))
                    {
                        found = window;
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);
                return found != IntPtr.Zero ? found : foreground;
            }
        }

        private static void Show(string message, MessageBoxIcon icon)
        {
            MessageBox.Show(new OfficeWindow(), message, Title, MessageBoxButtons.OK, icon);
        }

        private static bool Ask(string message)
        {
            return MessageBox.Show(new OfficeWindow(), message, Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        private DocumentHost Host()
        {
            DocumentHost host = DocumentHost.Active(application);
            if (host == null)
            {
                Show(Text.T("Open a document first.", "Ouvrir d'abord un document.", "Öffnen Sie zuerst ein Dokument."), MessageBoxIcon.Information);
            }
            return host;
        }

        //English or French, from the OfficeLanguage setting or the language of the Office user interface (msoLanguageIDUI = 2)
        public void ChooseLanguage()
        {
            int language = 0;
            object settings = Late.TryGet(application, "LanguageSettings");
            if (settings != null)
            {
                object id = Late.TryGet(settings, "LanguageID", 2);
                if (id != null) int.TryParse(Convert.ToString(id), out language);
            }
            Text.Choose(language);
        }

        //every button runs here, so an Office error shows as a message instead of unloading the add-in
        public void Run(string command)
        {
            try
            {
                Log.Write("command " + command + ", dialogs owned by window " + new OfficeWindow().Handle);
                ChooseLanguage();
                switch (command)
                {
                    case "Classify": Classify(); break;
                    case "PortionMark": PortionMark(); break;
                    case "RollUp": RollUpDocument(true); break;
                    case "AuthorityBlock": WriteAuthorityBlock(); break;
                    case "Downgrade": Downgrade(); break;
                }
            }
            catch (Exception ex)
            {
                Exception inner = ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                Log.Error(command, ex);
                Show(Text.T("The command failed: ", "La commande a échoué : ", "Der Befehl ist fehlgeschlagen: ") + inner.Message + Environment.NewLine + Environment.NewLine + Text.T("Details are in ", "Le détail est dans ", "Einzelheiten stehen in ") + Log.FilePath, MessageBoxIcon.Error);
            }
        }

        //the systems this computer allows: none set (no Simple value) allows all four, a national value its own system only
        private static List<string> AvailableSystems(SystemSettings settings)
        {
            List<string> systems = new List<string>();
            if (settings.FrenchAvailable) systems.Add(DocumentHost.FrenchSystem);
            if (GermanSettings.Available(settings)) systems.Add(DocumentHost.GermanSystem);
            if (UkSettings.Available(settings)) systems.Add(DocumentHost.BritishSystem);
            if (settings.IcAvailable) systems.Add(DocumentHost.IcSystem);
            return systems;
        }

        //the system of this document; null when the classifier cancels or the computer is set to NATO or UN
        private string ChooseSystem(DocumentHost host, SystemSettings settings, bool ask)
        {
            List<string> systems = AvailableSystems(settings);
            string stored = host.StoredSystem();
            if (systems.Count == 0) return null;
            if (systems.Count == 1) return systems[0];
            if (!ask && stored != null && systems.Contains(stored)) return stored;
            SystemForm form = new SystemForm(stored != null && systems.Contains(stored) ? stored : systems[0], systems.ToArray());
            return form.ShowDialog(new OfficeWindow()) == DialogResult.OK ? form.Result : null;
        }

        public void Classify()
        {
            DocumentHost host = Host();
            if (host == null) return;
            SystemSettings settings = SystemSettings.Read();
            MarkingFormatter.Abbreviate = !settings.SpellOut;
            if (AvailableSystems(settings).Count == 0)
            {
                //NATO and UN computers: the stock levels of SystemBanner
                NationalForm national = new NationalForm(settings, host.StoredClassification());
                if (national.ShowDialog(new OfficeWindow()) == DialogResult.OK)
                {
                    host.Mark(national.Result);
                }
                return;
            }
            switch (ChooseSystem(host, settings, true))
            {
                case DocumentHost.FrenchSystem: ClassifyFrench(host, settings); break;
                case DocumentHost.GermanSystem: ClassifyGerman(host, settings); break;
                case DocumentHost.BritishSystem: ClassifyBritish(host, settings); break;
                case DocumentHost.IcSystem: ClassifyIc(host, settings); break;
            }
        }

        //a document with marks takes its level from them, the classifier changes the marks instead
        private bool OfferRollUp(DocumentHost host, int marks)
        {
            if (!host.SupportsPortions || marks == 0)
            {
                return false;
            }
            if (Ask(Text.T("This document has portion marks: its classification is rolled up from them.\n\nRoll up the portion marks now?",
                "Ce document comporte des marques de paragraphe : sa classification en est déduite.\n\nCalculer la classification maintenant ?",
                "Dieses Dokument enthält Teilkennzeichnungen: Seine Einstufung wird daraus abgeleitet.\n\nEinstufung jetzt aus den Teilkennzeichnungen ableiten?")))
            {
                RollUpDocument(true);
            }
            return true;
        }

        private void ClassifyIc(DocumentHost host, SystemSettings settings)
        {
            if (OfferRollUp(host, host.SupportsPortions ? host.ScanPortions().Marks.Count : 0)) return;
            MarkingForm form = new MarkingForm(false, host.StoredBanner(), null, settings);
            if (form.ShowDialog(new OfficeWindow()) == DialogResult.OK && form.Result != null)
            {
                host.Mark(form.Result);
                SuggestAuthorityBlock(host, form.Result);
            }
        }

        private void ClassifyFrench(DocumentHost host, SystemSettings settings)
        {
            if (OfferRollUp(host, host.SupportsPortions ? host.ScanPortions().FrenchMarks.Count : 0)) return;
            FrenchForm form = new FrenchForm(false, host.StoredFrenchMarking(), host.StoredBound(), host.SupportsCover, settings);
            if (form.ShowDialog(new OfficeWindow()) == DialogResult.OK && form.Result != null)
            {
                MarkFrench(host, form.Result, form.Bound);
            }
        }

        private void ClassifyGerman(DocumentHost host, SystemSettings settings)
        {
            if (OfferRollUp(host, host.SupportsPortions ? host.ScanPortions().GermanMarks.Count : 0)) return;
            GermanForm form = new GermanForm(false, host.StoredGermanMarking(), settings);
            if (form.ShowDialog(new OfficeWindow()) == DialogResult.OK && form.Result != null)
            {
                MarkGerman(host, form.Result, true);
            }
        }

        private void ClassifyBritish(DocumentHost host, SystemSettings settings)
        {
            BritishForm form = new BritishForm(host.StoredUkMarking(), settings);
            if (form.ShowDialog(new OfficeWindow()) == DialogResult.OK && form.Result != null)
            {
                host.Mark(form.Result);
            }
        }

        //stamps the document and asks for its first page identification when it is incomplete
        private void MarkFrench(DocumentHost host, FrenchMarking marking, bool bound)
        {
            FrenchIdentification id = host.StoredFrenchIdentification();
            host.StoreDeclassification(null);
            host.Mark(marking, id, bound);
            if (host is WordHost && id.Validate(marking).Count > 0
                && Ask(Text.T("A " + marking.ToString() + " document is identified on its first page (issuing authority, date, registration number" + (marking.IsClassified ? ", classification deadline" : "") + ").\n\nWrite the identification now?",
                    "Un document " + marking.ToString() + " est identifié dès sa première page (autorité émettrice, date, numéro d'enregistrement" + (marking.IsClassified ? ", échéance de classification" : "") + ").\n\nÉcrire l'identification maintenant ?",
                    "Ein Dokument " + marking.ToString() + " wird auf seiner ersten Seite gekennzeichnet (ausstellende Stelle, Datum, Registriernummer" + (marking.IsClassified ? ", Ablauf der Einstufung" : "") + ").\n\nKennzeichnung jetzt schreiben?")))
            {
                WriteIdentification(host, marking);
            }
        }

        private void MarkGerman(DocumentHost host, GermanMarking marking, bool askIdentification)
        {
            host.Mark(marking);
            if (askIdentification && host is WordHost && marking.IsClassified && host.StoredGermanIdentification().Validate(marking).Count > 0
                && Ask(Text.T("A " + GermanFormatter.Stamp(marking.Level) + " document carries its issuer, date and page count" + (marking.AmtlichGeheimGehalten ? ", end of classification, reference and copy number" : "") + " (VSA § 20 Abs. 2).\n\nWrite them now?",
                    "Un document " + GermanFormatter.Stamp(marking.Level) + " porte son émetteur, sa date et son nombre de pages" + (marking.AmtlichGeheimGehalten ? ", la fin de classification, sa référence et son numéro d'exemplaire" : "") + " (VSA § 20 al. 2).\n\nLes écrire maintenant ?",
                    "Ein Dokument " + GermanFormatter.Stamp(marking.Level) + " trägt Herausgeber, Datum und Seitenzahl" + (marking.AmtlichGeheimGehalten ? ", Ende der Einstufung, Geschäftszeichen und Ausfertigungsnummer" : "") + " (VSA § 20 Abs. 2).\n\nJetzt schreiben?")))
            {
                WriteGermanIdentification(host, marking);
            }
        }

        public void PortionMark()
        {
            DocumentHost host = Host();
            if (host == null) return;
            if (!host.SupportsPortions)
            {
                Show(Text.T("Excel workbooks are marked as a whole: use Classify.", "Les classeurs Excel sont marqués en entier : utiliser Classify.", "Excel-Arbeitsmappen werden als Ganzes gekennzeichnet: Verwenden Sie Einstufen."), MessageBoxIcon.Information);
                return;
            }
            SystemSettings settings = SystemSettings.Read();
            string system = ChooseSystem(host, settings, false);
            switch (system)
            {
                case null:
                    if (AvailableSystems(settings).Count == 0)
                    {
                        Show(Text.T("Portion marks belong to the IC, French and German systems; this computer uses the " + settings.Ceiling.ClassificationSystem + " system.",
                            "Les marques de paragraphe existent dans les systèmes IC, français et allemand ; ce poste utilise le système " + settings.Ceiling.ClassificationSystem + ".",
                            "Teilkennzeichnungen gibt es im IC-, im französischen und im deutschen System; dieser Rechner verwendet das System " + settings.Ceiling.ClassificationSystem + "."), MessageBoxIcon.Information);
                    }
                    return;
                case DocumentHost.BritishSystem:
                    Show(Text.T("The Government Security Classifications Policy has no portion marking: classify the document as a whole.",
                        "La Government Security Classifications Policy ne prévoit pas de marquage par paragraphe : classifier le document en entier.",
                        "Die Government Security Classifications Policy kennt keine Teilkennzeichnung: Stufen Sie das Dokument als Ganzes ein."), MessageBoxIcon.Information);
                    return;
                case DocumentHost.FrenchSystem:
                    {
                        FrenchForm french = new FrenchForm(true, lastFrenchPortion, false, false, settings);
                        if (french.ShowDialog(new OfficeWindow()) != DialogResult.OK || french.Result == null) return;
                        lastFrenchPortion = french.Result;
                        host.InsertPortionMark(FrenchFormatter.Portion(french.Result));
                        RollUpDocument(false);
                        return;
                    }
                case DocumentHost.GermanSystem:
                    {
                        GermanForm german = new GermanForm(true, lastGermanPortion, settings);
                        if (german.ShowDialog(new OfficeWindow()) != DialogResult.OK || german.Result == null) return;
                        lastGermanPortion = german.Result;
                        host.InsertPortionMark(GermanFormatter.Portion(german.Result.Level));
                        RollUpDocument(false);
                        return;
                    }
            }
            Marking banner = host.StoredBanner();
            MarkingForm form = new MarkingForm(true, lastPortion, banner, settings);
            if (form.ShowDialog(new OfficeWindow()) != DialogResult.OK || form.Result == null)
            {
                return;
            }
            string ceiling = settings.CheckCeiling(form.Result);
            if (ceiling != null)
            {
                Show(ceiling, MessageBoxIcon.Warning);
                return;
            }
            lastPortion = form.Result;
            host.InsertPortionMark("(" + MarkingFormatter.Portion(form.Result, banner) + ")");
            RollUpDocument(false);
        }

        //rolls the portion marks up into the banner line or the stamps; verbose shows the outcome even when all is well
        public bool RollUpDocument(bool verbose)
        {
            DocumentHost host = Host();
            if (host == null) return false;
            SystemSettings settings = SystemSettings.Read();
            MarkingFormatter.Abbreviate = !settings.SpellOut;
            PortionScan scan = host.ScanPortions();
            if (scan.Errors.Count > 0)
            {
                Show(Text.T("Some portion marks are not valid, fix them first:\n\n", "Des marques de paragraphe ne sont pas valides, les corriger d'abord :\n\n", "Einige Teilkennzeichnungen sind ungültig, korrigieren Sie sie zuerst:\n\n") + string.Join("\n", scan.Errors.ToArray()), MessageBoxIcon.Warning);
                return false;
            }
            if (scan.Systems > 1)
            {
                Show(Text.T("This document mixes the marks of several systems (IC \"(S//NF)\", French \"[S]\", German \"[Geh.]\"): a document belongs to one system, and their levels are not equivalent.",
                    "Ce document mélange les marques de plusieurs systèmes (IC « (S//NF) », français « [S] », allemand « [Geh.] ») : un document relève d'un seul système, sans équivalence entre leurs niveaux.",
                    "Dieses Dokument mischt die Kennzeichnungen mehrerer Systeme (IC „(S//NF)“, französisch „[S]“, deutsch „[Geh.]“): Ein Dokument gehört zu einem System, und ihre Geheimhaltungsgrade sind nicht gleichwertig."), MessageBoxIcon.Warning);
                return false;
            }
            if (scan.FrenchMarks.Count > 0) return RollUpFrench(host, settings, scan, verbose);
            if (scan.GermanMarks.Count > 0) return RollUpGerman(host, settings, scan, verbose);
            if (scan.Marks.Count == 0)
            {
                if (verbose) Show(Text.T("The document has no portion marks yet: use Portion mark on each paragraph.", "Le document n'a pas encore de marques de paragraphe : utiliser Portion mark sur chaque paragraphe.", "Das Dokument hat noch keine Teilkennzeichnungen: Verwenden Sie Teilkennzeichnung für jeden Absatz."), MessageBoxIcon.Information);
                return false;
            }
            RollUpOptions options = new RollUpOptions();
            options.PreviousBanner = host.StoredBanner();
            options.DisseminatedAnalyticProduct = settings.AnalyticProducts;
            options.DateOfOrigin = host.StoredAuthorityBlock().DateOfOrigin;
            RollUpResult result = RollUp.Compute(scan.Marks, options);
            List<string> problems = new List<string>(result.Errors);
            problems.AddRange(MarkingValidator.Validate(result.Banner));
            string ceiling = settings.CheckCeiling(result.Banner);
            if (ceiling != null) problems.Add(ceiling);
            if (problems.Count > 0)
            {
                Show(Text.T("The banner line cannot be set:\n\n", "La ligne de bannière ne peut pas être fixée :\n\n", "Die Bannerzeile kann nicht gesetzt werden:\n\n") + string.Join("\n", problems.ToArray()), MessageBoxIcon.Warning);
                return false;
            }
            host.Mark(result.Banner);
            StringBuilder report = new StringBuilder();
            report.Append(Text.T("Banner line: ", "Ligne de bannière : ", "Bannerzeile: ")).Append(MarkingFormatter.Banner(result.Banner)).Append('\n');
            if (scan.Unmarked > 0 && result.Banner.HasControls | result.Banner.Level > Level.Unclassified)
            {
                report.Append('\n').Append(scan.Unmarked).Append(Text.T(" paragraph(s) have no portion mark: every portion of a classified or controlled document must be marked.\n", " paragraphe(s) sans marque : chaque partie d'un document classifié ou contrôlé doit être marquée.\n", " Absatz/Absätze ohne Teilkennzeichnung: Jeder Teil eines eingestuften oder kontrollierten Dokuments muss gekennzeichnet sein.\n"));
            }
            foreach (string note in result.Notes)
            {
                report.Append('\n').Append(note).Append('\n');
            }
            if (verbose || scan.Unmarked > 0 || result.Notes.Count > 0)
            {
                Show(report.ToString(), MessageBoxIcon.Information);
            }
            SuggestAuthorityBlock(host, result.Banner);
            return true;
        }

        private bool RollUpFrench(DocumentHost host, SystemSettings settings, PortionScan scan, bool verbose)
        {
            if (!string.IsNullOrEmpty(host.StoredDeclassification()))
            {
                Show(Text.T("This document was declassified (" + host.StoredDeclassification() + "): its paragraph marks do not classify it again. To classify it again, use Classify.",
                    "Ce document a été déclassifié (" + host.StoredDeclassification() + ") : ses marques de paragraphe ne le reclassifient pas. Pour le classifier à nouveau, utiliser Classify.",
                    "Die Einstufung dieses Dokuments wurde aufgehoben (" + host.StoredDeclassification() + "): Seine Absatzkennzeichnungen stufen es nicht erneut ein. Um es erneut einzustufen, verwenden Sie Einstufen."), MessageBoxIcon.Information);
                return false;
            }
            FrenchRollUpResult result = FrenchRollUp.Compute(scan.FrenchMarks);
            List<string> problems = new List<string>(result.Errors);
            string ceiling = settings.CheckCeiling(result.Document);
            if (ceiling != null) problems.Add(ceiling);
            if (problems.Count > 0)
            {
                Show(Text.T("The level of the document cannot be set:\n\n", "Le niveau du document ne peut pas être fixé :\n\n", "Der Geheimhaltungsgrad des Dokuments kann nicht gesetzt werden:\n\n") + string.Join("\n", problems.ToArray()), MessageBoxIcon.Warning);
                return false;
            }
            FrenchMarking stored = host.StoredFrenchMarking();
            if (stored != null && result.Document.IsClassified)
            {
                result.Document.MentionParticuliere = stored.MentionParticuliere; //kept from the document
            }
            if (result.Document.Level == FrenchLevel.NonProtege)
            {
                Show(Text.T("Every paragraph is [NP]: the document is neither classified nor protected, no stamp is applied.", "Tous les paragraphes sont [NP] : le document n'est ni classifié ni protégé, aucun timbre n'est apposé.", "Alle Absätze sind [NP]: Das Dokument ist weder eingestuft noch geschützt, es wird kein Stempel angebracht."), MessageBoxIcon.Information);
            }
            host.Mark(result.Document, host.StoredFrenchIdentification(), host.StoredBound());
            StringBuilder report = new StringBuilder();
            report.Append(Text.T("Level of the document: ", "Niveau du document : ", "Geheimhaltungsgrad des Dokuments: ")).Append(result.Document.ToString()).Append('\n');
            if (scan.Unmarked > 0 && result.Document.Level > FrenchLevel.NonProtege)
            {
                report.Append('\n').Append(scan.Unmarked).Append(Text.T(" paragraph(s) without a mark: the IGI 1300 recommends marking the level in the margin of each part (7.1.1.2 a).\n", " paragraphe(s) sans marque : l'IGI 1300 recommande de préciser le niveau en marge de chaque partie (7.1.1.2 a).\n", " Absatz/Absätze ohne Kennzeichnung: Die IGI 1300 empfiehlt, den Geheimhaltungsgrad am Rand jedes Teils anzugeben (7.1.1.2 a).\n"));
            }
            if (verbose || scan.Unmarked > 0)
            {
                Show(report.ToString(), MessageBoxIcon.Information);
            }
            if (host is WordHost && host.StoredFrenchIdentification().Validate(result.Document).Count > 0 && verbose
                && Ask(Text.T("The first page identification is incomplete. Write it now?", "L'identification de première page est incomplète. L'écrire maintenant ?", "Die Kennzeichnung der ersten Seite ist unvollständig. Jetzt schreiben?")))
            {
                WriteIdentification(host, result.Document);
            }
            return true;
        }

        private bool RollUpGerman(DocumentHost host, SystemSettings settings, PortionScan scan, bool verbose)
        {
            GermanMarking document = GermanRollUp.Compute(scan.GermanMarks);
            string ceiling = GermanSettings.CheckCeiling(settings, document);
            if (ceiling != null)
            {
                Show(ceiling, MessageBoxIcon.Warning);
                return false;
            }
            GermanMarking stored = host.StoredGermanMarking();
            if (stored != null && document.IsClassified)
            {
                document.Vermerke = stored.Vermerke; //set by the issuer on the whole document
            }
            MarkGerman(host, document, verbose);
            StringBuilder report = new StringBuilder();
            report.Append(Text.T("Level of the document: ", "Niveau du document : ", "Geheimhaltungsgrad des Dokuments: ")).Append(document.IsClassified ? GermanFormatter.Stamp(document.Level) : Text.T("not classified", "non classifié", "nicht eingestuft")).Append('\n');
            if (scan.Unmarked > 0 && document.IsClassified)
            {
                report.Append('\n').Append(scan.Unmarked).Append(Text.T(" paragraph(s) without a mark: the start and end of each part with its own level must be recognisable (VSA § 20 Abs. 6).\n", " paragraphe(s) sans marque : le début et la fin de chaque partie de niveau propre doivent être reconnaissables (VSA § 20 al. 6).\n", " Absatz/Absätze ohne Kennzeichnung: Anfang und Ende jedes Teils mit eigenem Geheimhaltungsgrad müssen erkennbar sein (VSA § 20 Abs. 6).\n"));
            }
            if (verbose || scan.Unmarked > 0)
            {
                Show(report.ToString(), MessageBoxIcon.Information);
            }
            return true;
        }

        private void SuggestAuthorityBlock(DocumentHost host, Marking banner)
        {
            if (banner.Owner == Owner.NonUS || banner.Owner == Owner.Nato || banner.Level == Level.Unclassified)
            {
                return; //only US and JOINT classified documents carry the block
            }
            if (host.StoredAuthorityBlock().Validate(banner).Count > 0 && host is WordHost)
            {
                if (Ask(Text.T("A classified document carries a classification authority block. Write it now?", "Un document classifié porte un bloc d'autorité de classification. L'écrire maintenant ?", "Ein eingestuftes Dokument trägt einen Einstufungsvermerk (Classification Authority Block). Jetzt schreiben?")))
                {
                    WriteAuthorityBlock();
                }
            }
        }

        //the IC classification authority block, or the French and German first page
        public void WriteAuthorityBlock()
        {
            DocumentHost host = Host();
            if (host == null) return;
            string system = host.StoredSystem();
            if (system == DocumentHost.FrenchSystem)
            {
                FrenchMarking marking = host.StoredFrenchMarking();
                if (marking == null || marking.Level == FrenchLevel.NonProtege)
                {
                    Show(Text.T("Classify the document first.", "Classifier d'abord le document.", "Stufen Sie das Dokument zuerst ein."), MessageBoxIcon.Information);
                    return;
                }
                WriteIdentification(host, marking);
                return;
            }
            if (system == DocumentHost.GermanSystem)
            {
                GermanMarking marking = host.StoredGermanMarking();
                if (marking == null || !marking.IsClassified)
                {
                    Show(Text.T("Classify the document first.", "Classifier d'abord le document.", "Stufen Sie das Dokument zuerst ein."), MessageBoxIcon.Information);
                    return;
                }
                WriteGermanIdentification(host, marking);
                return;
            }
            if (system == DocumentHost.BritishSystem)
            {
                Show(Text.T("The Government Security Classifications Policy has no authority block: the markings at the top and bottom of each page are the whole marking.",
                    "La Government Security Classifications Policy ne prévoit pas de bloc d'autorité : les marquages en haut et en bas de chaque page suffisent.",
                    "Die Government Security Classifications Policy kennt keinen Einstufungsvermerk: Die Kennzeichnungen oben und unten auf jeder Seite sind die gesamte Kennzeichnung."), MessageBoxIcon.Information);
                return;
            }
            Marking banner = host.StoredBanner();
            if (banner == null)
            {
                Show(Text.T("Classify the document first.", "Classifier d'abord le document.", "Stufen Sie das Dokument zuerst ein."), MessageBoxIcon.Information);
                return;
            }
            AuthorityForm form = new AuthorityForm(host.StoredAuthorityBlock(), banner);
            if (form.ShowDialog(new OfficeWindow()) != DialogResult.OK || form.Result == null)
            {
                return;
            }
            host.StoreAuthorityBlock(form.Result);
            if (host is WordHost)
            {
                host.WriteAuthorityBox(form.Result.Lines());
            }
        }

        private void WriteIdentification(DocumentHost host, FrenchMarking marking)
        {
            IdentificationForm form = new IdentificationForm(host.StoredFrenchIdentification(), marking);
            if (form.ShowDialog(new OfficeWindow()) != DialogResult.OK || form.Result == null)
            {
                return;
            }
            host.StoreFrenchIdentification(form.Result);
            host.Mark(marking, form.Result, host.StoredBound()); //the TRÈS SECRET reference goes on every page
            if (host is WordHost)
            {
                host.WriteBlock(DocumentHost.IdentificationBookmark, form.Result.Lines(marking));
            }
        }

        private void WriteGermanIdentification(DocumentHost host, GermanMarking marking)
        {
            GermanIdentificationForm form = new GermanIdentificationForm(host.StoredGermanIdentification(), marking);
            if (form.ShowDialog(new OfficeWindow()) != DialogResult.OK || form.Result == null)
            {
                return;
            }
            host.StoreGermanIdentification(form.Result);
            if (host is WordHost)
            {
                host.WriteBlock(DocumentHost.IdentificationBookmark, form.Result.Lines(marking));
            }
        }

        //lowering, raising or lifting the classification: annexe 38 of the IGI 1300, § 18 and § 19 of the VS-Anweisung
        public void Downgrade()
        {
            DocumentHost host = Host();
            if (host == null) return;
            string system = host.StoredSystem();
            if (system == DocumentHost.GermanSystem)
            {
                DowngradeGerman(host);
                return;
            }
            FrenchMarking marking = system == DocumentHost.FrenchSystem ? host.StoredFrenchMarking() : null;
            if (marking == null || !marking.IsClassified)
            {
                Show(Text.T("Downgrade works on French documents classified Secret or Très Secret (IGI 1300, annexe 38) and German classified documents (VSA § 18, § 19).",
                    "Le déclassement concerne les documents français classifiés Secret ou Très Secret (IGI 1300, annexe 38) et les documents allemands classifiés (VSA § 18, § 19).",
                    "Herabstufen gilt für französische Dokumente, die als Secret oder Très Secret eingestuft sind (IGI 1300, annexe 38), und für deutsche Verschlusssachen (VSA § 18, § 19)."), MessageBoxIcon.Information);
                return;
            }
            SystemSettings settings = SystemSettings.Read();
            ChangeForm form = new ChangeForm(marking, settings);
            if (form.ShowDialog(new OfficeWindow()) != DialogResult.OK)
            {
                return;
            }
            FrenchMarking changed = marking.Clone();
            switch (form.Result)
            {
                case ChangeForm.Change.Declassement:
                    changed.Level = FrenchLevel.Secret;
                    break;
                case ChangeForm.Change.Reclassement:
                    changed.Level = FrenchLevel.TresSecret;
                    break;
                default:
                    changed.Level = FrenchLevel.NonProtege;
                    changed.SpecialFrance = false;
                    changed.MentionParticuliere = null;
                    break;
            }
            host.Mark(changed, host.StoredFrenchIdentification(), host.StoredBound());
            if (form.Result == ChangeForm.Change.Declassification)
            {
                host.StoreDeclassification(string.Join(" ", form.Lines));
                host.SetProperty(DocumentHost.ClassificationProperty, "DÉCLASSIFIÉ"); //classified once: not an unclassified document for the save check
            }
            if (host is WordHost)
            {
                host.WriteBlock(DocumentHost.ChangeBookmark, form.Lines);
            }
            if (form.Result != ChangeForm.Change.Declassification && host.SupportsPortions && host.ScanPortions().FrenchMarks.Count > 0)
            {
                Show(Text.T("The [TS] or [S] paragraph marks of the document still need updating: the roll-up follows them.", "Les marques de paragraphe [TS] ou [S] du document restent à mettre à jour : le calcul du niveau les suit.", "Die Absatzkennzeichnungen [TS] oder [S] des Dokuments müssen noch aktualisiert werden: Die Ableitung richtet sich nach ihnen."), MessageBoxIcon.Information);
            }
        }

        private void DowngradeGerman(DocumentHost host)
        {
            GermanMarking marking = host.StoredGermanMarking();
            if (marking == null || !marking.IsClassified)
            {
                Show(Text.T("The document is not classified.", "Le document n'est pas classifié.", "Das Dokument ist nicht eingestuft."), MessageBoxIcon.Information);
                return;
            }
            GermanChangeForm form = new GermanChangeForm(marking);
            if (form.ShowDialog(new OfficeWindow()) != DialogResult.OK || form.Result == null)
            {
                return;
            }
            host.Mark(form.Result);
            if (!form.Result.IsClassified)
            {
                host.SetProperty(DocumentHost.ClassificationProperty, "VS-Einstufung aufgehoben"); //classified once: not an unclassified document for the save check
            }
            if (host is WordHost)
            {
                host.WriteBlock(DocumentHost.ChangeBookmark, form.Lines);
            }
            if (host.SupportsPortions && host.ScanPortions().GermanMarks.Count > 0)
            {
                Show(Text.T("The part marks of the document still need updating: the roll-up follows them.", "Les marques de partie du document restent à mettre à jour : le calcul du niveau les suit.", "Die Teilkennzeichnungen des Dokuments müssen noch aktualisiert werden: Die Ableitung richtet sich nach ihnen."), MessageBoxIcon.Information);
            }
        }

        //true to cancel saving or printing: the computer requires a classification and the document has none
        public bool BlockUnclassified(object document, string action)
        {
            try
            {
                SystemSettings settings = SystemSettings.Read();
                if (!settings.RequireClassification)
                {
                    return false;
                }
                DocumentHost host = DocumentHost.For(application, document);
                if (host == null || !string.IsNullOrEmpty(host.StoredClassification()))
                {
                    return false;
                }
                ChooseLanguage();
                string what = action == "printed" ? Text.T("printed", "imprimé", "gedruckt") : Text.T("saved", "enregistré", "gespeichert");
                if (MessageBox.Show(new OfficeWindow(), Text.T("This document has no classification. It must be classified before it is " + what + ".\n\nClassify it now?",
                    "Ce document n'a pas de classification. Il doit être classifié avant d'être " + what + ".\n\nLe classifier maintenant ?",
                    "Dieses Dokument ist nicht eingestuft. Es muss eingestuft werden, bevor es " + what + " wird.\n\nJetzt einstufen?"),
                    Title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    Classify();
                }
                return string.IsNullOrEmpty(host.StoredClassification());
            }
            catch (Exception ex)
            {
                Log.Error("save and print check", ex);
                return false; //never lose the user's work because the check itself failed
            }
        }
    }
}
