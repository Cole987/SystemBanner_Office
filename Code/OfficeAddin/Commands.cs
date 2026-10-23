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

        public Commands(object application)
        {
            this.application = application;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        //dialogs stay in front of the Office window; Office 2000 has no window handle property, the foreground window is it
        private class OfficeWindow : IWin32Window
        {
            private readonly IntPtr handle = GetForegroundWindow();
            public IntPtr Handle
            {
                get { return handle; }
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
                Show("Open a document first.", MessageBoxIcon.Information);
            }
            return host;
        }

        //every button runs here, so an Office error shows as a message instead of unloading the add-in
        public void Run(string command)
        {
            try
            {
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
                Show("The command failed: " + inner.Message + Environment.NewLine + Environment.NewLine + "Details are in " + Log.FilePath, MessageBoxIcon.Error);
            }
        }

        //IC or France for this document: the computer's Simple value decides, the classifier chooses when both are allowed;
        //null when the classifier cancels, or when the computer is set to another national system
        private string ChooseSystem(DocumentHost host, SystemSettings settings, bool ask)
        {
            string stored = host.StoredSystem();
            if (settings.IcAvailable && !settings.FrenchAvailable) return DocumentHost.IcSystem;
            if (settings.FrenchAvailable && !settings.IcAvailable) return DocumentHost.FrenchSystem;
            if (!settings.IcAvailable && !settings.FrenchAvailable) return null;
            if (!ask && (stored == DocumentHost.IcSystem || stored == DocumentHost.FrenchSystem)) return stored;
            SystemForm form = new SystemForm(stored ?? DocumentHost.FrenchSystem);
            return form.ShowDialog(new OfficeWindow()) == DialogResult.OK ? form.Result : null;
        }

        public void Classify()
        {
            DocumentHost host = Host();
            if (host == null) return;
            SystemSettings settings = SystemSettings.Read();
            MarkingFormatter.Abbreviate = !settings.SpellOut;
            if (!settings.IcAvailable && !settings.FrenchAvailable)
            {
                //German, British, NATO and UN computers: the stock levels of SystemBanner
                NationalForm national = new NationalForm(settings, host.StoredClassification());
                if (national.ShowDialog(new OfficeWindow()) == DialogResult.OK)
                {
                    host.Mark(national.Result);
                }
                return;
            }
            string system = ChooseSystem(host, settings, true);
            if (system == DocumentHost.FrenchSystem)
            {
                ClassifyFrench(host, settings);
            }
            else if (system == DocumentHost.IcSystem)
            {
                ClassifyIc(host, settings);
            }
        }

        private void ClassifyIc(DocumentHost host, SystemSettings settings)
        {
            if (host.SupportsPortions && host.ScanPortions().Marks.Count > 0)
            {
                //a portion marked document takes its banner from its portions, the classifier changes the portions instead
                if (Ask("This document has portion marks: its banner line is rolled up from them.\n\nRoll up the portion marks now?"))
                {
                    RollUpDocument(true);
                }
                return;
            }
            MarkingForm form = new MarkingForm(false, host.StoredBanner(), null, settings);
            if (form.ShowDialog(new OfficeWindow()) == DialogResult.OK && form.Result != null)
            {
                host.Mark(form.Result);
                SuggestAuthorityBlock(host, form.Result);
            }
        }

        private void ClassifyFrench(DocumentHost host, SystemSettings settings)
        {
            if (host.SupportsPortions && host.ScanPortions().FrenchMarks.Count > 0)
            {
                if (Ask("Ce document comporte des marques de paragraphe : son niveau est celui de son paragraphe le plus classifié (IGI 1300, 7.1.1.2).\n\nCalculer le niveau maintenant ?"))
                {
                    RollUpDocument(true);
                }
                return;
            }
            FrenchForm form = new FrenchForm(false, host.StoredFrenchMarking(), host.StoredBound(), host.SupportsCover, settings);
            if (form.ShowDialog(new OfficeWindow()) == DialogResult.OK && form.Result != null)
            {
                MarkFrench(host, form.Result, form.Bound);
            }
        }

        //stamps the document and asks for its first page identification when it is incomplete
        private void MarkFrench(DocumentHost host, FrenchMarking marking, bool bound)
        {
            FrenchIdentification id = host.StoredFrenchIdentification();
            host.StoreDeclassification(null);
            host.Mark(marking, id, bound);
            if (host is WordHost && id.Validate(marking).Count > 0
                && Ask("Un document " + marking.ToString() + " est identifié dès sa première page (autorité émettrice, date, numéro d'enregistrement" + (marking.IsClassified ? ", échéance de classification" : "") + ").\n\nÉcrire l'identification maintenant ?"))
            {
                WriteIdentification(host, marking);
            }
        }

        public void PortionMark()
        {
            DocumentHost host = Host();
            if (host == null) return;
            if (!host.SupportsPortions)
            {
                Show("Excel workbooks are marked as a whole: use Classify.", MessageBoxIcon.Information);
                return;
            }
            SystemSettings settings = SystemSettings.Read();
            string system = ChooseSystem(host, settings, false);
            if (system == null)
            {
                if (!settings.IcAvailable && !settings.FrenchAvailable)
                {
                    Show("Portion marks belong to the IC and French systems; this computer uses the " + settings.Ceiling.ClassificationSystem + " system.", MessageBoxIcon.Information);
                }
                return;
            }
            if (system == DocumentHost.FrenchSystem)
            {
                FrenchForm french = new FrenchForm(true, lastFrenchPortion, false, false, settings);
                if (french.ShowDialog(new OfficeWindow()) != DialogResult.OK || french.Result == null)
                {
                    return;
                }
                lastFrenchPortion = french.Result;
                host.InsertPortionMark(FrenchFormatter.Portion(french.Result));
                RollUpDocument(false);
                return;
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
                Show("Some portion marks are not valid, fix them first:\n\n" + string.Join("\n", scan.Errors.ToArray()), MessageBoxIcon.Warning);
                return false;
            }
            if (scan.Marks.Count > 0 && scan.FrenchMarks.Count > 0)
            {
                Show("Ce document mélange des marques IC « (S//NF) » et des marques françaises « [S] » : un document relève d'un seul système, sans équivalence entre eux.", MessageBoxIcon.Warning);
                return false;
            }
            if (scan.FrenchMarks.Count > 0)
            {
                return RollUpFrench(host, settings, scan, verbose);
            }
            if (scan.Marks.Count == 0)
            {
                if (verbose) Show("The document has no portion marks yet: use Portion mark on each paragraph.", MessageBoxIcon.Information);
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
                Show("The banner line cannot be set:\n\n" + string.Join("\n", problems.ToArray()), MessageBoxIcon.Warning);
                return false;
            }
            host.Mark(result.Banner);
            StringBuilder report = new StringBuilder();
            report.Append("Banner line: ").Append(MarkingFormatter.Banner(result.Banner)).Append('\n');
            if (scan.Unmarked > 0 && result.Banner.HasControls | result.Banner.Level > Level.Unclassified)
            {
                report.Append('\n').Append(scan.Unmarked).Append(" paragraph(s) have no portion mark: every portion of a classified or controlled document must be marked.\n");
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
                Show("Ce document a été déclassifié (" + host.StoredDeclassification() + ") : ses marques de paragraphe ne le reclassifient pas. Pour le classifier à nouveau, utiliser Classify.", MessageBoxIcon.Information);
                return false;
            }
            FrenchRollUpResult result = FrenchRollUp.Compute(scan.FrenchMarks);
            List<string> problems = new List<string>(result.Errors);
            string ceiling = settings.CheckCeiling(result.Document);
            if (ceiling != null) problems.Add(ceiling);
            if (problems.Count > 0)
            {
                Show("Le niveau du document ne peut pas être fixé :\n\n" + string.Join("\n", problems.ToArray()), MessageBoxIcon.Warning);
                return false;
            }
            FrenchMarking stored = host.StoredFrenchMarking();
            if (stored != null && result.Document.IsClassified)
            {
                result.Document.MentionParticuliere = stored.MentionParticuliere; //kept from the document
            }
            if (result.Document.Level == FrenchLevel.NonProtege)
            {
                Show("Tous les paragraphes sont [NP] : le document n'est ni classifié ni protégé, aucun timbre n'est apposé.", MessageBoxIcon.Information);
            }
            host.Mark(result.Document, host.StoredFrenchIdentification(), host.StoredBound());
            StringBuilder report = new StringBuilder();
            report.Append("Niveau du document : ").Append(result.Document.ToString()).Append('\n');
            if (scan.Unmarked > 0 && result.Document.Level > FrenchLevel.NonProtege)
            {
                report.Append('\n').Append(scan.Unmarked).Append(" paragraphe(s) sans marque : l'IGI 1300 recommande de préciser le niveau en marge de chaque partie (7.1.1.2 a).\n");
            }
            if (verbose || scan.Unmarked > 0)
            {
                Show(report.ToString(), MessageBoxIcon.Information);
            }
            if (host is WordHost && host.StoredFrenchIdentification().Validate(result.Document).Count > 0 && verbose
                && Ask("L'identification de première page est incomplète. L'écrire maintenant ?"))
            {
                WriteIdentification(host, result.Document);
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
                if (Ask("A classified document carries a classification authority block. Write it now?"))
                {
                    WriteAuthorityBlock();
                }
            }
        }

        //the IC classification authority block, or the French first page identification
        public void WriteAuthorityBlock()
        {
            DocumentHost host = Host();
            if (host == null) return;
            if (host.StoredSystem() == DocumentHost.FrenchSystem)
            {
                FrenchMarking marking = host.StoredFrenchMarking();
                if (marking == null || marking.Level == FrenchLevel.NonProtege)
                {
                    Show("Classifier d'abord le document.", MessageBoxIcon.Information);
                    return;
                }
                WriteIdentification(host, marking);
                return;
            }
            Marking banner = host.StoredBanner();
            if (banner == null)
            {
                Show("Classify the document first.", MessageBoxIcon.Information);
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
                host.WriteBlock(DocumentHost.AuthorityBookmark, form.Result.Lines());
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

        //annexe 38: lowering, raising or lifting the classification of a French document
        public void Downgrade()
        {
            DocumentHost host = Host();
            if (host == null) return;
            FrenchMarking marking = host.StoredSystem() == DocumentHost.FrenchSystem ? host.StoredFrenchMarking() : null;
            if (marking == null || !marking.IsClassified)
            {
                Show("Le déclassement et la déclassification (IGI 1300, annexe 38) concernent les documents français classifiés Secret ou Très Secret.", MessageBoxIcon.Information);
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
                Show("Les marques de paragraphe [TS] ou [S] du document restent à mettre à jour : le calcul du niveau les suit.", MessageBoxIcon.Information);
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
                if (MessageBox.Show(new OfficeWindow(), "This document has no classification. It must be classified before it is " + action + ".\n\nClassify it now?",
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
