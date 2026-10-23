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
                }
            }
            catch (Exception ex)
            {
                Exception inner = ex is System.Reflection.TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                Log.Error(command, ex);
                Show("The command failed: " + inner.Message + Environment.NewLine + Environment.NewLine + "Details are in " + Log.FilePath, MessageBoxIcon.Error);
            }
        }

        public void Classify()
        {
            DocumentHost host = Host();
            if (host == null) return;
            SystemSettings settings = SystemSettings.Read();
            MarkingFormatter.Abbreviate = !settings.SpellOut;
            if (!settings.IsUnitedStates)
            {
                NationalForm national = new NationalForm(settings, host.StoredClassification());
                if (national.ShowDialog(new OfficeWindow()) == DialogResult.OK)
                {
                    host.Mark(national.Result);
                }
                return;
            }
            if (host.SupportsPortions && host.ScanPortions().Marks.Count > 0)
            {
                //a portion marked document takes its banner from its portions, the classifier changes the portions instead
                if (MessageBox.Show(new OfficeWindow(), "This document has portion marks: its banner line is rolled up from them.\n\nRoll up the portion marks now?",
                    Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
            if (!settings.IsUnitedStates)
            {
                Show("Portion marks belong to the US IC marking system; this computer uses the " + settings.Ceiling.ClassificationSystem + " system.", MessageBoxIcon.Information);
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
            host.InsertPortionMark(MarkingFormatter.Portion(form.Result, banner));
            RollUpDocument(false);
        }

        //rolls the portion marks up into the banner line; verbose shows the outcome even when all is well
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

        private void SuggestAuthorityBlock(DocumentHost host, Marking banner)
        {
            if (banner.Owner == Owner.NonUS || banner.Owner == Owner.Nato || banner.Level == Level.Unclassified)
            {
                return; //only US and JOINT classified documents carry the block
            }
            if (host.StoredAuthorityBlock().Validate(banner).Count > 0 && host is WordHost)
            {
                if (MessageBox.Show(new OfficeWindow(), "A classified document carries a classification authority block. Write it now?", Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    WriteAuthorityBlock();
                }
            }
        }

        public void WriteAuthorityBlock()
        {
            DocumentHost host = Host();
            if (host == null) return;
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
                host.WriteAuthorityBlock(form.Result.Lines());
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
