//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //the classification authority block: Classified By, Reason or Derived From, Declassify On, checked as it is typed
    public class AuthorityForm : Form
    {
        private readonly Marking banner;
        private RadioButton original, derivative;
        private TextBox classifiedBy, agency, reason, derivedFrom, sources, declassifyOn, dateOfOrigin;
        private ListBox problems;
        private Button ok;
        public AuthorityBlock Result;

        public AuthorityForm(AuthorityBlock initial, Marking banner)
        {
            this.banner = banner;
            Text = "SystemBanner - " + Markings.Text.T("Classification authority block", "Bloc d'autorité de classification", "Einstufungsvermerk");
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Tahoma", 8.25f);
            int y = 10;
            original = new RadioButton();
            original.Text = Markings.Text.T("Original classification (OCA)", "Classification originale (OCA)", "Originäre Einstufung (OCA)");
            original.SetBounds(10, y, 200, 20);
            derivative = new RadioButton();
            derivative.Text = Markings.Text.T("Derivative classification", "Classification dérivée", "Abgeleitete Einstufung");
            derivative.SetBounds(220, y, 200, 20);
            derivative.Checked = true;
            original.CheckedChanged += UpdateChecks;
            Controls.Add(original);
            Controls.Add(derivative);
            y += 28;
            classifiedBy = AddText("Classified By", ref y, 1);
            agency = AddText(Markings.Text.T("Agency and office", "Agence et bureau", "Behörde und Dienststelle"), ref y, 1);
            reason = AddText("Reason (EO 13526 1.4)", ref y, 1);
            derivedFrom = AddText("Derived From", ref y, 1);
            sources = AddText(Markings.Text.T("Sources (one a line)", "Sources (une par ligne)", "Quellen (eine pro Zeile)"), ref y, 4);
            declassifyOn = AddText("Declassify On", ref y, 1);
            dateOfOrigin = AddText(Markings.Text.T("Date of origin", "Date d'origine", "Ursprungsdatum"), ref y, 1);
            string required = AuthorityBlock.RequiredNotApplicable(banner);
            Label hint = new Label();
            hint.Text = required != null ? Markings.Text.T("This document must read: ", "Ce document doit porter : ", "Dieses Dokument muss lauten: ") + required : Markings.Text.T("YYYYMMDD, an event, 25X#, YYYYMMDD, 50X1-HUM, 50X2-WMD...", "AAAAMMJJ, un événement, 25X#, AAAAMMJJ, 50X1-HUM, 50X2-WMD...", "JJJJMMTT, ein Ereignis, 25X#, JJJJMMTT, 50X1-HUM, 50X2-WMD...");
            hint.ForeColor = SystemColors.GrayText;
            hint.SetBounds(150, y - 4, 400, 30);
            Controls.Add(hint);
            y += 28;
            problems = new ListBox();
            problems.SetBounds(10, y, 540, 70);
            problems.ForeColor = Color.DarkRed;
            problems.HorizontalScrollbar = true;
            Controls.Add(problems);
            y += 76;
            ok = new Button();
            ok.Text = Markings.Text.T("Write", "Écrire", "Schreiben");
            ok.SetBounds(390, y, 75, 24);
            ok.Click += delegate { DialogResult = DialogResult.OK; Close(); };
            Controls.Add(ok);
            Button cancel = new Button();
            cancel.Text = Markings.Text.T("Cancel", "Annuler", "Abbrechen");
            cancel.SetBounds(475, y, 75, 24);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            CancelButton = cancel;
            ClientSize = new Size(560, y + 34);
            if (initial != null)
            {
                original.Checked = initial.Original;
                derivative.Checked = !initial.Original;
                classifiedBy.Text = initial.ClassifiedBy ?? "";
                agency.Text = initial.AgencyAndOffice ?? "";
                reason.Text = initial.Reason ?? "";
                derivedFrom.Text = initial.DerivedFrom ?? "";
                sources.Text = string.Join("\r\n", initial.Sources.ToArray());
                declassifyOn.Text = initial.DeclassifyOn ?? required ?? "";
                dateOfOrigin.Text = initial.DateOfOrigin.ToString("yyyyMMdd");
            }
            UpdateChecks(null, EventArgs.Empty);
        }

        private TextBox AddText(string label, ref int y, int lines)
        {
            Label caption = new Label();
            caption.Text = label;
            caption.SetBounds(10, y + 3, 140, 16);
            Controls.Add(caption);
            TextBox box = new TextBox();
            box.Multiline = lines > 1;
            if (lines > 1) box.ScrollBars = ScrollBars.Vertical;
            box.SetBounds(150, y, 400, 20 + 14 * (lines - 1));
            box.TextChanged += UpdateChecks;
            Controls.Add(box);
            y += 26 + 14 * (lines - 1);
            return box;
        }

        private void UpdateChecks(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            reason.Enabled = original.Checked;
            derivedFrom.Enabled = !original.Checked;
            sources.Enabled = !original.Checked;
            AuthorityBlock block = new AuthorityBlock();
            block.Original = original.Checked;
            block.ClassifiedBy = classifiedBy.Text.Trim();
            block.AgencyAndOffice = agency.Text.Trim();
            block.Reason = reason.Text.Trim();
            block.DerivedFrom = derivedFrom.Text.Trim();
            foreach (string line in sources.Text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                block.Sources.Add(line.Trim());
            }
            block.DeclassifyOn = declassifyOn.Text.Trim();
            List<string> errors = new List<string>();
            DateTime origin;
            if (DateTime.TryParseExact(dateOfOrigin.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out origin))
            {
                block.DateOfOrigin = origin;
            }
            else
            {
                errors.Add(Markings.Text.T("The date of origin is written YYYYMMDD.", "La date d'origine s'écrit AAAAMMJJ.", "Das Ursprungsdatum wird JJJJMMTT geschrieben."));
            }
            errors.AddRange(block.Validate(banner));
            problems.Items.Clear();
            foreach (string error in errors)
            {
                problems.Items.Add(error);
            }
            ok.Enabled = errors.Count == 0;
            Result = errors.Count == 0 ? block : null;
        }
    }
}
