//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //the level of an EU classified document or paragraph, with the additional markings, reference and copy number of the document
    public class EuForm : FrenchDialog
    {
        private readonly bool portionMode;
        private readonly SystemSettings settings;
        private readonly RadioButton[] levels = new RadioButton[4];
        private TextBox additional, reference, copy;
        private Label preview;
        public EuMarking Result;

        public EuForm(bool portionMode, EuMarking initial, SystemSettings settings)
            : base(portionMode ? Markings.Text.T("Mark the paragraph (EU)", "Marquer le paragraphe (UE)", "Absatz kennzeichnen (EU)") : Markings.Text.T("Classify the document (EU)", "Classifier le document (UE)", "Dokument einstufen (EU)"))
        {
            this.portionMode = portionMode;
            this.settings = settings;
            int y = 10;
            AddLabel(Markings.Text.T("Classification level", "Niveau de classification", "Geheimhaltungsgrad"), 10, y, 150);
            foreach (EuLevel level in new EuLevel[] { EuLevel.Restreint, EuLevel.Confidentiel, EuLevel.Secret, EuLevel.TresSecret })
            {
                RadioButton radio = new RadioButton();
                radio.Text = EuFormatter.Marking(level) + "  [" + EuFormatter.Abbreviation(level) + "]";
                radio.SetBounds(160, y, 330, 20);
                radio.CheckedChanged += OnChange;
                Controls.Add(radio);
                levels[(int)level] = radio;
                y += 22;
            }
            y += 6;
            if (!portionMode)
            {
                reference = AddText(Markings.Text.T("Reference number", "Numéro de référence", "Aktenzeichen"), ref y, OnChange);
                copy = AddText(Markings.Text.T("Copy number", "N° d'exemplaire", "Ausfertigung Nr."), ref y, OnChange);
                AddLabel(Markings.Text.T("Additional markings", "Marquages additionnels", "Zusatzvermerke"), 10, y, 150);
                additional = new TextBox();
                additional.Multiline = true;
                additional.ScrollBars = ScrollBars.Vertical;
                additional.SetBounds(160, y, 330, 52);
                additional.TextChanged += OnChange;
                Controls.Add(additional);
                y += 56;
                Label hint = AddLabel(Markings.Text.T("One per line, as written: originator, caveats, releasability, downgrading date or event.",
                    "Un par ligne, tels qu'écrits : origine, restrictions, communicabilité, date ou événement de déclassement.",
                    "Eine pro Zeile, wie geschrieben: Urheber, Einschränkungen, Freigabe, Datum oder Ereignis der Herabstufung."), 160, y, 330);
                hint.ForeColor = Color.Gray;
                hint.Height = 30;
                y += 34;
            }
            preview = new Label();
            preview.Font = new Font("Arial", 10f, FontStyle.Bold);
            preview.TextAlign = ContentAlignment.MiddleCenter;
            preview.BorderStyle = BorderStyle.FixedSingle;
            preview.SetBounds(10, y, 480, 34);
            Controls.Add(preview);
            y += 40;
            AddFooter(ref y, portionMode ? Markings.Text.T("Mark", "Marquer", "Kennzeichnen") : Markings.Text.T("Classify", "Classifier", "Einstufen"));
            EuMarking m = initial ?? new EuMarking();
            levels[(int)m.Level].Checked = true;
            if (!portionMode)
            {
                reference.Text = m.Reference ?? "";
                copy.Text = m.Copy > 0 ? m.Copy.ToString() : "";
                additional.Text = string.Join(Environment.NewLine, m.Additional.ToArray());
            }
            OnChange(null, EventArgs.Empty);
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            EuMarking m = new EuMarking();
            for (int i = 0; i < levels.Length; i++)
            {
                if (levels[i].Checked) m.Level = (EuLevel)i;
            }
            List<string> errors = new List<string>();
            if (!portionMode)
            {
                m.Reference = reference.Text.Trim().Length > 0 ? reference.Text.Trim() : null;
                int number;
                if (copy.Text.Trim().Length > 0 && !int.TryParse(copy.Text.Trim(), out number))
                {
                    errors.Add(Markings.Text.T("The copy number is a number.", "Le numéro d'exemplaire est un nombre.", "Die Ausfertigungsnummer ist eine Zahl."));
                }
                else if (int.TryParse(copy.Text.Trim(), out number))
                {
                    m.Copy = number;
                }
                copy.Enabled = m.Level >= EuLevel.Secret; //CSR Annex III pt 13(e): SECRET UE/EU SECRET and above
                if (!copy.Enabled) m.Copy = 0;
                foreach (string line in additional.Lines)
                {
                    if (line.Trim().Length > 0) m.Additional.Add(line.Trim().ToUpperInvariant());
                }
                errors.AddRange(EuValidator.Validate(m));
            }
            string ceiling = EuSettings.CheckCeiling(settings, m);
            if (ceiling != null) errors.Add(ceiling);
            preview.Text = portionMode ? EuFormatter.Portion(m.Level) : EuFormatter.Marking(m.Level);
            ShowProblems(errors);
            Result = errors.Count == 0 ? m : null;
        }
    }
}
