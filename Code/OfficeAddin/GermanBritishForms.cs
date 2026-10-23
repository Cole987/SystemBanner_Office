//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //the Geheimhaltungsgrad of a document or of a part, VS-Anweisung of 13 March 2023
    public class GermanForm : FrenchDialog
    {
        private readonly bool portionMode;
        private readonly SystemSettings settings;
        private RadioButton[] levels;
        private CheckedListBox vermerke;
        private TextBox weitergabe;
        private Label preview;
        public GermanMarking Result;

        public GermanForm(bool portionMode, GermanMarking initial, SystemSettings settings)
            : base(portionMode ? Markings.Text.T("Part mark (VS-Anweisung)", "Marque de partie (VS-Anweisung)") : Markings.Text.T("Classify the document (VS-Anweisung)", "Classifier le document (VS-Anweisung)"))
        {
            this.portionMode = portionMode;
            this.settings = settings;
            int y = 10;
            AddLabel(Markings.Text.T("Level (Geheimhaltungsgrad)", "Niveau (Geheimhaltungsgrad)"), 10, y, 150);
            List<GermanLevel> choices = new List<GermanLevel>();
            if (portionMode) choices.Add(GermanLevel.Offen);
            choices.AddRange(new GermanLevel[] { GermanLevel.VsNfd, GermanLevel.VsVertraulich, GermanLevel.Geheim, GermanLevel.StrengGeheim });
            levels = new RadioButton[choices.Count];
            for (int i = 0; i < choices.Count; i++)
            {
                RadioButton radio = new RadioButton();
                radio.Text = (choices[i] == GermanLevel.Offen ? Markings.Text.T("not classified", "non classifié") : GermanFormatter.Stamp(choices[i])) + " [" + GermanFormatter.Abbreviation(choices[i]) + "]";
                radio.Tag = choices[i];
                radio.SetBounds(160, y, 330, 20);
                radio.CheckedChanged += OnChange;
                Controls.Add(radio);
                levels[i] = radio;
                y += 22;
            }
            if (!portionMode)
            {
                y += 4;
                AddLabel(Markings.Text.T("Warn- und Sperrvermerke", "Warn- und Sperrvermerke"), 10, y, 150);
                vermerke = new CheckedListBox();
                vermerke.CheckOnClick = true;
                vermerke.Items.AddRange(GermanVermerke.Sperrvermerke);
                vermerke.Items.AddRange(GermanVermerke.Warnvermerke);
                vermerke.SetBounds(160, y, 330, 100);
                vermerke.ItemCheck += delegate { BeginInvoke(new EventHandler(OnChange), this, EventArgs.Empty); };
                Controls.Add(vermerke);
                y += 106;
                weitergabe = AddText(Markings.Text.T("Weitergabe an / intern", "Weitergabe an / intern"), ref y, OnChange);
                AddLabel(Markings.Text.T("e.g. \"Weitergabe an Referat X\" or \"BMI intern\" (Anlage IV § 1.3.2)", "ex. « Weitergabe an Referat X » ou « BMI intern » (Anlage IV § 1.3.2)"), 160, y - 4, 330).ForeColor = SystemColors.GrayText;
                y += 18;
            }
            preview = new Label();
            preview.Font = new Font("Arial", 11f, FontStyle.Bold);
            preview.SetBounds(160, y, 330, 40);
            Controls.Add(preview);
            y += 46;
            AddFooter(ref y, portionMode ? Markings.Text.T("Insert", "Insérer") : Markings.Text.T("Classify", "Classifier"));
            GermanMarking start = initial ?? new GermanMarking();
            if (!portionMode && start.Level == GermanLevel.Offen) start.Level = GermanLevel.VsNfd;
            foreach (RadioButton radio in levels)
            {
                radio.Checked = (GermanLevel)radio.Tag == start.Level;
            }
            if (vermerke != null)
            {
                for (int i = 0; i < vermerke.Items.Count; i++)
                {
                    vermerke.SetItemChecked(i, start.Vermerke.Contains(Convert.ToString(vermerke.Items[i])));
                }
                foreach (string vermerk in start.Vermerke)
                {
                    if (!vermerke.Items.Contains(vermerk)) weitergabe.Text = vermerk;
                }
            }
            OnChange(null, EventArgs.Empty);
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            GermanMarking m = new GermanMarking();
            foreach (RadioButton radio in levels)
            {
                if (radio.Checked) m.Level = (GermanLevel)radio.Tag;
            }
            if (vermerke != null)
            {
                for (int i = 0; i < vermerke.Items.Count; i++)
                {
                    if (vermerke.GetItemChecked(i)) m.Vermerke.Add(Convert.ToString(vermerke.Items[i]));
                }
                if (weitergabe.Text.Trim().Length > 0) m.Vermerke.Add(weitergabe.Text.Trim());
            }
            List<string> errors = new List<string>();
            string ceiling = GermanSettings.CheckCeiling(settings, m);
            if (ceiling != null) errors.Add(ceiling);
            if (m.Vermerke.Count > 0 && !m.IsClassified) errors.Add(Markings.Text.T("Warn- und Sperrvermerke go with classified information.", "Les Warn- und Sperrvermerke accompagnent une information classifiée."));
            int[] rgb = GermanFormatter.Color(m.Level);
            preview.ForeColor = Color.FromArgb(rgb[0], rgb[1], rgb[2]);
            if (portionMode)
            {
                preview.Text = GermanFormatter.Portion(m.Level) + " " + Markings.Text.T("text of the part", "texte de la partie");
            }
            else
            {
                preview.Text = GermanFormatter.Stamp(m.Level) + (m.AmtlichGeheimGehalten ? "\n" + GermanFormatter.AmtlichGeheimGehalten : "")
                    + "\n" + (GermanFormatter.AtBottom(m.Level) ? Markings.Text.T("(top and bottom of each page)", "(en haut et en bas de chaque page)") : Markings.Text.T("(top of each page)", "(en haut de chaque page)"));
                preview.Height = 52;
            }
            ShowProblems(errors);
            Result = errors.Count == 0 ? m : null;
        }
    }

    //the first page of a German document, § 20 Abs. 2
    public class GermanIdentificationForm : FrenchDialog
    {
        private readonly GermanMarking marking;
        private TextBox herausgeber, kennzeichen, ausfertigung, empfaenger, endet;
        private DateTimePicker datum;
        public GermanIdentification Result;

        public GermanIdentificationForm(GermanIdentification initial, GermanMarking marking)
            : base(Markings.Text.T("First page (VSA § 20 Abs. 2)", "Première page (VSA § 20 al. 2)"))
        {
            this.marking = marking;
            int y = 10;
            AddLabel(GermanFormatter.Stamp(marking.Level), 10, y, 480).Font = new Font("Tahoma", 8.25f, FontStyle.Bold);
            y += 24;
            herausgeber = AddText("Herausgeber", ref y, OnChange);
            AddLabel("Datum", 10, y, 150);
            datum = AddDate(160, y, OnChange);
            y += 26;
            endet = AddText(Markings.Text.T("Classification ends with the year", "Fin de classification (année)"), ref y, OnChange);
            AddLabel(marking.Level == GermanLevel.VsNfd ? Markings.Text.T("empty: 30 years (§ 16 Abs. 1)", "vide : 30 ans (§ 16 al. 1)") : Markings.Text.T("30 years at most (§ 16 Abs. 2)", "30 ans au plus (§ 16 al. 2)"), 160, y - 4, 330).ForeColor = SystemColors.GrayText;
            y += 18;
            if (marking.AmtlichGeheimGehalten)
            {
                kennzeichen = AddText(Markings.Text.T("Reference (Tgb.-Nr.)", "Référence (Tgb.-Nr.)"), ref y, OnChange);
                ausfertigung = AddText(Markings.Text.T("Copy number", "Numéro d'exemplaire"), ref y, OnChange);
                empfaenger = AddText(Markings.Text.T("Recipient (Empfänger)", "Destinataire (Empfänger)"), ref y, OnChange);
            }
            AddFooter(ref y, Markings.Text.T("Write", "Écrire"));
            GermanIdentification id = initial ?? new GermanIdentification();
            herausgeber.Text = id.Herausgeber ?? "";
            datum.Value = id.Datum;
            endet.Text = id.EinstufungEndet > 0 ? id.EinstufungEndet.ToString() : (marking.AmtlichGeheimGehalten ? (id.Datum.Year + 10).ToString() : "");
            if (kennzeichen != null)
            {
                kennzeichen.Text = id.Kennzeichen ?? "";
                ausfertigung.Text = id.Ausfertigung > 0 ? id.Ausfertigung.ToString() : "1";
                empfaenger.Text = id.Empfaenger ?? "";
            }
            OnChange(null, EventArgs.Empty);
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            GermanIdentification id = new GermanIdentification();
            id.Herausgeber = herausgeber.Text.Trim();
            id.Datum = datum.Value.Date;
            int number;
            List<string> errors = new List<string>();
            if (endet.Text.Trim().Length > 0 && !int.TryParse(endet.Text.Trim(), out number))
            {
                errors.Add(Markings.Text.T("The end of the classification is a year.", "La fin de classification est une année."));
            }
            else if (int.TryParse(endet.Text.Trim(), out number))
            {
                id.EinstufungEndet = number;
            }
            if (kennzeichen != null)
            {
                id.Kennzeichen = kennzeichen.Text.Trim();
                if (int.TryParse(ausfertigung.Text.Trim(), out number)) id.Ausfertigung = number;
                id.Empfaenger = empfaenger.Text.Trim();
            }
            errors.AddRange(id.Validate(marking));
            ShowProblems(errors);
            Result = errors.Count == 0 ? id : null;
        }
    }

    //Herabstufung and Aufhebung of a German classification, VSA § 18 and § 19
    public class GermanChangeForm : FrenchDialog
    {
        private readonly GermanMarking marking;
        private ComboBox newLevel;
        private DateTimePicker datum;
        private TextBox stelle;
        public GermanMarking Result;
        public string[] Lines;

        public GermanChangeForm(GermanMarking marking)
            : base(Markings.Text.T("Downgrade or lift (VSA § 18, § 19)", "Déclasser ou lever (VSA § 18, § 19)"))
        {
            this.marking = marking;
            int y = 10;
            AddLabel(GermanFormatter.Stamp(marking.Level), 10, y, 480).Font = new Font("Tahoma", 8.25f, FontStyle.Bold);
            y += 24;
            AddLabel(Markings.Text.T("New level", "Nouveau niveau"), 10, y, 150);
            newLevel = new ComboBox();
            newLevel.DropDownStyle = ComboBoxStyle.DropDownList;
            for (GermanLevel level = marking.Level - 1; level >= GermanLevel.Offen; level--)
            {
                newLevel.Items.Add(level == GermanLevel.Offen ? Markings.Text.T("lifted (Aufhebung)", "levée (Aufhebung)") : GermanFormatter.Stamp(level));
            }
            newLevel.SetBounds(160, y, 330, 20);
            newLevel.SelectedIndex = 0;
            newLevel.SelectedIndexChanged += OnChange;
            Controls.Add(newLevel);
            y += 26;
            AddLabel("Datum", 10, y, 150);
            datum = AddDate(160, y, OnChange);
            y += 26;
            stelle = AddText(Markings.Text.T("Decided by (Stelle)", "Décidé par (Stelle)"), ref y, OnChange);
            AddLabel(Markings.Text.T("The VSA asks for a visible note of the change and of who decided it; the wording is a choice of this project.", "La VSA demande une mention visible du changement et de qui l'a décidé ; la formulation est un choix du projet."), 160, y - 4, 330).ForeColor = SystemColors.GrayText;
            y += 32;
            AddFooter(ref y, Markings.Text.T("Apply", "Appliquer"));
            OnChange(null, EventArgs.Empty);
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            GermanMarking changed = marking.Clone();
            changed.Level = marking.Level - 1 - newLevel.SelectedIndex;
            if (!changed.IsClassified) changed.Vermerke.Clear();
            string date = datum.Value.ToString("dd.MM.yyyy");
            Lines = changed.IsClassified
                ? new string[] { "Herabgestuft auf " + GermanFormatter.Stamp(changed.Level) + " am " + date + " durch " + stelle.Text.Trim() }
                : new string[] { "VS-Einstufung aufgehoben am " + date + " durch " + stelle.Text.Trim() };
            List<string> errors = new List<string>();
            if (stelle.Text.Trim().Length == 0) errors.Add(Markings.Text.T("Name who decided the change (VSA § 19 Abs. 1).", "Indiquer qui a décidé le changement (VSA § 19 al. 1)."));
            ShowProblems(errors);
            Result = errors.Count == 0 ? changed : null;
        }
    }

    //the classification and additional markings of the Government Security Classifications Policy
    public class BritishForm : FrenchDialog
    {
        private readonly SystemSettings settings;
        private RadioButton official, secret, topSecret;
        private CheckBox sensitive, prefix, relEu;
        private CheckedListBox caveats, descriptors;
        private ComboBox handling;
        private TextBox organisation, embargo, codeword;
        private Label preview;
        public UkMarking Result;

        public BritishForm(UkMarking initial, SystemSettings settings)
            : base(Markings.Text.T("Classify the document (GSCP)", "Classifier le document (GSCP)"))
        {
            this.settings = settings;
            int y = 10;
            AddLabel(Markings.Text.T("Classification", "Classification"), 10, y, 150);
            official = Radio("OFFICIAL", 160, y);
            secret = Radio("SECRET", 260, y);
            topSecret = Radio("TOP SECRET", 360, y);
            y += 24;
            sensitive = Check("-SENSITIVE (OFFICIAL-SENSITIVE)", 160, y);
            y += 24;
            prefix = Check(Markings.Text.T("UK prefix (sent to a foreign government or organisation)", "Préfixe UK (envoi à un État ou une organisation étrangère)"), 160, y);
            y += 22;
            relEu = Check("REL EU", 160, y);
            y += 26;
            AddLabel(Markings.Text.T("National caveats", "Caveats nationaux"), 10, y, 150);
            caveats = List(UkRegister.Caveats, y, 52);
            y += 58;
            AddLabel(Markings.Text.T("Handling instruction", "Handling instruction"), 10, y, 150);
            handling = new ComboBox();
            handling.DropDownStyle = ComboBoxStyle.DropDownList;
            handling.Items.Add(Markings.Text.T("(none)", "(aucune)"));
            foreach (string instruction in UkRegister.HandlingInstructions)
            {
                handling.Items.Add(instruction == UkRegister.OrganisationUseOnly ? "[ORGANISATION] USE ONLY" : instruction);
            }
            handling.SetBounds(160, y, 330, 20);
            handling.SelectedIndexChanged += OnChange;
            Controls.Add(handling);
            y += 26;
            organisation = AddText(Markings.Text.T("Organisation(s)", "Organisation(s)"), ref y, OnChange);
            embargo = AddText(Markings.Text.T("Embargo: date, then handling", "Embargo : date, puis traitement"), ref y, OnChange);
            codeword = AddText("Codeword", ref y, OnChange);
            AddLabel(Markings.Text.T("Descriptors", "Descripteurs"), 10, y, 150);
            descriptors = List(UkRegister.Descriptors, y, 96);
            y += 102;
            preview = new Label();
            preview.Font = new Font("Arial", 10f, FontStyle.Bold);
            preview.TextAlign = ContentAlignment.MiddleCenter;
            preview.BorderStyle = BorderStyle.FixedSingle;
            preview.SetBounds(10, y, 480, 34);
            Controls.Add(preview);
            y += 40;
            AddFooter(ref y, Markings.Text.T("Classify", "Classifier"));
            UkMarking m = initial ?? new UkMarking();
            official.Checked = m.Level == UkLevel.Official;
            secret.Checked = m.Level == UkLevel.Secret;
            topSecret.Checked = m.Level == UkLevel.TopSecret;
            sensitive.Checked = m.Sensitive;
            prefix.Checked = m.UkPrefix;
            relEu.Checked = m.RelEu;
            for (int i = 0; i < caveats.Items.Count; i++) caveats.SetItemChecked(i, m.Caveats.Contains(Convert.ToString(caveats.Items[i])));
            for (int i = 0; i < descriptors.Items.Count; i++) descriptors.SetItemChecked(i, m.Descriptors.Contains(Convert.ToString(descriptors.Items[i])));
            int index = Array.IndexOf(UkRegister.HandlingInstructions, m.HandlingInstruction);
            handling.SelectedIndex = index + 1;
            organisation.Text = m.Organisation ?? "";
            embargo.Text = m.Embargo ?? "";
            codeword.Text = m.Codeword ?? "";
            OnChange(null, EventArgs.Empty);
        }

        private RadioButton Radio(string text, int x, int y)
        {
            RadioButton radio = new RadioButton();
            radio.Text = text;
            radio.SetBounds(x, y, 100, 20);
            radio.CheckedChanged += OnChange;
            Controls.Add(radio);
            return radio;
        }

        private CheckBox Check(string text, int x, int y)
        {
            CheckBox box = new CheckBox();
            box.Text = text;
            box.SetBounds(x, y, 330, 20);
            box.CheckedChanged += OnChange;
            Controls.Add(box);
            return box;
        }

        private CheckedListBox List(string[] items, int y, int height)
        {
            CheckedListBox list = new CheckedListBox();
            list.CheckOnClick = true;
            list.Items.AddRange(items);
            list.SetBounds(160, y, 330, height);
            list.ItemCheck += delegate { BeginInvoke(new EventHandler(OnChange), this, EventArgs.Empty); };
            Controls.Add(list);
            return list;
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            UkMarking m = new UkMarking();
            m.Level = topSecret.Checked ? UkLevel.TopSecret : (secret.Checked ? UkLevel.Secret : UkLevel.Official);
            sensitive.Enabled = m.Level == UkLevel.Official;
            m.Sensitive = sensitive.Checked && sensitive.Enabled;
            if (relEu.Checked && !prefix.Checked) prefix.Checked = true; //REL EU goes with the UK prefix
            m.UkPrefix = prefix.Checked;
            m.RelEu = relEu.Checked;
            for (int i = 0; i < caveats.Items.Count; i++) if (caveats.GetItemChecked(i)) m.Caveats.Add(Convert.ToString(caveats.Items[i]));
            for (int i = 0; i < descriptors.Items.Count; i++) if (descriptors.GetItemChecked(i)) m.Descriptors.Add(Convert.ToString(descriptors.Items[i]));
            if (handling.SelectedIndex > 0) m.HandlingInstruction = UkRegister.HandlingInstructions[handling.SelectedIndex - 1];
            organisation.Enabled = m.HandlingInstruction == UkRegister.OrganisationUseOnly;
            embargo.Enabled = m.HandlingInstruction == UkRegister.Embargoed;
            if (organisation.Enabled && organisation.Text.Trim().Length > 0) m.Organisation = organisation.Text.Trim();
            if (embargo.Enabled && embargo.Text.Trim().Length > 0) m.Embargo = embargo.Text.Trim();
            if (codeword.Text.Trim().Length > 0) m.Codeword = codeword.Text.Trim();
            List<string> errors = UkValidator.Validate(m);
            string ceiling = UkSettings.CheckCeiling(settings, m);
            if (ceiling != null) errors.Add(ceiling);
            preview.Text = UkFormatter.Marking(m);
            ShowProblems(errors);
            Result = errors.Count == 0 ? m : null;
        }
    }
}
