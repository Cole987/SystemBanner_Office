//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //shared look of the French dialogs, built in code like the IC ones
    public class FrenchDialog : Form
    {
        protected ListBox problems;
        protected Button ok;

        protected FrenchDialog(string title)
        {
            Text = "SystemBanner - " + title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Tahoma", 8.25f);
        }

        protected Label AddLabel(string text, int x, int y, int width)
        {
            Label label = new Label();
            label.Text = text;
            label.SetBounds(x, y + 3, width, 16);
            Controls.Add(label);
            return label;
        }

        protected TextBox AddText(string label, ref int y, EventHandler changed)
        {
            AddLabel(label, 10, y, 150);
            TextBox box = new TextBox();
            box.SetBounds(160, y, 330, 20);
            box.TextChanged += changed;
            Controls.Add(box);
            y += 26;
            return box;
        }

        protected DateTimePicker AddDate(int x, int y, EventHandler changed)
        {
            DateTimePicker picker = new DateTimePicker();
            picker.Format = DateTimePickerFormat.Short;
            picker.SetBounds(x, y, 110, 20);
            picker.ValueChanged += changed;
            Controls.Add(picker);
            return picker;
        }

        //the problems list and the OK and Cancel buttons at the bottom
        protected void AddFooter(ref int y, string okText)
        {
            problems = new ListBox();
            problems.SetBounds(10, y, 480, 70);
            problems.ForeColor = Color.DarkRed;
            problems.HorizontalScrollbar = true;
            Controls.Add(problems);
            y += 76;
            ok = new Button();
            ok.Text = okText;
            ok.SetBounds(330, y, 75, 24);
            ok.Click += delegate { DialogResult = DialogResult.OK; Close(); };
            Controls.Add(ok);
            Button cancel = new Button();
            cancel.Text = Markings.Text.T("Cancel", "Annuler", "Abbrechen");
            cancel.SetBounds(415, y, 75, 24);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            ClientSize = new Size(500, y + 34);
        }

        protected void ShowProblems(List<string> errors)
        {
            problems.Items.Clear();
            foreach (string error in errors)
            {
                problems.Items.Add(error);
            }
            ok.Enabled = errors.Count == 0;
        }
    }

    //the level of a document or of a paragraph, IGI 1300
    public class FrenchForm : FrenchDialog
    {
        private readonly bool portionMode;
        private readonly SystemSettings settings;
        private readonly bool coverAvailable;
        private RadioButton[] levels;
        private CheckBox specialFrance, bound;
        private TextBox mention;
        private Label top, special, bottom, portion;
        public FrenchMarking Result;
        public bool Bound;

        public FrenchForm(bool portionMode, FrenchMarking initial, bool initialBound, bool coverAvailable, SystemSettings settings)
            : base(portionMode ? Markings.Text.T("Paragraph mark (IGI 1300)", "Marque de paragraphe (IGI 1300)", "Absatzkennzeichnung (IGI 1300)") : Markings.Text.T("Classify the document (IGI 1300)", "Classifier le document (IGI 1300)", "Dokument einstufen (IGI 1300)"))
        {
            this.portionMode = portionMode;
            this.settings = settings;
            this.coverAvailable = coverAvailable;
            int y = 10;
            AddLabel(Markings.Text.T("Level", "Niveau", "Geheimhaltungsgrad"), 10, y, 150);
            FrenchLevel[] choices = portionMode
                ? new FrenchLevel[] { FrenchLevel.NonProtege, FrenchLevel.DiffusionRestreinte, FrenchLevel.Secret, FrenchLevel.TresSecret }
                : new FrenchLevel[] { FrenchLevel.DiffusionRestreinte, FrenchLevel.Secret, FrenchLevel.TresSecret };
            levels = new RadioButton[choices.Length];
            for (int i = 0; i < choices.Length; i++)
            {
                RadioButton radio = new RadioButton();
                radio.Text = FrenchFormatter.Name(choices[i]) + " [" + FrenchFormatter.Abbreviation(choices[i]) + "]";
                radio.Tag = choices[i];
                radio.SetBounds(160, y, 200, 20);
                radio.CheckedChanged += OnChange;
                Controls.Add(radio);
                levels[i] = radio;
                y += 22;
            }
            specialFrance = new CheckBox();
            specialFrance.Text = Markings.Text.T("Spécial France (French nationals only)", "Spécial France (aux seuls ressortissants français)", "Spécial France (nur frz. Staatsangehörige)");
            specialFrance.SetBounds(160, y, 330, 20);
            specialFrance.CheckedChanged += OnChange;
            Controls.Add(specialFrance);
            y += 26;
            if (!portionMode)
            {
                mention = AddText(Markings.Text.T("Mention particulière", "Mention particulière", "Mention particulière"), ref y, OnChange);
                AddLabel(Markings.Text.T("services, states or organisations allowed, or special classification (7.1.1.3 b)", "services, États ou organisations autorisés, ou classification spéciale (7.1.1.3 b)", "zugelassene Stellen, Staaten oder Organisationen, oder Sondereinstufung (7.1.1.3 b)"), 160, y - 4, 330).ForeColor = SystemColors.GrayText;
                y += 18;
                bound = new CheckBox();
                bound.Text = Markings.Text.T("Bound document: cover stamp at the bottom of the first page", "Document relié : timbre de couverture en bas de la première page", "Gebundenes Dokument: Deckblattstempel unten auf Seite 1");
                bound.SetBounds(160, y, 330, 20);
                bound.Enabled = coverAvailable;
                bound.CheckedChanged += OnChange;
                Controls.Add(bound);
                y += 28;
            }
            AddLabel(portionMode ? Markings.Text.T("Mark", "Marque", "Kennzeichnung") : Markings.Text.T("Stamps", "Timbres", "Stempel"), 10, y, 150);
            if (portionMode)
            {
                portion = new Label();
                portion.Font = new Font("Arial", 11f, FontStyle.Bold);
                portion.SetBounds(160, y, 330, 22);
                Controls.Add(portion);
                y += 28;
            }
            else
            {
                top = Stamp(y, Color.FromArgb(200, 16, 46));
                special = Stamp(y + 30, Color.FromArgb(0, 51, 160));
                bottom = new Label();
                bottom.SetBounds(160, y + 60, 330, 16);
                bottom.ForeColor = SystemColors.GrayText;
                Controls.Add(bottom);
                y += 84;
            }
            AddFooter(ref y, portionMode ? Markings.Text.T("Insert", "Insérer", "Einfügen") : Markings.Text.T("Classify", "Classifier", "Einstufen"));
            FrenchMarking start = initial ?? new FrenchMarking();
            if (!portionMode && start.Level == FrenchLevel.NonProtege) start.Level = FrenchLevel.Secret;
            foreach (RadioButton radio in levels)
            {
                radio.Checked = (FrenchLevel)radio.Tag == start.Level;
            }
            specialFrance.Checked = start.SpecialFrance;
            if (mention != null) mention.Text = start.MentionParticuliere ?? "";
            if (bound != null) bound.Checked = initialBound && coverAvailable;
            OnChange(null, EventArgs.Empty);
        }

        private Label Stamp(int y, Color color)
        {
            Label label = new Label();
            label.Font = new Font("Arial", 12f, FontStyle.Bold);
            label.ForeColor = color;
            label.BorderStyle = BorderStyle.FixedSingle;
            label.TextAlign = ContentAlignment.MiddleCenter;
            label.SetBounds(220, y, 210, 26);
            Controls.Add(label);
            return label;
        }

        private FrenchMarking Build()
        {
            FrenchMarking m = new FrenchMarking();
            foreach (RadioButton radio in levels)
            {
                if (radio.Checked) m.Level = (FrenchLevel)radio.Tag;
            }
            m.SpecialFrance = specialFrance.Checked;
            if (mention != null && mention.Text.Trim().Length > 0) m.MentionParticuliere = mention.Text.Trim();
            return m;
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            FrenchMarking m = Build();
            specialFrance.Enabled = m.Level != FrenchLevel.NonProtege;
            List<string> errors = FrenchValidator.Validate(m);
            string ceiling = settings == null ? null : settings.CheckCeiling(m);
            if (ceiling != null) errors.Add(ceiling);
            if (portionMode)
            {
                portion.Text = FrenchFormatter.Portion(m) + " " + Markings.Text.T("text of the paragraph", "texte du paragraphe", "Text des Absatzes");
            }
            else
            {
                top.Text = FrenchFormatter.Stamp(m);
                special.Text = m.SpecialFrance ? FrenchFormatter.SpecialFranceStamp : "";
                special.Visible = m.SpecialFrance;
                bottom.Text = m.IsClassified ? Markings.Text.T("top and bottom of each page, Spécial France at the top only", "en haut et en bas de chaque page, Spécial France en haut seulement", "oben und unten auf jeder Seite, Spécial France nur oben") : Markings.Text.T("top of each page only (annexe 1)", "en haut de chaque page seulement (annexe 1)", "nur oben auf jeder Seite (annexe 1)");
                if (bound != null) bound.Enabled = coverAvailable && m.IsClassified;
            }
            ShowProblems(errors);
            Result = errors.Count == 0 ? m : null;
            Bound = bound != null && bound.Checked && bound.Enabled;
        }
    }

    //the first page of the document: 7.1.2.3 b) and c), annexe 1 § 2
    public class IdentificationForm : FrenchDialog
    {
        private readonly FrenchMarking marking;
        private TextBox lieu, enregistrement, autorite, auteur, reference, annexes, exemplaire, total;
        private DateTimePicker emission, echeance, reevaluation;
        private RadioButton withDeadline, withReview, without;
        public FrenchIdentification Result;

        public IdentificationForm(FrenchIdentification initial, FrenchMarking marking)
            : base(Markings.Text.T("First page identification (IGI 1300, 7.1.2.3)", "Identification (IGI 1300, 7.1.2.3)", "Kennzeichnung der ersten Seite (IGI 1300, 7.1.2.3)"))
        {
            this.marking = marking;
            int y = 10;
            AddLabel(Markings.Text.T("Document ", "Document ", "Dokument ") + marking.ToString(), 10, y, 480).Font = new Font("Tahoma", 8.25f, FontStyle.Bold);
            y += 24;
            lieu = AddText(Markings.Text.T("Place", "Lieu", "Ort"), ref y, OnChange);
            AddLabel(Markings.Text.T("Date of issue", "Date d'émission", "Ausstellungsdatum"), 10, y, 150);
            emission = AddDate(160, y, OnChange);
            y += 26;
            enregistrement = AddText(Markings.Text.T("Registration number", "N° d'enregistrement", "Registriernummer"), ref y, OnChange);
            autorite = AddText(Markings.Text.T("Issuing authority", "Autorité émettrice", "Herausgeber"), ref y, OnChange);
            auteur = AddText(Markings.Text.T("Author", "Auteur", "Verfasser"), ref y, OnChange);
            annexes = AddText(Markings.Text.T("Annexes or plans", "Annexes ou plans", "Anlagen oder Pläne"), ref y, OnChange);
            if (marking.IsClassified)
            {
                AddLabel(Markings.Text.T("Deadline", "Échéance", "Frist"), 10, y, 150);
                withDeadline = Radio(Markings.Text.T("Declassified on", "Déclassifié le", "Aufhebung am"), 160, y);
                echeance = AddDate(300, y, OnChange);
                y += 24;
                withReview = Radio(Markings.Text.T("Re-evaluate on", "À réévaluer le", "Überprüfung am"), 160, y);
                reevaluation = AddDate(300, y, OnChange);
                AddLabel(Markings.Text.T("(by exception, 20 years at most)", "(à titre exceptionnel, 20 ans au plus)", "(ausnahmsweise, höchstens 20 Jahre)"), 415, y, 90).ForeColor = SystemColors.GrayText;
                y += 24;
                without = Radio(Markings.Text.T("No deadline (NRBC weapons, 7.6.1)", "Aucune échéance (armes NRBC, 7.6.1)", "Keine Frist (ABC-Waffen, 7.6.1)"), 160, y);
                without.Width = 330;
                y += 28;
            }
            if (marking.Level == FrenchLevel.TresSecret)
            {
                AddLabel(Markings.Text.T("Copy number / of", "Exemplaire n° / sur", "Ausfertigung Nr. / von"), 10, y, 150);
                exemplaire = new TextBox();
                exemplaire.SetBounds(160, y, 50, 20);
                exemplaire.TextChanged += OnChange;
                Controls.Add(exemplaire);
                total = new TextBox();
                total.SetBounds(220, y, 50, 20);
                total.TextChanged += OnChange;
                Controls.Add(total);
                y += 26;
                reference = AddText(Markings.Text.T("Reference (every page)", "Référence (chaque page)", "Aktenzeichen (jede Seite)"), ref y, OnChange);
            }
            AddFooter(ref y, Markings.Text.T("Write", "Écrire", "Schreiben"));
            FrenchIdentification id = initial ?? new FrenchIdentification();
            lieu.Text = id.Lieu ?? "";
            emission.Value = id.DateEmission;
            enregistrement.Text = id.NumeroEnregistrement ?? "";
            autorite.Text = id.AutoriteEmettrice ?? "";
            auteur.Text = id.Auteur ?? "";
            annexes.Text = id.Annexes ?? "";
            if (marking.IsClassified)
            {
                echeance.Value = id.Echeance ?? id.DateEmission.AddYears(10);
                reevaluation.Value = id.Reevaluation ?? id.DateEmission.AddYears(FrenchIdentification.MaximumReevaluationYears);
                withReview.Checked = id.Reevaluation.HasValue;
                without.Checked = id.SansEcheance;
                withDeadline.Checked = !withReview.Checked && !without.Checked;
            }
            if (exemplaire != null)
            {
                exemplaire.Text = id.Exemplaire > 0 ? id.Exemplaire.ToString() : "";
                total.Text = id.NombreExemplaires > 0 ? id.NombreExemplaires.ToString() : "";
                reference.Text = id.Reference ?? "";
            }
            OnChange(null, EventArgs.Empty);
        }

        private RadioButton Radio(string text, int x, int y)
        {
            RadioButton radio = new RadioButton();
            radio.Text = text;
            radio.SetBounds(x, y, 135, 20);
            radio.CheckedChanged += OnChange;
            Controls.Add(radio);
            return radio;
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            FrenchIdentification id = new FrenchIdentification();
            id.Lieu = lieu.Text.Trim();
            id.DateEmission = emission.Value.Date;
            id.NumeroEnregistrement = enregistrement.Text.Trim();
            id.AutoriteEmettrice = autorite.Text.Trim();
            id.Auteur = auteur.Text.Trim();
            id.Annexes = annexes.Text.Trim();
            if (marking.IsClassified)
            {
                echeance.Enabled = withDeadline.Checked;
                reevaluation.Enabled = withReview.Checked;
                if (withDeadline.Checked) id.Echeance = echeance.Value.Date;
                if (withReview.Checked) id.Reevaluation = reevaluation.Value.Date;
                id.SansEcheance = without.Checked;
            }
            if (exemplaire != null)
            {
                int number;
                if (int.TryParse(exemplaire.Text, out number)) id.Exemplaire = number;
                if (int.TryParse(total.Text, out number)) id.NombreExemplaires = number;
                id.Reference = reference.Text.Trim();
            }
            List<string> errors = id.Validate(marking);
            ShowProblems(errors);
            Result = errors.Count == 0 ? id : null;
        }
    }

    //annexe 38: lowering, raising or lifting the classification
    public class ChangeForm : FrenchDialog
    {
        public enum Change { Declassement, Reclassement, Declassification }

        private readonly FrenchMarking marking;
        private readonly SystemSettings settings;
        private RadioButton lower, raise, lift;
        private DateTimePicker effective, decisionDate;
        private TextBox decision;
        public Change Result;
        public string[] Lines;

        public ChangeForm(FrenchMarking marking, SystemSettings settings)
            : base(Markings.Text.T("Downgrade and declassification (annexe 38)", "Déclassement et déclassification (annexe 38)", "Herabstufung und Aufhebung der Einstufung (annexe 38)"))
        {
            this.marking = marking;
            this.settings = settings;
            int y = 10;
            AddLabel(Markings.Text.T("Document ", "Document ", "Dokument ") + marking.ToString(), 10, y, 480).Font = new Font("Tahoma", 8.25f, FontStyle.Bold);
            y += 24;
            lower = Radio(Markings.Text.T("Downgrade from Très Secret to Secret", "Déclasser de Très Secret à Secret", "Herabstufen von Très Secret auf Secret"), ref y);
            raise = Radio(Markings.Text.T("Upgrade from Secret to Très Secret", "Reclasser de Secret à Très Secret", "Hochstufen von Secret auf Très Secret"), ref y);
            lift = Radio(Markings.Text.T("Declassify", "Déclassifier", "Einstufung aufheben"), ref y);
            lower.Enabled = marking.Level == FrenchLevel.TresSecret;
            raise.Enabled = marking.Level == FrenchLevel.Secret;
            AddLabel(Markings.Text.T("Takes effect on", "Intervient le", "Wirksam ab"), 10, y, 150);
            effective = AddDate(160, y, OnChange);
            y += 26;
            decision = AddText(Markings.Text.T("Decision number", "Décision n°", "Entscheidung Nr."), ref y, OnChange);
            AddLabel(Markings.Text.T("of", "du", "vom"), 10, y, 150);
            decisionDate = AddDate(160, y, OnChange);
            y += 30;
            AddFooter(ref y, Markings.Text.T("Apply", "Appliquer", "Übernehmen"));
            if (lower.Enabled) lower.Checked = true; else if (raise.Enabled) raise.Checked = true; else lift.Checked = true;
            OnChange(null, EventArgs.Empty);
        }

        private RadioButton Radio(string text, ref int y)
        {
            RadioButton radio = new RadioButton();
            radio.Text = text;
            radio.SetBounds(160, y, 330, 20);
            radio.CheckedChanged += OnChange;
            Controls.Add(radio);
            y += 22;
            return radio;
        }

        private void OnChange(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return;
            }
            List<string> errors = new List<string>();
            if (decision.Text.Trim().Length == 0)
            {
                errors.Add(Markings.Text.T("The decision number is required.", "Le numéro de la décision est requis.", "Die Nummer der Entscheidung ist erforderlich."));
            }
            string date = FrenchFormatter.Date(effective.Value);
            string number = decision.Text.Trim();
            string of = FrenchFormatter.Date(decisionDate.Value);
            if (lower.Checked)
            {
                Result = Change.Declassement;
                Lines = new string[] { "Le déclassement du niveau Très Secret", "au niveau Secret", "intervient le " + date, "par décision n° " + number, "du " + of };
            }
            else if (raise.Checked)
            {
                Result = Change.Reclassement;
                Lines = new string[] { "Le reclassement du niveau Secret", "au niveau Très Secret", "intervient le " + date, "par décision n° " + number, "du " + of };
                FrenchMarking raised = marking.Clone();
                raised.Level = FrenchLevel.TresSecret;
                string ceiling = settings.CheckCeiling(raised);
                if (ceiling != null) errors.Add(ceiling);
            }
            else
            {
                Result = Change.Declassification;
                Lines = new string[] { "DÉCLASSIFIÉ", "Le " + date, "par décision n° " + number, "du " + of };
            }
            if (!lower.Checked && !raise.Checked && !lift.Checked)
            {
                errors.Add(Markings.Text.T("Choose an operation.", "Choisir une opération.", "Wählen Sie einen Vorgang aus."));
            }
            ShowProblems(errors);
        }
    }

    //the marking system of a document, on a computer where several are allowed (no Simple value)
    public class SystemForm : FrenchDialog
    {
        public string Result;

        public SystemForm(string current, string[] systems)
            : base(Markings.Text.T("Marking system", "Système de classification", "Kennzeichnungssystem"))
        {
            AddLabel(Markings.Text.T("This computer accepts several systems: which one applies to this document?", "Ce poste accepte plusieurs systèmes : lequel s'applique à ce document ?", "Dieser Rechner unterstützt mehrere Systeme: Welches gilt für dieses Dokument?"), 10, 10, 480);
            int y = 36;
            foreach (string system in systems)
            {
                AddButton(system, Describe(system), y, current);
                y += 34;
            }
            Button cancel = new Button();
            cancel.Text = Markings.Text.T("Cancel", "Annuler", "Abbrechen");
            cancel.SetBounds(415, y + 4, 75, 24);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            CancelButton = cancel;
            ClientSize = new Size(500, y + 38);
        }

        private static string Describe(string system)
        {
            switch (system)
            {
                case DocumentHost.FrenchSystem: return Markings.Text.T("France - IGI 1300 (SECRET, TRÈS SECRET, DIFFUSION RESTREINTE)", "France - IGI 1300 (SECRET, TRÈS SECRET, DIFFUSION RESTREINTE)", "Frankreich - IGI 1300 (SECRET, TRÈS SECRET, DIFFUSION RESTREINTE)");
                case DocumentHost.GermanSystem: return Markings.Text.T("Germany - VS-Anweisung (VS-NfD, VS-VERTRAULICH, GEHEIM, STRENG GEHEIM)", "Allemagne - VS-Anweisung (VS-NfD, VS-VERTRAULICH, GEHEIM, STRENG GEHEIM)", "Deutschland - VS-Anweisung (VS-NfD, VS-VERTRAULICH, GEHEIM, STRENG GEHEIM)");
                case DocumentHost.BritishSystem: return Markings.Text.T("United Kingdom - GSCP (OFFICIAL, OFFICIAL-SENSITIVE, SECRET, TOP SECRET)", "Royaume-Uni - GSCP (OFFICIAL, OFFICIAL-SENSITIVE, SECRET, TOP SECRET)", "Vereinigtes Königreich - GSCP (OFFICIAL, OFFICIAL-SENSITIVE, SECRET, TOP SECRET)");
                default: return Markings.Text.T("United States - IC Markings System (TOP SECRET//SI//NOFORN...)", "États-Unis - IC Markings System (TOP SECRET//SI//NOFORN...)", "Vereinigte Staaten - IC Markings System (TOP SECRET//SI//NOFORN...)");
            }
        }

        private void AddButton(string system, string text, int y, string current)
        {
            Button button = new Button();
            button.Text = text;
            button.SetBounds(10, y, 480, 28);
            if (system == current)
            {
                button.Font = new Font(Font, FontStyle.Bold);
                AcceptButton = button;
            }
            button.Click += delegate { Result = system; DialogResult = DialogResult.OK; Close(); };
            Controls.Add(button);
        }
    }
}
