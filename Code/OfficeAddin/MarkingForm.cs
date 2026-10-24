//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice
{
    //picks a classification and its control markings, with the banner line and portion mark shown as they are typed
    //and the reasons a combination is refused; built in code, .NET 2.0 Windows Forms, no designer file
    internal static class ItemCheckUpdate
    {
        //ItemCheck comes before the box changes state: the update runs right after the click. Before the window exists
        //(a dialog filled from the document's marking), BeginInvoke would fail: the update runs when the dialog opens.
        public static void After(Form form, EventHandler update)
        {
            if (form.IsHandleCreated)
            {
                form.BeginInvoke(update, form, EventArgs.Empty);
                return;
            }
            EventHandler once = null;
            once = delegate(object sender, EventArgs e)
            {
                form.Load -= once;
                update(sender, e);
            };
            form.Load += once;
        }
    }

    public class MarkingForm : Form
    {
        private readonly bool portionMode;
        private readonly SystemSettings settings;
        private readonly Marking documentBanner;

        private ComboBox owner, level;
        private TextBox ownerCodes, sci, sap, sigma, fgi, relTo, eyes, display;
        private CheckBox rd, cnwdi, frd, dodUcni, doeUcni, tfni, fgiBox, atomal, bohemia, balk;
        private CheckedListBox dissem, nonIc;
        private Label banner, portion;
        private ListBox problems;
        private Button ok;

        public Marking Result;

        public MarkingForm(bool portionMode, Marking initial, Marking documentBanner, SystemSettings settings)
        {
            this.portionMode = portionMode;
            this.settings = settings;
            this.documentBanner = documentBanner;
            Text = "SystemBanner - " + (portionMode ? Markings.Text.T("Portion mark", "Marque de portion") : Markings.Text.T("Classify document", "Classifier le document"));
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(640, 610);
            Font = new Font("Tahoma", 8.25f);
            Build();
            if (initial != null)
            {
                LoadMarking(initial);
            }
            UpdatePreview(null, EventArgs.Empty);
        }

        private int y = 10;

        private Label AddLabel(string text, int x, int top, int width)
        {
            Label label = new Label();
            label.Text = text;
            label.SetBounds(x, top + 3, width, 16);
            Controls.Add(label);
            return label;
        }

        private TextBox AddText(string label, string hint)
        {
            AddLabel(label, 10, y, 130);
            TextBox box = new TextBox();
            box.SetBounds(140, y, 230, 20);
            box.TextChanged += UpdatePreview;
            Controls.Add(box);
            AddLabel(hint, 375, y, 255).ForeColor = SystemColors.GrayText;
            y += 26;
            return box;
        }

        private CheckBox AddCheck(string text, int x, int top, int width)
        {
            CheckBox box = new CheckBox();
            box.Text = text;
            box.SetBounds(x, top, width, 20);
            box.CheckedChanged += UpdatePreview;
            Controls.Add(box);
            return box;
        }

        private void Build()
        {
            AddLabel(Markings.Text.T("Classification", "Classification"), 10, y, 130);
            owner = new ComboBox();
            owner.DropDownStyle = ComboBoxStyle.DropDownList;
            owner.Items.AddRange(new object[] { "US", "Non-US (//GBR ...)", "NATO", "JOINT" });
            owner.SetBounds(140, y, 120, 20);
            owner.SelectedIndex = 0;
            owner.SelectedIndexChanged += UpdatePreview;
            Controls.Add(owner);
            level = new ComboBox();
            level.DropDownStyle = ComboBoxStyle.DropDownList;
            level.Items.AddRange(new object[] { "TOP SECRET", "SECRET", "CONFIDENTIAL", "RESTRICTED", "UNCLASSIFIED" });
            level.SetBounds(265, y, 105, 20);
            level.SelectedIndex = 4;
            level.SelectedIndexChanged += UpdatePreview;
            Controls.Add(level);
            y += 26;
            ownerCodes = AddText(Markings.Text.T("Owners / NAC activity", "Propriétaires / activité NAC"), Markings.Text.T("non-US: GBR, or CAN GBR; JOINT: CAN GBR USA; NATO: ISAF", "non US : GBR, ou CAN GBR ; JOINT : CAN GBR USA ; OTAN : ISAF"));
            sci = AddText("SCI", "SI-G ABCD/TK-BLFH/123");
            sap = AddText("SAP", Markings.Text.T("BP-123/CA-XYZ (without SAR-)", "BP-123/CA-XYZ (sans SAR-)"));

            AddLabel("AEA", 10, y, 130);
            rd = AddCheck("RD", 140, y, 45);
            cnwdi = AddCheck("CNWDI", 185, y, 65);
            frd = AddCheck("FRD", 250, y, 50);
            tfni = AddCheck("TFNI", 300, y, 55);
            dodUcni = AddCheck("DOD UCNI", 355, y, 80);
            doeUcni = AddCheck("DOE UCNI", 435, y, 80);
            atomal = AddCheck("ATOMAL", 515, y, 70);
            y += 24;
            sigma = AddText("SIGMA", Markings.Text.T("14 18 (with RD or FRD)", "14 18 (avec RD ou FRD)"));
            AddLabel("NATO SAP", 10, y, 130);
            bohemia = AddCheck("BOHEMIA", 140, y, 80);
            balk = AddCheck("BALK", 220, y, 60);
            fgiBox = AddCheck(Markings.Text.T("FGI in US information", "FGI dans une information US"), 300, y, 180);
            y += 24;
            fgi = AddText(Markings.Text.T("FGI sources", "Sources FGI"), Markings.Text.T("DEU GBR, empty to conceal the source", "DEU GBR, vide pour masquer la source"));

            AddLabel(Markings.Text.T("Dissemination", "Diffusion"), 10, y, 130);
            dissem = new CheckedListBox();
            dissem.CheckOnClick = true;
            foreach (RegisterEntry entry in Register.Dissemination)
            {
                dissem.Items.Add(entry.BannerText(true) + (entry.Abbreviation == null ? "" : "  (" + entry.Title + ")"));
            }
            dissem.SetBounds(140, y, 230, 120);
            dissem.ItemCheck += delegate { ItemCheckUpdate.After(this, UpdatePreview); };
            Controls.Add(dissem);
            AddLabel("Non-IC", 375, y, 60);
            nonIc = new CheckedListBox();
            nonIc.CheckOnClick = true;
            foreach (RegisterEntry entry in Register.NonIcDissemination)
            {
                nonIc.Items.Add(entry.BannerText(true));
            }
            nonIc.SetBounds(435, y, 195, 120);
            nonIc.ItemCheck += delegate { ItemCheckUpdate.After(this, UpdatePreview); };
            Controls.Add(nonIc);
            y += 126;
            relTo = AddText("REL TO", Markings.Text.T("USA, AUS, GBR, NATO (with REL TO checked)", "USA, AUS, GBR, NATO (avec REL TO coché)"));
            eyes = AddText("EYES ONLY", Markings.Text.T("USA/CAN/GBR (NSA SIGINT only)", "USA/CAN/GBR (SIGINT de la NSA uniquement)"));
            display = AddText("DISPLAY ONLY", "AFG, IRQ");

            AddLabel(Markings.Text.T("Banner line", "Ligne de bannière"), 10, y, 130);
            banner = new Label();
            banner.SetBounds(140, y, 490, 34);
            banner.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            banner.TextAlign = ContentAlignment.MiddleCenter;
            Controls.Add(banner);
            y += 38;
            AddLabel(Markings.Text.T("Portion mark", "Marque de portion"), 10, y, 130);
            portion = new Label();
            portion.SetBounds(140, y, 490, 18);
            portion.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            Controls.Add(portion);
            y += 22;
            problems = new ListBox();
            problems.SetBounds(10, y, 620, 70);
            problems.ForeColor = Color.DarkRed;
            problems.HorizontalScrollbar = true;
            Controls.Add(problems);
            y += 76;
            ok = new Button();
            ok.Text = portionMode ? Markings.Text.T("Insert", "Insérer") : Markings.Text.T("Classify", "Classifier");
            ok.SetBounds(470, y, 75, 24);
            ok.Click += delegate { DialogResult = DialogResult.OK; Close(); };
            Controls.Add(ok);
            Button cancel = new Button();
            cancel.Text = Markings.Text.T("Cancel", "Annuler");
            cancel.SetBounds(555, y, 75, 24);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            ClientSize = new Size(640, y + 34);
        }

        private void LoadMarking(Marking m)
        {
            owner.SelectedIndex = (int)m.Owner == 0 ? 0 : (m.Owner == Markings.Owner.NonUS ? 1 : (m.Owner == Markings.Owner.Nato ? 2 : 3));
            level.SelectedIndex = 4 - (int)m.Level;
            ownerCodes.Text = m.Owner == Markings.Owner.Joint ? string.Join(" ", m.JointOwners.ToArray()) : (m.Owner == Markings.Owner.Nato ? m.NatoActivity : m.OwnerCode) ?? "";
            Marking sciOnly = new Marking();
            sciOnly.Level = Level.TopSecret;
            sciOnly.Sci = m.Sci;
            sci.Text = Category(MarkingFormatter.Portion(sciOnly), 1);
            Marking sapOnly = new Marking();
            sapOnly.Level = Level.TopSecret;
            sapOnly.Sap = m.Sap;
            string sapText = Category(MarkingFormatter.Portion(sapOnly), 1);
            sap.Text = sapText.StartsWith("SAR-") ? sapText.Substring(4) : sapText;
            rd.Checked = m.RestrictedData;
            cnwdi.Checked = m.Cnwdi;
            frd.Checked = m.FormerlyRestrictedData;
            tfni.Checked = m.Tfni;
            dodUcni.Checked = m.DodUcni;
            doeUcni.Checked = m.DoeUcni;
            List<string> sigmas = new List<string>();
            foreach (int n in m.RdSigma) sigmas.Add(n.ToString());
            foreach (int n in m.FrdSigma) sigmas.Add(n.ToString());
            sigma.Text = string.Join(" ", sigmas.ToArray());
            atomal.Checked = m.NatoSap.Contains("ATOMAL");
            bohemia.Checked = m.NatoSap.Contains("BOHEMIA");
            balk.Checked = m.NatoSap.Contains("BALK");
            fgiBox.Checked = m.Fgi;
            fgi.Text = string.Join(" ", m.FgiSources.ToArray());
            for (int i = 0; i < Register.Dissemination.Length; i++)
            {
                dissem.SetItemChecked(i, m.HasDissem(Register.Dissemination[i].Key));
            }
            for (int i = 0; i < Register.NonIcDissemination.Length; i++)
            {
                nonIc.SetItemChecked(i, m.HasNonIcDissem(Register.NonIcDissemination[i].Key));
            }
            relTo.Text = string.Join(", ", m.RelTo.ToArray());
            eyes.Text = string.Join("/", m.EyesOnly.ToArray());
            display.Text = string.Join(", ", m.DisplayOnly.ToArray());
        }

        private static string Category(string marking, int index)
        {
            string[] parts = marking.Split(new string[] { "//" }, StringSplitOptions.None);
            return parts.Length > index ? parts[index] : "";
        }

        private static List<string> Codes(string text, params char[] separators)
        {
            List<string> codes = new List<string>();
            foreach (string code in text.ToUpperInvariant().Split(separators, StringSplitOptions.RemoveEmptyEntries))
            {
                if (code.Trim().Length > 0) codes.Add(code.Trim());
            }
            return codes;
        }

        //builds the marking from the fields; parse errors of the free text fields go to errors
        private Marking Build(List<string> errors)
        {
            Marking m = new Marking();
            m.Owner = owner.SelectedIndex == 1 ? Markings.Owner.NonUS : owner.SelectedIndex == 2 ? Markings.Owner.Nato : owner.SelectedIndex == 3 ? Markings.Owner.Joint : Markings.Owner.UnitedStates;
            m.Level = (Level)(4 - level.SelectedIndex);
            string codes = ownerCodes.Text.Trim().ToUpperInvariant();
            if (m.Owner == Markings.Owner.NonUS) m.OwnerCode = codes.Length == 0 ? null : string.Join(" ", Codes(codes, ' ', ',').ToArray());
            if (m.Owner == Markings.Owner.Joint) m.JointOwners = Codes(codes, ' ', ',');
            if (m.Owner == Markings.Owner.Nato && codes.Length > 0) m.NatoActivity = codes;
            if (sci.Text.Trim().Length > 0) m.Sci = ParseCategory("TS//" + sci.Text.Trim().ToUpperInvariant(), errors).Sci;
            if (sap.Text.Trim().Length > 0) m.Sap = ParseCategory("TS//SAR-" + sap.Text.Trim().ToUpperInvariant(), errors).Sap;
            m.RestrictedData = rd.Checked;
            m.Cnwdi = cnwdi.Checked;
            m.FormerlyRestrictedData = frd.Checked;
            m.Tfni = tfni.Checked;
            m.DodUcni = dodUcni.Checked;
            m.DoeUcni = doeUcni.Checked;
            foreach (string number in Codes(sigma.Text, ' ', ','))
            {
                int n;
                if (int.TryParse(number, out n))
                {
                    if (m.RestrictedData) m.RdSigma.Add(n); else m.FrdSigma.Add(n);
                }
                else
                {
                    errors.Add(Markings.Text.T("SIGMA numbers are numbers: \"" + number + "\".", "Les numéros SIGMA sont des nombres : « " + number + " »."));
                }
            }
            if (atomal.Checked) m.NatoSap.Add("ATOMAL");
            if (balk.Checked) m.NatoSap.Add("BALK");
            if (bohemia.Checked) m.NatoSap.Add("BOHEMIA");
            m.Fgi = fgiBox.Checked;
            if (m.Fgi) m.FgiSources = Codes(fgi.Text, ' ', ',');
            for (int i = 0; i < Register.Dissemination.Length; i++)
            {
                if (dissem.GetItemChecked(i)) m.Dissem.Add(Register.Dissemination[i].Key);
            }
            for (int i = 0; i < Register.NonIcDissemination.Length; i++)
            {
                if (nonIc.GetItemChecked(i)) m.NonIcDissem.Add(Register.NonIcDissemination[i].Key);
            }
            if (m.HasDissem("REL"))
            {
                m.RelTo = Codes(relTo.Text, ',', ' ');
                if (m.RelTo.Count > 0 && !m.RelTo.Contains("USA")) m.RelTo.Insert(0, "USA");
            }
            if (m.HasDissem("EYES"))
            {
                m.EyesOnly = Codes(eyes.Text, '/', ',', ' ');
                if (!m.EyesOnly.Contains("USA")) m.EyesOnly.Insert(0, "USA");
            }
            if (m.HasDissem("DISPLAYONLY")) m.DisplayOnly = Codes(display.Text, ',', ' ');
            return m;
        }

        private static Marking ParseCategory(string text, List<string> errors)
        {
            ParseResult result = MarkingParser.Parse(text);
            errors.AddRange(result.Errors);
            return result.Marking;
        }

        private void UpdatePreview(object sender, EventArgs e)
        {
            if (problems == null)
            {
                return; //still building
            }
            List<string> errors = new List<string>();
            Marking m = Build(errors);
            errors.AddRange(MarkingValidator.Validate(m));
            if (m.HasDissem("REL") && m.RelTo.Count == 0) errors.Add(Markings.Text.T("REL TO needs its countries.", "REL TO a besoin de ses pays."));
            if (m.HasDissem("DISPLAYONLY") && m.DisplayOnly.Count == 0) errors.Add(Markings.Text.T("DISPLAY ONLY needs its countries.", "DISPLAY ONLY a besoin de ses pays."));
            if (m.Owner == Markings.Owner.NonUS && m.OwnerCode == null && !portionMode) errors.Add(Markings.Text.T("A non-US document names its owner country.", "Un document non américain nomme son pays propriétaire."));
            if (m.Owner == Markings.Owner.Joint && m.JointOwners.Count < 2) errors.Add(Markings.Text.T("JOINT information names its co-owners, USA among them.", "Une information JOINT nomme ses copropriétaires, dont USA."));
            if (!portionMode && settings != null)
            {
                string ceiling = settings.CheckCeiling(m);
                if (ceiling != null) errors.Add(ceiling);
            }
            banner.Text = MarkingFormatter.Banner(m);
            int[] colors = Presets.ColorsFor(m);
            banner.BackColor = Color.FromArgb(colors[0], colors[1], colors[2]);
            banner.ForeColor = Color.FromArgb(colors[3], colors[4], colors[5]);
            portion.Text = "(" + MarkingFormatter.Portion(m, documentBanner) + ")";
            problems.Items.Clear();
            foreach (string error in errors)
            {
                problems.Items.Add(error);
            }
            ok.Enabled = errors.Count == 0;
            Result = errors.Count == 0 ? m : null;
        }
    }

    //the stock levels of SystemBanner for the French, German, British, NATO and UN systems, up to the computer's level
    public class NationalForm : Form
    {
        public Preset Result;

        public NationalForm(SystemSettings settings, string current)
        {
            Text = "SystemBanner - " + settings.Ceiling.ClassificationSystem;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(320, 100 + 26 * Presets.LevelsOf(settings.Ceiling.ClassificationSystem, settings.Ceiling).Length);
            Font = new Font("Tahoma", 8.25f);
            int y = 10;
            foreach (Preset preset in Presets.LevelsOf(settings.Ceiling.ClassificationSystem, settings.Ceiling))
            {
                Button button = new Button();
                button.Text = preset.Text;
                button.SetBounds(10, y, 300, 22);
                button.BackColor = Color.FromArgb(preset.Colors[0], preset.Colors[1], preset.Colors[2]);
                button.ForeColor = Color.FromArgb(preset.Colors[3], preset.Colors[4], preset.Colors[5]);
                button.Font = new Font(Font, preset.Text == current ? FontStyle.Bold | FontStyle.Underline : FontStyle.Bold);
                Preset chosen = preset;
                button.Click += delegate { Result = chosen; DialogResult = DialogResult.OK; Close(); };
                Controls.Add(button);
                y += 26;
            }
            Label note = new Label();
            note.Text = Markings.Text.T("Levels above the classification of this computer (" + settings.Ceiling.Text + ") are not offered.", "Les niveaux supérieurs à la classification de ce poste (" + settings.Ceiling.Text + ") ne sont pas proposés.");
            note.SetBounds(10, y + 4, 300, 30);
            Controls.Add(note);
            Button cancel = new Button();
            cancel.Text = Markings.Text.T("Cancel", "Annuler");
            cancel.SetBounds(235, y + 40, 75, 24);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            CancelButton = cancel;
        }
    }
}
