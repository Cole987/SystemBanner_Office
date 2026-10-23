//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SystemBannerOffice.Markings
{
    //German classified information, Allgemeine Verwaltungsvorschrift zum materiellen Geheimschutz
    //(Verschlusssachenanweisung, VSA) of the federal government, 13 March 2023: § 16 (duration), § 20 (marking), Anlage IV
    public enum GermanLevel
    {
        Offen = 0,           //not classified; "[offen]" is a paragraph mark of this project, the VSA has none
        VsNfd = 1,           //VS-NUR FÜR DEN DIENSTGEBRAUCH
        VsVertraulich = 2,   //VS-VERTRAULICH
        Geheim = 3,          //GEHEIM
        StrengGeheim = 4     //STRENG GEHEIM
    }

    public class GermanMarking
    {
        public GermanLevel Level = GermanLevel.Offen;
        public List<string> Vermerke = new List<string>(); //Warn- und Sperrvermerke of Anlage IV § 1.3, applied by the issuer only (§ 20 Abs. 3)

        public bool IsClassified
        {
            get { return Level > GermanLevel.Offen; }
        }

        //"amtlich geheim gehalten" from VS-VERTRAULICH up (Anlage IV § 1.1)
        public bool AmtlichGeheimGehalten
        {
            get { return Level >= GermanLevel.VsVertraulich; }
        }

        public GermanMarking Clone()
        {
            GermanMarking clone = (GermanMarking)MemberwiseClone();
            clone.Vermerke = new List<string>(Vermerke);
            return clone;
        }

        public override string ToString()
        {
            return GermanFormatter.Stamp(Level);
        }
    }

    //what the first page carries, § 20 Abs. 2
    public class GermanIdentification
    {
        public string Herausgeber;          //Nr. 2
        public DateTime Datum = DateTime.Today; //Nr. 3
        public int EinstufungEndet;         //Nr. 4 and 5: the year the classification ends with; 0 for VS-NfD with the default 30 years
        public string Kennzeichen;          //Nr. 6, from VS-VERTRAULICH up: Geschäftszeichen and Tagebuchnummer
        public int Ausfertigung;            //Nr. 7, from VS-VERTRAULICH up: copy number
        public string Empfaenger;           //Nr. 7: the recipient of this copy

        public const int MaximumYears = 30; //§ 16

        public List<string> Validate(GermanMarking marking)
        {
            List<string> errors = new List<string>();
            if (!marking.IsClassified)
            {
                return errors;
            }
            if (string.IsNullOrEmpty(Herausgeber))
            {
                errors.Add(Text.T("The issuer (Herausgeber) is marked on the document (VSA § 20 Abs. 2 Nr. 2).", "L'émetteur (Herausgeber) figure sur le document (VSA § 20 al. 2 n° 2)."));
            }
            int lastYear = Datum.Year + MaximumYears;
            if (marking.Level == GermanLevel.VsNfd)
            {
                //VS-NfD: 30 years unless the issuer sets a shorter time, never longer (§ 16 Abs. 1, § 17 Abs. 1)
                if (EinstufungEndet != 0 && (EinstufungEndet < Datum.Year || EinstufungEndet > lastYear))
                {
                    errors.Add(Text.T("A VS-NfD classification ends within 30 years (VSA § 16 Abs. 1): " + Datum.Year + " to " + lastYear + ".", "Une classification VS-NfD prend fin dans les 30 ans (VSA § 16 al. 1) : de " + Datum.Year + " à " + lastYear + "."));
                }
                return errors;
            }
            if (EinstufungEndet == 0)
            {
                errors.Add(Text.T("From VS-VERTRAULICH up, the issuer sets when the classification ends: \"Die Einstufung endet mit Ablauf des Jahres …\" (VSA § 20 Abs. 2 Nr. 5).", "À partir de VS-VERTRAULICH, l'émetteur fixe la fin de la classification : « Die Einstufung endet mit Ablauf des Jahres … » (VSA § 20 al. 2 n° 5)."));
            }
            else if (EinstufungEndet < Datum.Year || EinstufungEndet > lastYear)
            {
                errors.Add(Text.T("The classification should not last more than 30 years (VSA § 16 Abs. 2): " + Datum.Year + " to " + lastYear + "; a longer time is an extension under § 17.", "La classification ne devrait pas dépasser 30 ans (VSA § 16 al. 2) : de " + Datum.Year + " à " + lastYear + " ; au-delà, c'est une prolongation selon le § 17."));
            }
            if (string.IsNullOrEmpty(Kennzeichen))
            {
                errors.Add(Text.T("From VS-VERTRAULICH up, the document carries an individual reference, such as the Geschäftszeichen and Tagebuchnummer (VSA § 20 Abs. 2 Nr. 6).", "À partir de VS-VERTRAULICH, le document porte une référence individuelle, par exemple le Geschäftszeichen et la Tagebuchnummer (VSA § 20 al. 2 n° 6)."));
            }
            if (Ausfertigung <= 0 || string.IsNullOrEmpty(Empfaenger))
            {
                errors.Add(Text.T("From VS-VERTRAULICH up, each copy carries its number and its recipient (VSA § 20 Abs. 2 Nr. 7).", "À partir de VS-VERTRAULICH, chaque exemplaire porte son numéro et son destinataire (VSA § 20 al. 2 n° 7)."));
            }
            return errors;
        }

        //the first page lines, in German: the document content is German whatever the language of the add-in
        public string[] Lines(GermanMarking marking)
        {
            List<string> lines = new List<string>();
            if (!string.IsNullOrEmpty(Herausgeber)) lines.Add("Herausgeber: " + Herausgeber);
            lines.Add("Datum: " + Datum.ToString("dd.MM.yyyy"));
            if (marking.AmtlichGeheimGehalten && !string.IsNullOrEmpty(Kennzeichen))
            {
                lines.Add(Kennzeichen + " " + GermanFormatter.Abbreviation(marking.Level)); //"supplemented by the abbreviation of the level", § 20 Abs. 2 Nr. 6
            }
            if (marking.AmtlichGeheimGehalten && Ausfertigung > 0)
            {
                lines.Add(Ausfertigung + ". Ausfertigung" + (string.IsNullOrEmpty(Empfaenger) ? "" : ", Empfänger: " + Empfaenger));
            }
            if (EinstufungEndet > 0)
            {
                lines.Add("Die Einstufung endet mit Ablauf des Jahres " + EinstufungEndet + ".");
            }
            lines.Add("{NUMPAGES} Seiten");
            return lines.ToArray();
        }
    }

    public static class GermanFormatter
    {
        public const string AmtlichGeheimGehalten = "amtlich geheim gehalten";

        //written out in full, § 20 Abs. 4
        public static string Stamp(GermanLevel level)
        {
            switch (level)
            {
                case GermanLevel.StrengGeheim: return "STRENG GEHEIM";
                case GermanLevel.Geheim: return "GEHEIM";
                case GermanLevel.VsVertraulich: return "VS-VERTRAULICH";
                case GermanLevel.VsNfd: return "VS-NUR FÜR DEN DIENSTGEBRAUCH";
                default: return "";
            }
        }

        //the abbreviations of § 20 Abs. 4
        public static string Abbreviation(GermanLevel level)
        {
            switch (level)
            {
                case GermanLevel.StrengGeheim: return "Str. Geh.";
                case GermanLevel.Geheim: return "Geh.";
                case GermanLevel.VsVertraulich: return "VS-Vertr.";
                case GermanLevel.VsNfd: return "VS-NfD";
                default: return "offen";
            }
        }

        //Anlage IV § 1.1: STRENG GEHEIM and GEHEIM at the top and bottom, VS-VERTRAULICH and VS-NfD at the top only
        public static bool AtBottom(GermanLevel level)
        {
            return level >= GermanLevel.Geheim;
        }

        //red for STRENG GEHEIM and GEHEIM; black or blue for the others, blue for VS-VERTRAULICH like its folder stripe, black for VS-NfD
        public static int[] Color(GermanLevel level)
        {
            if (level >= GermanLevel.Geheim) return new int[] { 200, 16, 46 };
            if (level == GermanLevel.VsVertraulich) return new int[] { 0, 51, 160 };
            return new int[] { 0, 0, 0 };
        }

        //the paragraph mark of a part of the document, § 20 Abs. 6: "[Geh.]", "[VS-NfD]"; the form is a choice of this project
        public static string Portion(GermanLevel level)
        {
            return "[" + Abbreviation(level) + "]";
        }
    }

    public static class GermanParser
    {
        private static readonly Regex PortionMark = new Regex(@"^\[(Str\. Geh\.|Geh\.|VS-Vertr\.|VS-NfD|offen)\]$");

        public static GermanMarking ParsePortion(string text)
        {
            Match match = PortionMark.Match((text ?? "").Trim());
            if (!match.Success)
            {
                return null;
            }
            GermanMarking m = new GermanMarking();
            foreach (GermanLevel level in new GermanLevel[] { GermanLevel.StrengGeheim, GermanLevel.Geheim, GermanLevel.VsVertraulich, GermanLevel.VsNfd, GermanLevel.Offen })
            {
                if (GermanFormatter.Abbreviation(level) == match.Groups[1].Value)
                {
                    m.Level = level;
                }
            }
            return m;
        }
    }

    public static class GermanRollUp
    {
        //a document with parts of different levels is marked with the highest one (§ 20 Abs. 6)
        public static GermanMarking Compute(IList<GermanMarking> parts)
        {
            GermanMarking document = new GermanMarking();
            foreach (GermanMarking part in parts)
            {
                if (part.Level > document.Level)
                {
                    document.Level = part.Level;
                }
                foreach (string vermerk in part.Vermerke)
                {
                    if (!document.Vermerke.Contains(vermerk)) document.Vermerke.Add(vermerk);
                }
            }
            return document;
        }
    }

    //the Warn- und Sperrvermerke of Anlage IV § 1.3; "Weitergabe an …" and "… intern" take a name
    public static class GermanVermerke
    {
        public static readonly string[] Warnvermerke = new string[] { "KRYPTOSICHERHEIT", "Controlled COMSEC Item", "Atomal", "Auswertesache" };
        public static readonly string[] Sperrvermerke = new string[] { "Nur Deutschen zur Kenntnis", "Nur zu ihrer Information", "Die Fertigung von Kopien dieser VS ist untersagt", "Für Geheimschutzbeauftragte" };
    }

    public static class GermanSettings
    {
        public const string Germany = "Germany";

        //German markings: no Simple value, or a German one (21 to 24)
        public static bool Available(SystemSettings settings)
        {
            return settings.Ceiling == null || settings.Ceiling.ClassificationSystem == Germany;
        }

        public static string CheckCeiling(SystemSettings settings, GermanMarking marking)
        {
            if (!Available(settings))
            {
                return Text.T("This computer is set to the " + settings.Ceiling.ClassificationSystem + " system (" + settings.Ceiling.Text + "): German markings are not offered.",
                    "Ce poste est réglé sur le système " + settings.Ceiling.ClassificationSystem + " (" + settings.Ceiling.Text + ") : les marquages allemands n'y sont pas proposés.");
            }
            //presets 21 to 24: VS-NfD rank 0 to STRENG GEHEIM rank 3
            if (settings.Ceiling != null && marking.IsClassified && (int)marking.Level - 1 > settings.Ceiling.Rank)
            {
                return Text.T("This computer is set to " + settings.Ceiling.Text + ": a " + marking + " document cannot be written on it.",
                    "Ce poste est réglé au niveau " + settings.Ceiling.Text + " : un document " + marking + " ne peut pas y être rédigé.");
            }
            return null;
        }
    }
}
