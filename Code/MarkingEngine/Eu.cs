//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SystemBannerOffice.Markings
{
    //EU classified information (EUCI): Council Decision 2013/488/EU (CSR) and Commission Decision (EU, Euratom) 2015/444 (CSD),
    //the same marking rules in both; see Docs\EU-marking-rules.md for every rule and what the texts leave open
    public enum EuLevel
    {
        Restreint = 0,
        Confidentiel = 1,
        Secret = 2,
        TresSecret = 3
    }

    public class EuMarking
    {
        public EuLevel Level = EuLevel.Restreint;
        //additional markings (CSR Annex III pt 10, CSD Art. 23): originator, caveats, releasability, downgrading date or event;
        //the texts give no format, each line is written as typed under the classification at the top of the pages
        public List<string> Additional = new List<string>();
        public string Reference;   //reference (CSR) or registration (CSD) number, every level: CSR Annex III pt 13(c)
        public int Copy;           //copy number on every page, SECRET UE/EU SECRET and above when several copies: CSR Annex III pt 13(e)

        public EuMarking Clone()
        {
            EuMarking clone = (EuMarking)MemberwiseClone();
            clone.Additional = new List<string>(Additional);
            return clone;
        }

        public override string ToString()
        {
            return EuFormatter.Marking(Level);
        }
    }

    public static class EuFormatter
    {
        //CSR Art. 2(2), CSD Art. 3(2): the bilingual string is the marking itself, French part first, in every language
        public static string Marking(EuLevel level)
        {
            switch (level)
            {
                case EuLevel.TresSecret: return "TRÈS SECRET UE/EU TOP SECRET";
                case EuLevel.Secret: return "SECRET UE/EU SECRET";
                case EuLevel.Confidentiel: return "CONFIDENTIEL UE/EU CONFIDENTIAL";
                default: return "RESTREINT UE/EU RESTRICTED";
            }
        }

        //CSR Annex III pt 12, CSD Art. 24(2); the brackets are a choice of this project, as for the French and German marks
        public static string Abbreviation(EuLevel level)
        {
            switch (level)
            {
                case EuLevel.TresSecret: return "TS-UE/EU-TS";
                case EuLevel.Secret: return "S-UE/EU-S";
                case EuLevel.Confidentiel: return "C-UE/EU-C";
                default: return "R-UE/EU-R";
            }
        }

        public static string Portion(EuLevel level)
        {
            return "[" + Abbreviation(level) + "]";
        }

        //what goes after the page number on every page: the reference, and the copy number from SECRET UE/EU SECRET up
        public static string Footer(EuMarking m)
        {
            List<string> parts = new List<string>();
            if (!string.IsNullOrEmpty(m.Reference)) parts.Add(m.Reference);
            if (m.Copy > 0 && m.Level >= EuLevel.Secret) parts.Add("Ex./Copy " + m.Copy);
            return string.Join(" - ", parts.ToArray());
        }
    }

    public static class EuParser
    {
        //"[R-UE/EU-R]" or "(R-UE/EU-R)" at the start of a paragraph
        public static readonly Regex PortionStart = new Regex(@"^\s*[\[(]((?:TS|S|C|R)-UE/EU-(?:TS|S|C|R))[\])]");

        public static EuMarking ParsePortion(string mark)
        {
            Match match = PortionStart.Match(mark ?? "");
            if (!match.Success) return null;
            foreach (EuLevel level in new EuLevel[] { EuLevel.Restreint, EuLevel.Confidentiel, EuLevel.Secret, EuLevel.TresSecret })
            {
                if (EuFormatter.Abbreviation(level) == match.Groups[1].Value)
                {
                    EuMarking m = new EuMarking();
                    m.Level = level;
                    return m;
                }
            }
            return null; //"S-UE/EU-TS": the two halves do not match
        }

        //the level written in full, with or without the accent of TRÈS (the English text of the CSD prints TRES)
        public static EuMarking ParseMarking(string text)
        {
            string normalized = (text ?? "").Trim().ToUpperInvariant().Replace("TRES ", "TRÈS ");
            foreach (EuLevel level in new EuLevel[] { EuLevel.Restreint, EuLevel.Confidentiel, EuLevel.Secret, EuLevel.TresSecret })
            {
                if (EuFormatter.Marking(level) == normalized)
                {
                    EuMarking m = new EuMarking();
                    m.Level = level;
                    return m;
                }
            }
            return null;
        }
    }

    public static class EuRollUp
    {
        //the document is classified at least as high as its most highly classified part: CSR Annex III pt 7, CSD Art. 22(6)
        public static EuMarking Compute(List<EuMarking> parts)
        {
            EuMarking document = new EuMarking();
            foreach (EuMarking part in parts)
            {
                if (part.Level > document.Level) document.Level = part.Level;
            }
            return document;
        }
    }

    public static class EuValidator
    {
        public static List<string> Validate(EuMarking m)
        {
            List<string> errors = new List<string>();
            if (string.IsNullOrEmpty(m.Reference))
            {
                errors.Add(Text.T("Every EU classified document carries a reference or registration number (CSR Annex III pt 13(c), CSD Art. 25(1)(c)).",
                    "Tout document classifié UE porte un numéro de référence ou d'enregistrement (CSR annexe III pt 13 c), CSD art. 25, par. 1, c)).",
                    "Jedes EU-Verschlussdokument trägt ein Aktenzeichen oder eine Registriernummer (CSR Anhang III Nr. 13 Buchst. c, CSD Art. 25 Abs. 1 Buchst. c)."));
            }
            if (m.Copy < 0)
            {
                errors.Add(Text.T("The copy number is a positive number.", "Le numéro d'exemplaire est un nombre positif.", "Die Ausfertigungsnummer ist eine positive Zahl."));
            }
            return errors;
        }
    }

    public static class EuSettings
    {
        public const string EuropeanUnion = "European Union";

        //EU markings: no Simple value, or an EU one (61 to 64)
        public static bool Available(SystemSettings settings)
        {
            return settings.Ceiling == null || settings.Ceiling.ClassificationSystem == EuropeanUnion;
        }

        public static string CheckCeiling(SystemSettings settings, EuMarking marking)
        {
            if (!Available(settings))
            {
                return Text.T("This computer is set to the " + settings.Ceiling.ClassificationSystem + " system (" + settings.Ceiling.Text + "): EU markings are not offered.",
                    "Ce poste est réglé sur le système " + settings.Ceiling.ClassificationSystem + " (" + settings.Ceiling.Text + ") : les marquages de l'UE n'y sont pas proposés.",
                    "Dieser Rechner ist auf das System " + settings.Ceiling.ClassificationSystem + " eingestellt (" + settings.Ceiling.Text + "): EU-Kennzeichnungen werden nicht angeboten.");
            }
            //presets 61 to 64: RESTREINT rank 0 to TRÈS SECRET rank 3
            if (settings.Ceiling != null && (int)marking.Level > settings.Ceiling.Rank)
            {
                return Text.T("This computer is set to " + settings.Ceiling.Text + ": a " + marking + " document cannot be written on it.",
                    "Ce poste est réglé au niveau " + settings.Ceiling.Text + " : un document " + marking + " ne peut pas y être rédigé.",
                    "Dieser Rechner ist auf " + settings.Ceiling.Text + " eingestellt: Ein Dokument " + marking + " kann hier nicht erstellt werden.");
            }
            return null;
        }
    }
}
