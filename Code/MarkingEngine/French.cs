//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SystemBannerOffice.Markings
{
    //French protection of national defence secrets, IGI 1300/SGDSN/PSE/PSD of 9 August 2021:
    //two classification levels since 1 July 2021, and Diffusion Restreinte, a protection mention that is not a classification
    public enum FrenchLevel
    {
        NonProtege = 0,          //[NP]
        DiffusionRestreinte = 1, //[DR], annexe 1
        Secret = 2,              //[S]
        TresSecret = 3           //[TS]
    }

    //the marking of one paragraph or of a whole document
    public class FrenchMarking
    {
        public FrenchLevel Level = FrenchLevel.NonProtege;
        public bool SpecialFrance;          //7.1.1.3 a): not a classification, never released to a foreign state or organisation
        public string MentionParticuliere;  //7.1.1.3 b): services, states or organisations allowed, or the special classification; free text

        public bool IsClassified
        {
            get { return Level >= FrenchLevel.Secret; }
        }

        public FrenchMarking Clone()
        {
            return (FrenchMarking)MemberwiseClone();
        }

        public override string ToString()
        {
            return FrenchFormatter.Stamp(this) + (SpecialFrance ? " " + FrenchFormatter.SpecialFranceStamp : "");
        }
    }

    //what the first page of a document carries, 7.1.2.3 b) and annexe 1 § 2
    public class FrenchIdentification
    {
        public string Lieu;
        public DateTime DateEmission = DateTime.Today;
        public string NumeroEnregistrement;
        public string AutoriteEmettrice;
        public string Auteur;
        public DateTime? Echeance;          //date of automatic declassification: "Déclassifié le [date]"
        public DateTime? Reevaluation;      //by exception, when no deadline can be set: "Classification à réévaluer le [date]", 20 years at most
        public bool SansEcheance;           //information on nuclear, radiological, biological or chemical weapons: no deadline, never declassified
        public int Exemplaire;              //TRÈS SECRET: copy number
        public int NombreExemplaires;       //TRÈS SECRET: total number of copies
        public string Reference;            //TRÈS SECRET: carried on every page
        public string Annexes;              //the annexes or plans of the document, listed on the first page (7.1.2.3 c)

        public const int MaximumReevaluationYears = 20;
        public const int CommunicabilityYears = 50;

        //checks the identification against the level of the document; empty when it is complete
        public List<string> Validate(FrenchMarking marking)
        {
            List<string> errors = new List<string>();
            if (marking.Level == FrenchLevel.NonProtege)
            {
                return errors;
            }
            //every classified or Diffusion Restreinte document is identified on its first page
            if (string.IsNullOrEmpty(AutoriteEmettrice))
            {
                errors.Add(Text.T("The issuing authority (or originating body) appears on the first page.", "L'autorité émettrice (ou l'organisme auteur) figure sur la première page.", "Die herausgebende Stelle (oder die verfassende Stelle) steht auf der ersten Seite."));
            }
            if (string.IsNullOrEmpty(NumeroEnregistrement))
            {
                errors.Add(Text.T("The registration number appears on the first page.", "Le numéro d'enregistrement figure sur la première page.", "Die Registriernummer steht auf der ersten Seite."));
            }
            if (!marking.IsClassified)
            {
                return errors; //Diffusion Restreinte: no deadline, no copy number
            }
            if (string.IsNullOrEmpty(Auteur))
            {
                errors.Add(Text.T("The author of the classified information appears on the first page.", "L'auteur de l'information classifiée figure sur la première page.", "Der Verfasser der eingestuften Information steht auf der ersten Seite."));
            }
            int choices = (Echeance.HasValue ? 1 : 0) + (Reevaluation.HasValue ? 1 : 0) + (SansEcheance ? 1 : 0);
            if (choices == 0)
            {
                errors.Add(Text.T("Classified information carries its classification deadline (\"Déclassifié le\") or, by exception, its re-evaluation date (IGI 1300, 7.6.1).", "Une information classifiée porte son échéance de classification (« Déclassifié le »), ou, à titre exceptionnel, sa date de réévaluation (IGI 1300, 7.6.1).", "Eingestufte Informationen tragen das Ende ihrer Einstufung (\"Déclassifié le\") oder ausnahmsweise ihr Datum der Neubewertung (IGI 1300, 7.6.1)."));
            }
            else if (choices > 1)
            {
                errors.Add(Text.T("Choose a single option: deadline, re-evaluation date, or no deadline.", "Choisir une seule option : échéance, date de réévaluation, ou aucune échéance.", "Wählen Sie nur eine Option: Ende der Einstufung, Datum der Neubewertung oder kein Ende."));
            }
            if (Echeance.HasValue)
            {
                if (Echeance.Value.Date <= DateEmission.Date)
                {
                    errors.Add(Text.T("The classification deadline comes after the issue date.", "L'échéance de classification est postérieure à la date d'émission.", "Das Ende der Einstufung liegt nach dem Ausstellungsdatum."));
                }
                else if (Echeance.Value.Date >= DateEmission.Date.AddYears(CommunicabilityYears))
                {
                    errors.Add(Text.T("The classification deadline comes before the end of the fifty-year period after which records may be released (IGI 1300, 7.6.1).", "L'échéance de classification est antérieure au délai de cinquante ans de communicabilité (IGI 1300, 7.6.1).", "Das Ende der Einstufung liegt vor Ablauf der Frist von fünfzig Jahren, nach der Unterlagen zugänglich gemacht werden dürfen (IGI 1300, 7.6.1)."));
                }
            }
            if (Reevaluation.HasValue)
            {
                if (Reevaluation.Value.Date <= DateEmission.Date)
                {
                    errors.Add(Text.T("The re-evaluation date comes after the issue date.", "La date de réévaluation est postérieure à la date d'émission.", "Das Datum der Neubewertung liegt nach dem Ausstellungsdatum."));
                }
                else if (Reevaluation.Value.Date > DateEmission.Date.AddYears(MaximumReevaluationYears))
                {
                    errors.Add(Text.T("The re-evaluation date is no more than twenty years from the production date (IGI 1300, 7.6.1).", "La date de réévaluation n'excède pas vingt ans à compter de la date de production (IGI 1300, 7.6.1).", "Das Datum der Neubewertung liegt höchstens zwanzig Jahre nach dem Erstellungsdatum (IGI 1300, 7.6.1)."));
                }
            }
            if (marking.Level == FrenchLevel.TresSecret)
            {
                if (Exemplaire <= 0 || NombreExemplaires <= 0)
                {
                    errors.Add(Text.T("At the Très Secret level, each document carries its copy number and the total number of copies (IGI 1300, 7.1.2.3 b).", "Au niveau Très Secret, chaque document porte son numéro d'exemplaire et le nombre total d'exemplaires (IGI 1300, 7.1.2.3 b).", "Auf der Stufe Très Secret trägt jedes Dokument seine Ausfertigungsnummer und die Gesamtzahl der Ausfertigungen (IGI 1300, 7.1.2.3 b)."));
                }
                else if (Exemplaire > NombreExemplaires)
                {
                    errors.Add(Text.T("The copy number cannot exceed the total number of copies.", "Le numéro d'exemplaire ne peut dépasser le nombre total d'exemplaires.", "Die Ausfertigungsnummer darf die Gesamtzahl der Ausfertigungen nicht übersteigen."));
                }
                if (string.IsNullOrEmpty(Reference))
                {
                    errors.Add(Text.T("At the Très Secret level, each page carries the document reference.", "Au niveau Très Secret, chaque page porte la référence du document.", "Auf der Stufe Très Secret trägt jede Seite das Aktenzeichen des Dokuments."));
                }
            }
            return errors;
        }

        //the lines of the identification block, the page count left to the caller (a field in Word)
        public string[] Lines(FrenchMarking marking)
        {
            List<string> lines = new List<string>();
            string place = string.IsNullOrEmpty(Lieu) ? "" : Lieu + ", ";
            lines.Add(place + "le " + FrenchFormatter.Date(DateEmission));
            if (!string.IsNullOrEmpty(NumeroEnregistrement)) lines.Add("N° " + NumeroEnregistrement);
            if (!string.IsNullOrEmpty(AutoriteEmettrice)) lines.Add("Autorité émettrice : " + AutoriteEmettrice);
            if (!string.IsNullOrEmpty(Auteur)) lines.Add("Auteur : " + Auteur);
            if (marking.IsClassified)
            {
                if (Echeance.HasValue) lines.Add("Déclassifié le " + FrenchFormatter.Date(Echeance.Value));
                if (Reevaluation.HasValue) lines.Add("Classification à réévaluer le " + FrenchFormatter.Date(Reevaluation.Value));
            }
            if (marking.Level == FrenchLevel.TresSecret)
            {
                lines.Add("Exemplaire n° " + Exemplaire + " sur " + NombreExemplaires);
                if (!string.IsNullOrEmpty(Reference)) lines.Add("Référence : " + Reference);
            }
            if (!string.IsNullOrEmpty(marking.MentionParticuliere)) lines.Add(marking.MentionParticuliere);
            if (marking.IsClassified)
            {
                //7.1.2.3 c): the total number of pages and the annexes on the first page; {NUMPAGES} becomes a field in Word
                lines.Add("Ce document comporte {NUMPAGES} pages" + (string.IsNullOrEmpty(Annexes) ? "" : ", annexes : " + Annexes));
            }
            return lines.ToArray();
        }
    }

    public static class FrenchFormatter
    {
        public const string SpecialFranceStamp = "SPÉCIAL FRANCE";
        //annexe 37 § 1: under the stamp of the cover of bound documents
        public const string CoverWarning = "Toute personne qui détient ce document sans avoir qualité pour le connaître tombe sous le coup des dispositions du code pénal réprimant les atteintes au secret de la défense nationale";
        //7.1.2.3 c): blank pages of TRÈS SECRET documents
        public const string BlankPage = "PAS DE TEXTE";

        //the stamp in full letters, empty for non protected information
        public static string Stamp(FrenchMarking m)
        {
            return Stamp(m.Level);
        }

        public static string Stamp(FrenchLevel level)
        {
            switch (level)
            {
                case FrenchLevel.TresSecret: return "TRÈS SECRET";
                case FrenchLevel.Secret: return "SECRET";
                case FrenchLevel.DiffusionRestreinte: return "DIFFUSION RESTREINTE";
                default: return "";
            }
        }

        //the level written as in the running text: "Très Secret", "Secret"
        public static string Name(FrenchLevel level)
        {
            switch (level)
            {
                case FrenchLevel.TresSecret: return "Très Secret";
                case FrenchLevel.Secret: return "Secret";
                case FrenchLevel.DiffusionRestreinte: return "Diffusion Restreinte";
                default: return "Non protégé";
            }
        }

        //the stamp at the bottom of each page: classification stamps only, Diffusion Restreinte and Spécial France stay at the top
        public static string BottomStamp(FrenchMarking m)
        {
            return m.IsClassified ? Stamp(m) : "";
        }

        public static string Abbreviation(FrenchLevel level)
        {
            switch (level)
            {
                case FrenchLevel.TresSecret: return "TS";
                case FrenchLevel.Secret: return "S";
                case FrenchLevel.DiffusionRestreinte: return "DR";
                default: return "NP";
            }
        }

        //the mark in the margin of a paragraph, 7.1.2.1 and annexe 36: "[S]", "[TS]", "[DR]", "[NP]";
        //the IGI gives no paragraph mark for Spécial France, the add-in writes "[S SF]"
        public static string Portion(FrenchMarking m)
        {
            return "[" + Abbreviation(m.Level) + (m.SpecialFrance && m.Level != FrenchLevel.NonProtege ? " SF" : "") + "]";
        }

        public static string Date(DateTime date)
        {
            return date.ToString("d MMMM yyyy", CultureInfo.GetCultureInfo("fr-FR"));
        }
    }

    public static class FrenchParser
    {
        private static readonly Regex PortionMark = new Regex(@"^\[(NP|DR|S|TS)( SF)?\]$");

        //a paragraph mark "[S]", "[TS SF]"..., null when the text is not one
        public static FrenchMarking ParsePortion(string text)
        {
            Match match = PortionMark.Match((text ?? "").Trim());
            if (!match.Success)
            {
                return null;
            }
            FrenchMarking m = new FrenchMarking();
            switch (match.Groups[1].Value)
            {
                case "TS": m.Level = FrenchLevel.TresSecret; break;
                case "S": m.Level = FrenchLevel.Secret; break;
                case "DR": m.Level = FrenchLevel.DiffusionRestreinte; break;
                default: m.Level = FrenchLevel.NonProtege; break;
            }
            m.SpecialFrance = match.Groups[2].Success;
            return m;
        }

        public static FrenchLevel? ParseStamp(string text)
        {
            foreach (FrenchLevel level in new FrenchLevel[] { FrenchLevel.TresSecret, FrenchLevel.Secret, FrenchLevel.DiffusionRestreinte })
            {
                if (text == FrenchFormatter.Stamp(level))
                {
                    return level;
                }
            }
            return null;
        }
    }

    public class FrenchRollUpResult
    {
        public FrenchMarking Document;
        public List<string> Errors = new List<string>();
        public List<string> Notes = new List<string>();
    }

    public static class FrenchRollUp
    {
        //a document is classified at the highest level of the information it contains (7.1.1.2 a); Spécial France,
        //which may cover only some parts (7.1.1.3 a), applies to the whole document as soon as one part carries it
        public static FrenchRollUpResult Compute(IList<FrenchMarking> portions)
        {
            FrenchRollUpResult result = new FrenchRollUpResult();
            FrenchMarking document = new FrenchMarking();
            result.Document = document;
            if (portions.Count == 0)
            {
                result.Errors.Add(Text.T("The document has no paragraph marks.", "Le document ne comporte aucune marque de paragraphe.", "Das Dokument enthält keine Absatzkennzeichnungen."));
                return result;
            }
            foreach (FrenchMarking portion in portions)
            {
                if (portion.Level > document.Level)
                {
                    document.Level = portion.Level;
                }
                if (portion.SpecialFrance)
                {
                    document.SpecialFrance = true;
                }
            }
            if (document.SpecialFrance && document.Level == FrenchLevel.NonProtege)
            {
                result.Errors.Add(Text.T("The Spécial France mention only applies to classified or Diffusion Restreinte information.", "La mention Spécial France ne s'applique qu'à des informations classifiées ou Diffusion Restreinte.", "Der Vermerk Spécial France gilt nur für eingestufte Informationen oder Diffusion Restreinte."));
            }
            return result;
        }
    }

    //checks one marking on its own
    public static class FrenchValidator
    {
        public static List<string> Validate(FrenchMarking m)
        {
            List<string> errors = new List<string>();
            if (m.SpecialFrance && m.Level == FrenchLevel.NonProtege)
            {
                errors.Add(Text.T("The Spécial France mention only applies to classified or Diffusion Restreinte information.", "La mention Spécial France ne s'applique qu'à des informations classifiées ou Diffusion Restreinte.", "Der Vermerk Spécial France gilt nur für eingestufte Informationen oder Diffusion Restreinte."));
            }
            if (!string.IsNullOrEmpty(m.MentionParticuliere) && !m.IsClassified)
            {
                errors.Add(Text.T("Particular mentions (7.1.1.3 b) go with classified information.", "Les mentions particulières (7.1.1.3 b) accompagnent des informations classifiées.", "Besondere Vermerke (7.1.1.3 b) gehören zu eingestuften Informationen."));
            }
            return errors;
        }
    }
}
