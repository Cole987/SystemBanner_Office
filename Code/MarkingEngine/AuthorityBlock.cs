//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace SystemBannerOffice.Markings
{
    //the classification authority block of the Register and Manual, section E (EO 13526, 32 CFR 2001, ICD 710):
    //original decisions carry a Classification Reason, derivative ones a Derived From line
    public class AuthorityBlock
    {
        public bool Original;            //true for an OCA decision, false for a derivative one
        public string ClassifiedBy;      //name and position, or personal identifier
        public string AgencyAndOffice;   //if not otherwise evident
        public string Reason;            //original only: cites at least one EO 13526 section 1.4 category
        public string DerivedFrom;       //derivative only: source or guide, "Multiple Sources" with a source list
        public string DeclassifyOn;
        public DateTime DateOfOrigin = DateTime.Today;
        public List<string> Sources = new List<string>(); //listed on or attached to the document when DerivedFrom is "Multiple Sources"

        public const string MultipleSources = "Multiple Sources";

        public string[] Lines()
        {
            List<string> lines = new List<string>();
            string by = ClassifiedBy ?? "";
            if (!string.IsNullOrEmpty(AgencyAndOffice))
            {
                by += ", " + AgencyAndOffice;
            }
            lines.Add("Classified By: " + by);
            if (Original)
            {
                lines.Add("Reason: " + (Reason ?? ""));
            }
            else
            {
                lines.Add("Derived From: " + (DerivedFrom ?? ""));
            }
            lines.Add("Declassify On: " + (DeclassifyOn ?? ""));
            return lines.ToArray();
        }

        //checks the block against the banner of the document; empty when it is complete and valid
        public List<string> Validate(Marking banner)
        {
            List<string> errors = new List<string>();
            if (banner.Owner != Owner.UnitedStates || banner.Level == Level.Unclassified)
            {
                return errors; //the block is required on the face of US classified NSI documents only
            }
            if (string.IsNullOrEmpty(ClassifiedBy))
            {
                errors.Add(Text.T("Classified By is required: name and position, or personal identifier.", "La ligne Classified By est obligatoire : nom et fonction, ou identifiant personnel.", "Classified By ist erforderlich: Name und Funktion oder persönliche Kennung."));
            }
            if (Original)
            {
                if (string.IsNullOrEmpty(Reason))
                {
                    errors.Add(Text.T("Reason is required on an original classification: cite at least one category of EO 13526, section 1.4.", "La ligne Reason est obligatoire pour une classification originale : citer au moins une catégorie de l'EO 13526, section 1.4.", "Reason ist bei einer Ersteinstufung erforderlich: Nennen Sie mindestens eine Kategorie aus EO 13526, Abschnitt 1.4."));
                }
                else if (!Regex.IsMatch(Reason, @"1\.4\s*\(?[a-h]\)?", RegexOptions.IgnoreCase))
                {
                    errors.Add(Text.T("The Reason should cite a category of EO 13526, section 1.4, for example 1.4(c).", "La ligne Reason doit citer une catégorie de l'EO 13526, section 1.4, par exemple 1.4(c).", "Reason sollte eine Kategorie aus EO 13526, Abschnitt 1.4, nennen, zum Beispiel 1.4(c)."));
                }
            }
            else
            {
                if (string.IsNullOrEmpty(DerivedFrom))
                {
                    errors.Add(Text.T("Derived From is required on a derivative classification: the source document or classification guide, with its agency and date.", "La ligne Derived From est obligatoire pour une classification dérivée : document source ou guide de classification, avec son organisme et sa date.", "Derived From ist bei einer abgeleiteten Einstufung erforderlich: das Quelldokument oder der Einstufungsleitfaden, mit Behörde und Datum."));
                }
                else if (DerivedFrom == MultipleSources && Sources.Count == 0)
                {
                    errors.Add(Text.T("\"Multiple Sources\" needs the list of all sources on or attached to the document.", "« Multiple Sources » exige la liste de toutes les sources, portée sur le document ou jointe à celui-ci.", "\"Multiple Sources\" erfordert die Liste aller Quellen, auf dem Dokument oder als Anlage."));
                }
            }
            string required = RequiredNotApplicable(banner);
            if (required != null)
            {
                if (DeclassifyOn != required)
                {
                    errors.Add(Text.T("This document commingles NSI with " + (required.Substring(7, required.IndexOf(" portions") - 7)) + " information: Declassify On must read \"" + required + "\", with the NSI source list.", "Ce document mêle des informations NSI et " + (required.Substring(7, required.IndexOf(" portions") - 7)) + " : la ligne Declassify On doit indiquer « " + required + " », avec la liste des sources NSI.", "Dieses Dokument vermischt NSI mit " + (required.Substring(7, required.IndexOf(" portions") - 7)) + "-Informationen: Declassify On muss \"" + required + "\" lauten, mit der Liste der NSI-Quellen."));
                }
                return errors;
            }
            DeclassifyValue value = DeclassifyValue.Parse(DeclassifyOn);
            if (value.Error != null)
            {
                errors.Add(value.Error);
            }
            else if (value.Kind == DeclassifyKind.Date && Original && value.Date > DateOfOrigin.AddYears(25))
            {
                errors.Add(Text.T("An original declassification date may not be more than 25 years after the decision (" + DateOfOrigin.AddYears(25).ToString("yyyyMMdd") + ").", "Une date de déclassification originale ne peut pas dépasser 25 ans après la décision (" + DateOfOrigin.AddYears(25).ToString("yyyyMMdd") + ").", "Ein ursprüngliches Datum für die Aufhebung der Einstufung darf höchstens 25 Jahre nach der Entscheidung liegen (" + DateOfOrigin.AddYears(25).ToString("yyyyMMdd") + ")."));
            }
            else if (value.Kind == DeclassifyKind.Date && value.Date <= DateTime.Today)
            {
                errors.Add(Text.T("The declassification date has passed: do not assume the information is unclassified, ask the originating agency.", "La date de déclassification est passée : ne pas considérer l'information comme non classifiée, interroger l'organisme d'origine.", "Das Datum für die Aufhebung der Einstufung ist überschritten: Gehen Sie nicht davon aus, dass die Information offen ist, fragen Sie die herausgebende Behörde."));
            }
            else if (value.Kind == DeclassifyKind.NotApplicable)
            {
                errors.Add(Text.T("\"N/A to ... portions\" is only for documents with RD, FRD, TFNI or NATO portions.", "« N/A to ... portions » est réservé aux documents comportant des portions RD, FRD, TFNI ou NATO.", "\"N/A to ... portions\" gilt nur für Dokumente mit RD-, FRD-, TFNI- oder NATO-Abschnitten."));
            }
            return errors;
        }

        //"N/A to RD/FRD/TFNI [and NATO] portions. See source list for NSI portions." when the banner carries such information, null otherwise
        public static string RequiredNotApplicable(Marking banner)
        {
            List<string> kinds = new List<string>();
            if (banner.RestrictedData) kinds.Add("RD");
            if (banner.FormerlyRestrictedData) kinds.Add("FRD");
            if (banner.Tfni) kinds.Add("TFNI");
            bool nato = banner.Fgi && banner.FgiSources.Contains("NATO");
            if (kinds.Count == 0 && !nato)
            {
                return null;
            }
            string text = string.Join("/", kinds.ToArray());
            if (nato)
            {
                text = text.Length == 0 ? "NATO" : text + " and NATO";
            }
            return "N/A to " + text + " portions. See source list for NSI portions.";
        }
    }

    public enum DeclassifyKind
    {
        NotApplicable,  //N/A to RD/FRD/TFNI/NATO portions
        Hum,            //50X1-HUM
        Wmd,            //50X2-WMD
        Exemption50,    //50X#, date or event (and 75X#)
        Imagery,        //25X1, EO 12951
        Exemption25,    //25X#, date or event
        Exemption25Bare,//25X# with no date or event: 50 years from the source date
        Date,
        Event,
        Retired
    }

    //one Declassify On value, ranked by the hierarchy of section E.3 to pick the most restrictive of several sources
    public class DeclassifyValue
    {
        public DeclassifyKind Kind;
        public string Text;
        public int Exemption;       //the # of 25X#, 50X#, 75X#
        public int ExemptionYears;  //25, 50 or 75
        public DateTime Date;       //for dates, and exemptions with a date
        public string Event;        //for events, and exemptions with an event
        public string Error;        //why the value is not valid, null when it is

        private static readonly Regex DateFormat = new Regex(@"^\d{8}$");
        private static readonly Regex ExemptionFormat = new Regex(@"^(25|50|75)X([1-9])(?:,\s*(.+))?$");

        public static DeclassifyValue Parse(string text)
        {
            DeclassifyValue value = new DeclassifyValue();
            value.Text = (text ?? "").Trim();
            string t = value.Text;
            if (t.Length == 0)
            {
                value.Error = SystemBannerOffice.Markings.Text.T("Declassify On is required: a date (YYYYMMDD), an event, or an exemption.", "La ligne Declassify On est obligatoire : une date (AAAAMMJJ), un événement ou une exemption.", "Declassify On ist erforderlich: ein Datum (JJJJMMTT), ein Ereignis oder eine Ausnahme.");
                return value;
            }
            string retired = RetiredReplacement(t);
            if (retired != null)
            {
                value.Kind = DeclassifyKind.Retired;
                value.Error = SystemBannerOffice.Markings.Text.T("\"" + t + "\" is retired or invalid: " + retired, "« " + t + " » est retiré ou non valide : " + retired, "\"" + t + "\" ist zurückgezogen oder ungültig:" + retired);
                return value;
            }
            if (t.StartsWith("N/A to "))
            {
                value.Kind = DeclassifyKind.NotApplicable;
                return value;
            }
            if (t == "50X1-HUM")
            {
                value.Kind = DeclassifyKind.Hum;
                return value;
            }
            if (t == "50X2-WMD")
            {
                value.Kind = DeclassifyKind.Wmd;
                return value;
            }
            if (t == "25X1, EO 12951")
            {
                value.Kind = DeclassifyKind.Imagery;
                value.Error = null; //reserved for D/NGA original imagery, the add-in cannot check who classifies
                return value;
            }
            if (DateFormat.IsMatch(t))
            {
                value.Kind = DeclassifyKind.Date;
                if (!TryDate(t, out value.Date))
                {
                    value.Error = SystemBannerOffice.Markings.Text.T("\"" + t + "\" is not a valid date: use YYYYMMDD.", "« " + t + " » n'est pas une date valide : utiliser AAAAMMJJ.", "\"" + t + "\" ist kein gültiges Datum: Verwenden Sie JJJJMMTT.");
                }
                return value;
            }
            Match exemption = ExemptionFormat.Match(t);
            if (exemption.Success)
            {
                value.ExemptionYears = int.Parse(exemption.Groups[1].Value);
                value.Exemption = int.Parse(exemption.Groups[2].Value);
                string rest = exemption.Groups[3].Success ? exemption.Groups[3].Value.Trim() : null;
                if (value.ExemptionYears == 25)
                {
                    value.Kind = rest == null ? DeclassifyKind.Exemption25Bare : DeclassifyKind.Exemption25;
                }
                else
                {
                    value.Kind = DeclassifyKind.Exemption50;
                    if (rest == null)
                    {
                        value.Error = SystemBannerOffice.Markings.Text.T("\"" + t + "\" needs a date or an event: \"" + t + ", YYYYMMDD\".", "« " + t + " » exige une date ou un événement : « " + t + ", AAAAMMJJ ».", "\"" + t + "\" erfordert ein Datum oder ein Ereignis: \"" + t + ", JJJJMMTT\".");
                    }
                }
                if (rest != null)
                {
                    if (DateFormat.IsMatch(rest))
                    {
                        if (!TryDate(rest, out value.Date))
                        {
                            value.Error = SystemBannerOffice.Markings.Text.T("\"" + rest + "\" is not a valid date: use YYYYMMDD.", "« " + rest + " » n'est pas une date valide : utiliser AAAAMMJJ.", "\"" + rest + "\" ist kein gültiges Datum: Verwenden Sie JJJJMMTT.");
                        }
                    }
                    else
                    {
                        value.Event = rest;
                    }
                }
                return value;
            }
            if (Regex.IsMatch(t, @"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$"))
            {
                value.Error = SystemBannerOffice.Markings.Text.T("\"" + t + "\": dates are written YYYYMMDD.", "« " + t + " » : les dates s'écrivent AAAAMMJJ.", "\"" + t + "\": Datumsangaben werden als JJJJMMTT geschrieben.");
                return value;
            }
            //anything else is an event, which must be reasonably definite, foreseeable and less than 10 years away: the add-in cannot check that
            value.Kind = DeclassifyKind.Event;
            value.Event = t;
            return value;
        }

        private static bool TryDate(string text, out DateTime date)
        {
            return DateTime.TryParseExact(text, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        //what to do instead of a retired value, section E.6; null when the value is not retired
        public static string RetiredReplacement(string text)
        {
            string t = text.ToUpperInvariant();
            if (t == "OADR" || t.StartsWith("ORIGINATING AGENCY'S DETERMINATION REQUIRED") || t.StartsWith("SOURCE MARKED OADR")
                || t == "MR" || t.StartsWith("MANUAL REVIEW") || t.StartsWith("SOURCE MARKED MR")
                || Regex.IsMatch(t, @"^(SOURCE MARKED )?X[1-8]\b"))
            {
                return SystemBannerOffice.Markings.Text.T("use the date 25 years after the date of the source document, or 25 years from today when it is unknown.", "utiliser la date située 25 ans après la date du document source, ou 25 ans à compter d'aujourd'hui si elle est inconnue.", "verwenden Sie das Datum 25 Jahre nach dem Datum des Quelldokuments oder 25 Jahre ab heute, wenn es unbekannt ist.");
            }
            if (t == "DNI ONLY" || t == "DCI ONLY")
            {
                return SystemBannerOffice.Markings.Text.T("use a date 25 years after the source document, or the 25X exemption of the source.", "utiliser une date située 25 ans après le document source, ou l'exemption 25X de la source.", "verwenden Sie ein Datum 25 Jahre nach dem Quelldokument oder die Ausnahme 25X der Quelle.");
            }
            if (t.StartsWith("SUBJECT TO TREATY OR INTERNATIONAL AGREEMENT"))
            {
                return SystemBannerOffice.Markings.Text.T("use the exemption set by OCA guidance, or the date 25 years after the source document.", "utiliser l'exemption fixée par les instructions de l'OCA, ou la date située 25 ans après le document source.", "verwenden Sie die in den OCA-Vorgaben festgelegte Ausnahme oder das Datum 25 Jahre nach dem Quelldokument.");
            }
            if (t == "25X1-HUMAN")
            {
                return SystemBannerOffice.Markings.Text.T("use \"50X1-HUM\".", "utiliser « 50X1-HUM ».", "verwenden Sie \"50X1-HUM\".");
            }
            if (Regex.IsMatch(t, @"^SOURCE MARKED 25X[1-9]"))
            {
                return SystemBannerOffice.Markings.Text.T("use the date 50 years after the date of the source document, or 50 years from today when it is unknown.", "utiliser la date située 50 ans après la date du document source, ou 50 ans à compter d'aujourd'hui si elle est inconnue.", "verwenden Sie das Datum 50 Jahre nach dem Datum des Quelldokuments oder 50 Jahre ab heute, wenn es unbekannt ist.");
            }
            return null;
        }

        //rank in the section E.3 hierarchy, lower is more restrictive
        private int Rank
        {
            get
            {
                switch (Kind)
                {
                    case DeclassifyKind.NotApplicable: return 0;
                    case DeclassifyKind.Hum:
                    case DeclassifyKind.Wmd: return 1;
                    case DeclassifyKind.Exemption50: return 2;
                    case DeclassifyKind.Imagery: return 3;
                    case DeclassifyKind.Exemption25: return 4;
                    case DeclassifyKind.Exemption25Bare: return 5;
                    case DeclassifyKind.Date: return 6;
                    case DeclassifyKind.Event: return 7;
                    default: return 8;
                }
            }
        }

        //the single Declassify On value of a document derived from several sources, section E.3;
        //sourceDates gives the date of each source for 25X# values with no date (50 years from the source)
        public static string MostRestrictive(IList<string> values, IList<DateTime> sourceDates)
        {
            DeclassifyValue best = null;
            DateTime bestDate = DateTime.MinValue;
            List<string> notApplicable = new List<string>();
            for (int i = 0; i < values.Count; i++)
            {
                DeclassifyValue value = Parse(values[i]);
                if (value.Kind == DeclassifyKind.Retired || value.Error != null)
                {
                    continue; //the caller converts retired values first, see RetiredReplacement
                }
                if (value.Kind == DeclassifyKind.NotApplicable)
                {
                    notApplicable.Add(value.Text);
                }
                DateTime date = value.Date;
                if (value.Kind == DeclassifyKind.Exemption25Bare)
                {
                    DateTime source = sourceDates != null && i < sourceDates.Count ? sourceDates[i] : DateTime.Today;
                    date = source.AddYears(50);
                }
                if (best == null || value.Rank < best.Rank || (value.Rank == best.Rank && Better(value, date, best, bestDate)))
                {
                    best = value;
                    bestDate = date;
                }
            }
            if (best == null)
            {
                return null;
            }
            if (best.Kind == DeclassifyKind.NotApplicable)
            {
                return CombineNotApplicable(notApplicable);
            }
            if (best.Kind == DeclassifyKind.Exemption25Bare)
            {
                return "25X" + best.Exemption + ", " + bestDate.ToString("yyyyMMdd");
            }
            return best.Text;
        }

        //same rank: the longest protection wins, then the lowest exemption number
        private static bool Better(DeclassifyValue value, DateTime date, DeclassifyValue best, DateTime bestDate)
        {
            if (value.Kind == DeclassifyKind.Hum || value.Kind == DeclassifyKind.Wmd)
            {
                return value.Kind == DeclassifyKind.Hum && best.Kind == DeclassifyKind.Wmd; //50X1-HUM has the lowest number
            }
            if (value.Event == null && best.Event == null && date != bestDate)
            {
                return date > bestDate;
            }
            if (value.Exemption != best.Exemption && (value.Event == best.Event || (value.Event == null && date == bestDate)))
            {
                return value.Exemption < best.Exemption;
            }
            return false; //events cannot be compared: keep the first one, the classifier checks it
        }

        private static string CombineNotApplicable(List<string> values)
        {
            bool rd = false, frd = false, tfni = false, nato = false;
            foreach (string value in values)
            {
                string head = value.Substring(0, value.IndexOf(" portions") > 0 ? value.IndexOf(" portions") : value.Length);
                rd |= Regex.IsMatch(head, @"\bRD\b");
                frd |= head.Contains("FRD");
                tfni |= head.Contains("TFNI");
                nato |= head.Contains("NATO");
            }
            Marking marking = new Marking();
            marking.Level = Level.Secret;
            marking.RestrictedData = rd;
            marking.FormerlyRestrictedData = frd;
            marking.Tfni = tfni;
            if (nato)
            {
                marking.Fgi = true;
                marking.FgiSources.Add("NATO");
            }
            return AuthorityBlock.RequiredNotApplicable(marking);
        }
    }
}
