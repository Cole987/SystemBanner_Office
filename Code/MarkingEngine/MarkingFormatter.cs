//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Text;

namespace SystemBannerOffice.Markings
{
    //writes banner lines and portion marks, in the order and with the separators of the Register and Manual, A.6 Formatting:
    //categories separated by "//", values inside a category by "/", sub-markings joined by "-", sub-compartments by " "
    public static class MarkingFormatter
    {
        //banner lines may spell control markings out or use their authorized abbreviations, the add-in uses abbreviations unless told otherwise
        public static bool Abbreviate = true;

        public static string Banner(Marking marking)
        {
            return Write(marking, false, Abbreviate, null);
        }

        public static string Banner(Marking marking, bool abbreviate)
        {
            return Write(marking, false, abbreviate, null);
        }

        //portion mark, without its parentheses
        public static string Portion(Marking marking)
        {
            return Write(marking, true, true, null);
        }

        //portion mark of a portion inside a document with this banner: "REL" may stand for the REL TO list
        //when the portion is releasable to the same list as the banner line
        public static string Portion(Marking marking, Marking banner)
        {
            return Write(marking, true, true, banner);
        }

        public static string PortionWithParentheses(Marking marking)
        {
            return "(" + Portion(marking) + ")";
        }

        private static string Write(Marking m, bool portion, bool abbreviate, Marking banner)
        {
            List<string> categories = new List<string>();
            categories.Add(Classification(m, portion, banner));
            string sci = Sci(m, portion, abbreviate);
            if (sci.Length > 0) categories.Add(sci);
            string sap = Sap(m, portion, abbreviate);
            if (sap.Length > 0) categories.Add(sap);
            string aea = Aea(m, portion, abbreviate);
            if (aea.Length > 0) categories.Add(aea);
            string fgi = Fgi(m, portion);
            if (fgi.Length > 0) categories.Add(fgi);
            string dissem = Dissemination(m, portion, abbreviate, banner);
            if (dissem.Length > 0) categories.Add(dissem);
            string nonIc = NonIcDissemination(m, portion, abbreviate);
            if (nonIc.Length > 0) categories.Add(nonIc);
            //non-US, NATO and JOINT markings start with "//": the first category is empty
            return string.Join("//", categories.ToArray());
        }

        public static string LevelTitle(Level level)
        {
            switch (level)
            {
                case Level.TopSecret: return "TOP SECRET";
                case Level.Secret: return "SECRET";
                case Level.Confidential: return "CONFIDENTIAL";
                case Level.Restricted: return "RESTRICTED";
                default: return "UNCLASSIFIED";
            }
        }

        public static string LevelPortion(Level level)
        {
            switch (level)
            {
                case Level.TopSecret: return "TS";
                case Level.Secret: return "S";
                case Level.Confidential: return "C";
                case Level.Restricted: return "R";
                default: return "U";
            }
        }

        //NATO levels, Appendix B: COSMIC TOP SECRET (CTS), NATO SECRET (NS), NATO CONFIDENTIAL (NC), NATO RESTRICTED (NR), NATO UNCLASSIFIED (NU)
        public static string NatoTitle(Marking m)
        {
            string activity = m.NatoActivity != null ? m.NatoActivity + " " : "";
            if (m.Level == Level.TopSecret)
            {
                return "COSMIC TOP SECRET";
            }
            return "NATO " + activity + LevelTitle(m.Level);
        }

        public static string NatoPortion(Marking m)
        {
            if (m.Level == Level.TopSecret)
            {
                return "CTS";
            }
            return "N" + (m.NatoActivity != null ? Register.NacPortionCode(m.NatoActivity) : "") + LevelPortion(m.Level);
        }

        private static string Classification(Marking m, bool portion, Marking banner)
        {
            switch (m.Owner)
            {
                case Owner.NonUS:
                    {
                        //"//GBR SECRET", portion "//GBR S", several sources "//CAN GBR S", concealed source "//FGI S"
                        string owners = m.OwnerCode == null ? "FGI" : string.Join(" ", Sorting.SortCountries(m.OwnerCode.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries), false).ToArray());
                        return "//" + owners + " " + (portion ? LevelPortion(m.Level) : LevelTitle(m.Level));
                    }
                case Owner.Nato:
                    return "//" + (portion ? NatoPortion(m) : NatoTitle(m));
                case Owner.Joint:
                    {
                        //"//JOINT SECRET CAN GBR USA", co-owners in ascending order; a portion with the co-owners of the banner leaves them out
                        string level = portion ? LevelPortion(m.Level) : LevelTitle(m.Level);
                        bool sameAsBanner = m.JointOwners.Count == 0 || (portion && banner != null && banner.Owner == Owner.Joint && SameList(m.JointOwners, banner.JointOwners));
                        if (sameAsBanner && portion)
                        {
                            return "//JOINT " + level;
                        }
                        return "//JOINT " + level + " " + string.Join(" ", Sorting.SortCountries(m.JointOwners, false).ToArray());
                    }
                default:
                    return portion ? LevelPortion(m.Level) : LevelTitle(m.Level);
            }
        }

        //NATO special access program markings in Appendix B order
        private static List<string> SortedNatoSap(Marking m, bool atomal)
        {
            List<string> saps = new List<string>();
            foreach (string sap in Register.NatoSap)
            {
                if (m.NatoSap.Contains(sap) && (sap == "ATOMAL") == atomal)
                {
                    saps.Add(sap);
                }
            }
            return saps;
        }

        //SCI: "123/SI-G ABCD DEFG-MMM AACD", systems and compartments numbers first then letters, at each level
        private static string Sci(Marking m, bool portion, bool abbreviate)
        {
            List<ControlSystem> systems = new List<ControlSystem>(m.Sci);
            systems.Sort(delegate (ControlSystem a, ControlSystem b) { return Sorting.CompareAlphanumeric(SciName(a.Name, true), SciName(b.Name, true)); });
            List<string> parts = new List<string>();
            foreach (ControlSystem system in systems)
            {
                SciEntry entry = Register.FindSci(system.Name);
                bool shortForm = portion || abbreviate;
                StringBuilder text = new StringBuilder(SciName(system.Name, shortForm));
                AppendCompartments(text, system, entry != null ? entry.Children : null, shortForm);
                parts.Add(text.ToString());
            }
            return string.Join("/", parts.ToArray());
        }

        //abbreviation (or title) of a published SCI value, the value itself when it is not published
        private static string SciName(string name, bool shortForm)
        {
            SciEntry entry = Register.FindSci(name);
            if (entry == null)
            {
                return name;
            }
            return shortForm ? entry.Abbreviation : entry.Title;
        }

        private static void AppendCompartments(StringBuilder text, ControlSystem system, SciEntry[] published, bool shortForm)
        {
            List<Compartment> compartments = new List<Compartment>(system.Compartments);
            compartments.Sort(delegate (Compartment a, Compartment b) { return Sorting.CompareAlphanumeric(CompartmentName(a.Name, published, true), CompartmentName(b.Name, published, true)); });
            foreach (Compartment compartment in compartments)
            {
                text.Append('-').Append(CompartmentName(compartment.Name, published, shortForm));
                List<string> subs = new List<string>(compartment.SubCompartments);
                Sorting.SortAlphanumeric(subs);
                foreach (string sub in subs)
                {
                    text.Append(' ').Append(sub);
                }
            }
        }

        private static string CompartmentName(string name, SciEntry[] published, bool shortForm)
        {
            SciEntry entry = Register.FindSci(published, name);
            if (entry == null)
            {
                return name;
            }
            return shortForm ? entry.Abbreviation : entry.Title;
        }

        //SAP: "SAR-ABC-DEF 123/SDA-121", the category indicator is written once; NATO BALK and BOHEMIA follow the US programs
        private static string Sap(Marking m, bool portion, bool abbreviate)
        {
            List<string> nato = SortedNatoSap(m, false);
            if (m.Sap.Count == 0 && !m.SapProgramsHidden && nato.Count == 0)
            {
                return "";
            }
            List<string> parts = new List<string>();
            bool shortForm = portion || abbreviate;
            if (m.Sap.Count == 0 && m.SapProgramsHidden)
            {
                parts.Add("SAR"); //program identifiers are optional in the SAP hierarchy
            }
            else if (m.Sap.Count > 0)
            {
                List<ControlSystem> programs = new List<ControlSystem>(m.Sap);
                programs.Sort(delegate (ControlSystem a, ControlSystem b) { return Sorting.CompareAlphanumeric(ProgramName(a, true), ProgramName(b, true)); });
                List<string> written = new List<string>();
                foreach (ControlSystem program in programs)
                {
                    StringBuilder text = new StringBuilder(ProgramName(program, shortForm));
                    AppendCompartments(text, program, null, true);
                    written.Add(text.ToString());
                }
                parts.Add((shortForm ? "SAR-" : "SPECIAL ACCESS REQUIRED-") + string.Join("/", written.ToArray()));
            }
            parts.AddRange(nato);
            return string.Join("/", parts.ToArray());
        }

        //portion marks always use the abbreviation of the program identifier when there is one, banner lines may spell it out
        private static string ProgramName(ControlSystem program, bool shortForm)
        {
            return shortForm && program.Abbreviation != null ? program.Abbreviation : program.Name;
        }

        //AEA, in Register order: RD (with CNWDI and SIGMA), FRD (with SIGMA), DOD UCNI, DOE UCNI, TFNI
        private static string Aea(Marking m, bool portion, bool abbreviate)
        {
            List<string> parts = new List<string>();
            bool shortForm = portion || abbreviate;
            if (m.RestrictedData)
            {
                string rd = shortForm ? "RD" : "RESTRICTED DATA";
                if (m.Cnwdi)
                {
                    rd += "-" + (shortForm ? "CNWDI" : "CRITICAL NUCLEAR WEAPON DESIGN INFORMATION");
                }
                rd += Sigma(m.RdSigma, portion);
                parts.Add(rd);
            }
            if (m.FormerlyRestrictedData)
            {
                parts.Add((shortForm ? "FRD" : "FORMERLY RESTRICTED DATA") + Sigma(m.FrdSigma, portion));
            }
            if (m.DodUcni)
            {
                parts.Add(portion ? "DCNI" : (abbreviate ? "DOD UCNI" : "DOD UNCLASSIFIED CONTROLLED NUCLEAR INFORMATION"));
            }
            if (m.DoeUcni)
            {
                parts.Add(portion ? "UCNI" : (abbreviate ? "DOE UCNI" : "DOE UNCLASSIFIED CONTROLLED NUCLEAR INFORMATION"));
            }
            if (m.Tfni)
            {
                parts.Add(shortForm ? "TFNI" : "TRANSCLASSIFIED FOREIGN NUCLEAR INFORMATION");
            }
            parts.AddRange(SortedNatoSap(m, true)); //ATOMAL follows the US AEA markings: "SECRET//RD/ATOMAL//FGI NATO//NOFORN"
            return string.Join("/", parts.ToArray());
        }

        //"-SIGMA 1 2" in the banner line, "-SG 1 2" in portion marks, numbers in ascending order
        private static string Sigma(List<int> numbers, bool portion)
        {
            if (numbers.Count == 0)
            {
                return "";
            }
            List<int> sorted = new List<int>(numbers);
            sorted.Sort();
            StringBuilder text = new StringBuilder(portion ? "-SG" : "-SIGMA");
            int previous = -1;
            foreach (int number in sorted)
            {
                if (number != previous)
                {
                    text.Append(' ').Append(number);
                }
                previous = number;
            }
            return text.ToString();
        }

        //FGI inside US information: "FGI GBR JPN NATO", or "FGI" alone when the source is concealed
        private static string Fgi(Marking m, bool portion)
        {
            if (!m.Fgi)
            {
                return "";
            }
            if (m.FgiSources.Count == 0)
            {
                return "FGI";
            }
            return "FGI " + string.Join(" ", Sorting.SortCountries(m.FgiSources, false).ToArray());
        }

        private static string Dissemination(Marking m, bool portion, bool abbreviate, Marking banner)
        {
            List<string> keys = new List<string>(m.Dissem);
            Register.SortByRegister(keys, Register.Dissemination);
            List<string> parts = new List<string>();
            foreach (string key in keys)
            {
                RegisterEntry entry = Register.FindDissemination(key);
                if (entry == null)
                {
                    parts.Add(key);
                    continue;
                }
                switch (key)
                {
                    case "REL":
                        if (portion && (m.RelTo.Count == 0 || (banner != null && banner.HasDissem("REL") && SameList(m.RelTo, banner.RelTo))))
                        {
                            parts.Add("REL"); //same countries as the banner line
                        }
                        else
                        {
                            parts.Add("REL TO " + string.Join(", ", Sorting.SortCountries(m.RelTo, true).ToArray()));
                        }
                        break;
                    case "EYES":
                        //"USA/GBR EYES ONLY" in the banner line, "EYES" in portion marks with the countries of the banner line
                        if (portion && (m.EyesOnly.Count == 0 || (banner != null && banner.HasDissem("EYES") && SameList(m.EyesOnly, banner.EyesOnly))))
                        {
                            parts.Add("EYES");
                        }
                        else
                        {
                            parts.Add(string.Join("/", Sorting.SortCountries(m.EyesOnly, true).ToArray()) + " EYES ONLY");
                        }
                        break;
                    case "DISPLAYONLY":
                        parts.Add("DISPLAY ONLY " + string.Join(", ", Sorting.SortCountries(m.DisplayOnly, false).ToArray()));
                        break;
                    default:
                        parts.Add(portion ? entry.Portion : entry.BannerText(abbreviate));
                        break;
                }
            }
            return string.Join("/", parts.ToArray());
        }

        private static string NonIcDissemination(Marking m, bool portion, bool abbreviate)
        {
            List<string> keys = new List<string>(m.NonIcDissem);
            Register.SortByRegister(keys, Register.NonIcDissemination);
            List<string> parts = new List<string>();
            foreach (string key in keys)
            {
                RegisterEntry entry = Register.FindNonIcDissemination(key);
                if (entry == null)
                {
                    parts.Add(key);
                }
                else
                {
                    parts.Add(portion ? entry.Portion : entry.BannerText(abbreviate));
                }
            }
            return string.Join("/", parts.ToArray());
        }

        public static bool SameList(List<string> a, List<string> b)
        {
            List<string> sa = Sorting.SortCountries(a, true);
            List<string> sb = Sorting.SortCountries(b, true);
            if (sa.Count != sb.Count)
            {
                return false;
            }
            for (int i = 0; i < sa.Count; i++)
            {
                if (sa[i] != sb[i])
                {
                    return false;
                }
            }
            return true;
        }
    }
}
