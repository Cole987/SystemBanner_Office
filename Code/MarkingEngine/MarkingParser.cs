//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;

namespace SystemBannerOffice.Markings
{
    public class ParseResult
    {
        public Marking Marking;
        public List<string> Errors = new List<string>();
        public bool Success
        {
            get { return Errors.Count == 0; }
        }
    }

    //reads banner lines and portion marks back into a Marking, in their spelled out or abbreviated forms,
    //so portion marks typed by hand or found in a document can be rolled up
    public static class MarkingParser
    {
        public static ParseResult Parse(string text)
        {
            ParseResult result = new ParseResult();
            Marking m = new Marking();
            result.Marking = m;
            string value = (text ?? "").Trim();
            if (value.StartsWith("(") && value.EndsWith(")"))
            {
                value = value.Substring(1, value.Length - 2).Trim();
            }
            if (value.Length == 0)
            {
                result.Errors.Add("The marking is empty.");
                return result;
            }
            if (value != value.ToUpperInvariant())
            {
                result.Errors.Add("Markings must be in uppercase letters.");
                value = value.ToUpperInvariant();
            }
            List<string> categories = new List<string>(value.Split(new string[] { "//" }, StringSplitOptions.None));
            if (categories[0].Length == 0)
            {
                //non-US, NATO or JOINT: "//GBR SECRET//..."
                categories.RemoveAt(0);
                if (categories.Count == 0 || !ParseForeignClassification(categories[0], m, result))
                {
                    return result;
                }
            }
            else if (!ParseUsClassification(categories[0], m))
            {
                result.Errors.Add("\"" + categories[0] + "\" is not a classification marking.");
                return result;
            }
            int lastCategory = 0; //categories must come in Register order: SCI 1, SAP 2, AEA 3, FGI 4, dissemination 5, non-IC dissemination 6
            for (int i = 1; i < categories.Count; i++)
            {
                string category = categories[i].Trim();
                int kind;
                if (category.Length == 0)
                {
                    result.Errors.Add("Empty category: no \"////\" placeholder is allowed for categories that do not apply.");
                    continue;
                }
                else if (category.StartsWith("SAR-") || category == "SAR" || category.StartsWith("SPECIAL ACCESS REQUIRED-") || IsNatoSap(category.Split('/')[0], false))
                {
                    kind = 2;
                    ParseSap(category, m, result);
                }
                else if (IsAea(category))
                {
                    kind = 3;
                    ParseAea(category, m, result);
                }
                else if (category == "FGI" || category.StartsWith("FGI ") || category == "FOREIGN GOVERNMENT INFORMATION" || category.StartsWith("FOREIGN GOVERNMENT INFORMATION "))
                {
                    kind = 4;
                    ParseFgi(category, m, result);
                }
                else if (IsDissemination(category))
                {
                    kind = 5;
                    ParseDissemination(category, m, result);
                }
                else if (IsNonIcDissemination(category))
                {
                    kind = 6;
                    ParseNonIcDissemination(category, m, result);
                }
                else
                {
                    kind = 1;
                    ParseSci(category, m, result);
                }
                if (kind < lastCategory)
                {
                    result.Errors.Add("\"" + category + "\" is out of order: categories go classification, SCI, SAP, AEA, FGI, dissemination controls, non-IC dissemination controls.");
                }
                else if (kind == lastCategory)
                {
                    result.Errors.Add("\"" + category + "\" repeats a category: values of one category are separated by a single \"/\".");
                }
                lastCategory = kind;
            }
            return result;
        }

        public static Level? ParseLevel(string text)
        {
            foreach (Level level in new Level[] { Level.TopSecret, Level.Secret, Level.Confidential, Level.Restricted, Level.Unclassified })
            {
                if (text == MarkingFormatter.LevelTitle(level) || text == MarkingFormatter.LevelPortion(level))
                {
                    return level;
                }
            }
            return null;
        }

        private static bool ParseUsClassification(string text, Marking m)
        {
            Level? level = ParseLevel(text.Trim());
            if (level == null || level == Level.Restricted)
            {
                return false; //RESTRICTED is not a US classification
            }
            m.Owner = Owner.UnitedStates;
            m.Level = level.Value;
            return true;
        }

        private static bool ParseForeignClassification(string text, Marking m, ParseResult result)
        {
            text = text.Trim();
            if (ParseNato(text, m))
            {
                return true;
            }
            string[] words = text.Split(' ');
            if (words[0] == "JOINT")
            {
                //"JOINT SECRET CAN GBR USA" or "JOINT S CAN GBR USA"
                m.Owner = Owner.Joint;
                int start = 1;
                Level? level = null;
                if (words.Length > 2 && words[1] == "TOP" && words[2] == "SECRET")
                {
                    level = Level.TopSecret;
                    start = 3;
                }
                else if (words.Length > 1)
                {
                    level = ParseLevel(words[1]);
                    start = 2;
                }
                if (level == null || level == Level.Restricted)
                {
                    result.Errors.Add("\"" + text + "\" has no valid JOINT classification level.");
                    return false;
                }
                m.Level = level.Value;
                for (int i = start; i < words.Length; i++)
                {
                    m.JointOwners.Add(words[i]); //none in "(//JOINT S//REL)": the co-owners of the banner line
                }
                if (m.JointOwners.Count > 0 && !m.JointOwners.Contains("USA"))
                {
                    result.Errors.Add("The co-owners of JOINT information include USA.");
                }
                return true;
            }
            //"GBR SECRET", "GBR S", "CAN GBR S" for several sources, "FGI S" for a concealed foreign source
            for (int count = 1; count < words.Length; count++)
            {
                Level? level = ParseLevel(string.Join(" ", words, count, words.Length - count));
                if (level != null)
                {
                    m.Owner = Owner.NonUS;
                    m.Level = level.Value;
                    string owners = string.Join(" ", words, 0, count);
                    if (owners == "FGI")
                    {
                        m.OwnerCode = null;
                    }
                    else
                    {
                        for (int i = 0; i < count; i++)
                        {
                            if (!IsCountryCode(words[i]))
                            {
                                result.Errors.Add("\"" + words[i] + "\" is not a trigraph or tetragraph code.");
                            }
                        }
                        m.OwnerCode = owners;
                    }
                    return true;
                }
            }
            result.Errors.Add("\"//" + text + "\" is not a non-US, NATO or JOINT classification marking.");
            return false;
        }

        private static bool ParseNato(string text, Marking m)
        {
            Level? level = null;
            string activity = null;
            if (text == "COSMIC TOP SECRET" || text == "CTS")
            {
                level = Level.TopSecret;
            }
            else if (text.StartsWith("NATO "))
            {
                //"NATO SECRET" or "NATO [NAC activity] SECRET"
                foreach (Level candidate in new Level[] { Level.Secret, Level.Confidential, Level.Restricted, Level.Unclassified })
                {
                    string title = MarkingFormatter.LevelTitle(candidate);
                    if (text.EndsWith(" " + title))
                    {
                        level = candidate;
                        string middle = text.Substring(5, text.Length - 5 - title.Length).Trim();
                        activity = middle.Length > 0 ? middle : null;
                        break;
                    }
                }
            }
            else if (text == "NS" || text == "NC" || text == "NR" || text == "NU")
            {
                level = ParseLevel(text.Substring(1));
            }
            else if (text.Length >= 3 && text.IndexOf(' ') < 0 && text[0] == 'N' && "SCRU".IndexOf(text[text.Length - 1]) >= 0)
            {
                //NATO/NAC portion marks: "N[NAC Activity]S", NIS for NATO ISAF SECRET
                level = ParseLevel(text.Substring(text.Length - 1));
                activity = Register.NacActivity(text.Substring(1, text.Length - 2));
            }
            if (level == null)
            {
                return false;
            }
            m.Owner = Owner.Nato;
            m.Level = level.Value;
            m.NatoActivity = activity;
            return true;
        }

        //"123/SI-G ABCD DEFG-MMM AACD/TK"
        private static void ParseSci(string category, Marking m, ParseResult result)
        {
            foreach (string item in category.Split('/'))
            {
                ControlSystem system = ParseHierarchy(item, result, "SCI control system", false);
                if (system == null)
                {
                    continue;
                }
                SciEntry entry = Register.FindSci(system.Name);
                if (entry != null)
                {
                    system.Name = entry.Abbreviation;
                    foreach (Compartment compartment in system.Compartments)
                    {
                        SciEntry child = Register.FindSci(entry.Children, compartment.Name);
                        if (child != null)
                        {
                            compartment.Name = child.Abbreviation;
                        }
                    }
                }
                m.Sci.Add(system);
            }
        }

        //"NAME-COMP1 SUB1 SUB2-COMP2": hyphens between compartments, spaces before sub-compartments
        private static ControlSystem ParseHierarchy(string item, ParseResult result, string what, bool spacesInName)
        {
            string trimmed = item.Trim();
            if (trimmed.Length == 0)
            {
                result.Errors.Add("Empty " + what + ".");
                return null;
            }
            //spelled out titles contain spaces (TALENT KEYHOLE): match published titles first
            string name = null;
            string rest = trimmed;
            foreach (SciEntry entry in Register.Sci)
            {
                if (trimmed == entry.Title || trimmed.StartsWith(entry.Title + "-"))
                {
                    name = entry.Title;
                    rest = trimmed.Substring(name.Length);
                    break;
                }
            }
            if (name == null)
            {
                int hyphen = trimmed.IndexOf('-');
                name = hyphen < 0 ? trimmed : trimmed.Substring(0, hyphen);
                rest = hyphen < 0 ? "" : trimmed.Substring(hyphen);
            }
            if (name.IndexOf(' ') >= 0 && !spacesInName && Register.FindSci(name) == null) //SAP program identifiers are nicknames and may be spelled out: BUTTER POPCORN
            {
                result.Errors.Add("\"" + name + "\": a " + what + " is joined to its compartments by a hyphen.");
            }
            ControlSystem system = new ControlSystem(name);
            if (rest.Length > 0)
            {
                foreach (string part in rest.Substring(1).Split('-'))
                {
                    string compartmentText = part.Trim();
                    if (compartmentText.Length == 0)
                    {
                        result.Errors.Add("Empty compartment in \"" + trimmed + "\".");
                        continue;
                    }
                    //published compartment titles may contain spaces too (none do today), so split on spaces for the sub-compartments
                    string[] words = compartmentText.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    Compartment compartment = new Compartment(words[0]);
                    for (int i = 1; i < words.Length; i++)
                    {
                        compartment.SubCompartments.Add(words[i]);
                    }
                    system.Compartments.Add(compartment);
                }
            }
            return system;
        }

        //"SAR-ABC-DEF 123/SDA-121", or "SAR" alone when the program identifiers are not shown
        private static void ParseSap(string category, Marking m, ParseResult result)
        {
            if (category == "SAR")
            {
                m.SapProgramsHidden = true;
                return;
            }
            string body = category;
            if (category.StartsWith("SAR-"))
            {
                body = category.Substring(4);
            }
            else if (category.StartsWith("SPECIAL ACCESS REQUIRED-"))
            {
                body = category.Substring("SPECIAL ACCESS REQUIRED-".Length);
            }
            foreach (string item in body.Split('/'))
            {
                string programText = item.Trim();
                if (IsNatoSap(programText, false))
                {
                    m.NatoSap.Add(programText);
                    continue;
                }
                if (programText == "SAR")
                {
                    m.SapProgramsHidden = true;
                    continue;
                }
                if (programText.StartsWith("SAR-") || programText.StartsWith("SPECIAL ACCESS REQUIRED-"))
                {
                    result.Errors.Add("The SAR- category indicator is not repeated for each program: \"" + category + "\".");
                    programText = programText.Substring(programText.IndexOf('-') + 1);
                }
                ControlSystem program = ParseHierarchy(programText, result, "SAP program identifier", true);
                if (program != null)
                {
                    m.Sap.Add(program);
                }
            }
        }

        //BALK and BOHEMIA go in the SAP category, ATOMAL in the AEA category
        private static bool IsNatoSap(string text, bool atomal)
        {
            string t = text.Trim();
            return atomal ? t == "ATOMAL" : (t == "BALK" || t == "BOHEMIA");
        }

        private static readonly string[] AeaStarts = new string[]
        {
            "ATOMAL",
            "RD", "RESTRICTED DATA", "FRD", "FORMERLY RESTRICTED DATA", "DCNI", "UCNI", "DOD UCNI", "DOE UCNI",
            "DOD UNCLASSIFIED CONTROLLED NUCLEAR INFORMATION", "DOE UNCLASSIFIED CONTROLLED NUCLEAR INFORMATION",
            "TFNI", "TRANSCLASSIFIED FOREIGN NUCLEAR INFORMATION"
        };

        private static bool IsAea(string category)
        {
            string first = category.Split('/')[0].Split('-')[0].Trim();
            return Array.IndexOf(AeaStarts, first) >= 0;
        }

        //"RD-CNWDI-SG 1 2/FRD-SIGMA 14/DOD UCNI"
        private static void ParseAea(string category, Marking m, ParseResult result)
        {
            foreach (string item in category.Split('/'))
            {
                string[] parts = item.Trim().Split('-');
                string head = parts[0].Trim();
                if (IsNatoSap(head, true))
                {
                    m.NatoSap.Add(head);
                }
                else if (head == "RD" || head == "RESTRICTED DATA")
                {
                    m.RestrictedData = true;
                    for (int i = 1; i < parts.Length; i++)
                    {
                        string part = parts[i].Trim();
                        if (part == "CNWDI" || part == "CRITICAL NUCLEAR WEAPON DESIGN INFORMATION")
                        {
                            m.Cnwdi = true;
                        }
                        else if (!ParseSigma(part, m.RdSigma))
                        {
                            result.Errors.Add("\"" + part + "\" is not a sub-marking of RESTRICTED DATA (CNWDI, SIGMA [#]).");
                        }
                    }
                }
                else if (head == "FRD" || head == "FORMERLY RESTRICTED DATA")
                {
                    m.FormerlyRestrictedData = true;
                    for (int i = 1; i < parts.Length; i++)
                    {
                        if (!ParseSigma(parts[i].Trim(), m.FrdSigma))
                        {
                            result.Errors.Add("\"" + parts[i] + "\" is not a sub-marking of FORMERLY RESTRICTED DATA (SIGMA [#]).");
                        }
                    }
                }
                else if (head == "DCNI" || head == "DOD UCNI" || head == "DOD UNCLASSIFIED CONTROLLED NUCLEAR INFORMATION")
                {
                    m.DodUcni = true;
                }
                else if (head == "UCNI" || head == "DOE UCNI" || head == "DOE UNCLASSIFIED CONTROLLED NUCLEAR INFORMATION")
                {
                    m.DoeUcni = true;
                }
                else if (head == "TFNI" || head == "TRANSCLASSIFIED FOREIGN NUCLEAR INFORMATION")
                {
                    m.Tfni = true;
                }
                else
                {
                    result.Errors.Add("\"" + item + "\" is not an Atomic Energy Act marking.");
                }
            }
        }

        //"SIGMA 1 2" or "SG 1 2"
        private static bool ParseSigma(string text, List<int> numbers)
        {
            string[] words = text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length < 2 || (words[0] != "SIGMA" && words[0] != "SG"))
            {
                return false;
            }
            for (int i = 1; i < words.Length; i++)
            {
                int number;
                if (!Sorting.IsAllDigits(words[i]) || !int.TryParse(words[i], out number))
                {
                    return false;
                }
                numbers.Add(number);
            }
            return true;
        }

        //"FGI GBR JPN NATO" or "FGI"
        private static void ParseFgi(string category, Marking m, ParseResult result)
        {
            m.Fgi = true;
            string list = category.StartsWith("FOREIGN GOVERNMENT INFORMATION") ? category.Substring("FOREIGN GOVERNMENT INFORMATION".Length) : category.Substring(3);
            foreach (string code in list.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!IsCountryCode(code))
                {
                    result.Errors.Add("\"" + code + "\" is not a trigraph or tetragraph code in the FGI marking.");
                }
                m.FgiSources.Add(code);
            }
        }

        public static bool IsCountryCode(string code)
        {
            if (code.Length < 3 || code.Length > 4)
            {
                return false;
            }
            foreach (char c in code)
            {
                if (c < 'A' || c > 'Z')
                {
                    return false;
                }
            }
            return true;
        }

        //splits a dissemination category on "/", keeping "USA/GBR EYES ONLY" in one piece
        private static List<string> SplitDissemination(string category)
        {
            List<string> items = new List<string>();
            List<string> pendingCodes = new List<string>();
            foreach (string piece in category.Split('/'))
            {
                string item = piece.Trim();
                if (item.EndsWith(" EYES ONLY"))
                {
                    pendingCodes.Add(item);
                    items.Add(string.Join("/", pendingCodes.ToArray()));
                    pendingCodes.Clear();
                }
                else if (IsCountryCode(item) && Register.Recognize(Register.Dissemination, item) == null && Register.Recognize(Register.NonIcDissemination, item) == null)
                {
                    pendingCodes.Add(item); //a country of a following EYES ONLY list
                }
                else
                {
                    items.AddRange(pendingCodes);
                    pendingCodes.Clear();
                    items.Add(item);
                }
            }
            items.AddRange(pendingCodes);
            return items;
        }

        private static bool IsDisseminationItem(string item)
        {
            return item.StartsWith("REL TO ") || item == "REL" || item.EndsWith(" EYES ONLY") || item.StartsWith("DISPLAY ONLY ")
                || item.StartsWith("AUTHORIZED FOR RELEASE TO ") || Register.Recognize(Register.Dissemination, item) != null;
        }

        private static bool IsDissemination(string category)
        {
            return IsDisseminationItem(SplitDissemination(category)[0]);
        }

        private static bool IsNonIcDissemination(string category)
        {
            return Register.Recognize(Register.NonIcDissemination, category.Split('/')[0].Trim()) != null;
        }

        private static void ParseDissemination(string category, Marking m, ParseResult result)
        {
            foreach (string item in SplitDissemination(category))
            {
                if (item.StartsWith("REL TO ") || item.StartsWith("AUTHORIZED FOR RELEASE TO "))
                {
                    AddKey(m.Dissem, "REL", result);
                    string list = item.Substring(item.IndexOf(" TO ") + 4);
                    foreach (string code in list.Split(','))
                    {
                        string trimmed = code.Trim();
                        if (!IsCountryCode(trimmed))
                        {
                            result.Errors.Add("\"" + trimmed + "\" is not a trigraph or tetragraph code in REL TO; codes are separated by \", \".");
                        }
                        m.RelTo.Add(trimmed);
                    }
                    if (m.RelTo.Count == 0 || m.RelTo[0] != "USA")
                    {
                        result.Errors.Add("USA must come first in the REL TO list.");
                    }
                }
                else if (item == "REL")
                {
                    AddKey(m.Dissem, "REL", result); //the list is the one of the banner line
                }
                else if (item.EndsWith(" EYES ONLY"))
                {
                    AddKey(m.Dissem, "EYES", result);
                    foreach (string code in item.Substring(0, item.Length - " EYES ONLY".Length).Split('/'))
                    {
                        m.EyesOnly.Add(code.Trim());
                    }
                }
                else if (item.StartsWith("DISPLAY ONLY "))
                {
                    AddKey(m.Dissem, "DISPLAYONLY", result);
                    foreach (string code in item.Substring("DISPLAY ONLY ".Length).Split(','))
                    {
                        m.DisplayOnly.Add(code.Trim());
                    }
                }
                else
                {
                    RegisterEntry entry = Register.Recognize(Register.Dissemination, item);
                    if (entry == null)
                    {
                        result.Errors.Add("\"" + item + "\" is not a dissemination control marking of the Register.");
                    }
                    else
                    {
                        AddKey(m.Dissem, entry.Key, result);
                    }
                }
            }
            CheckRegisterOrder(m.Dissem, Register.Dissemination, result);
        }

        private static void ParseNonIcDissemination(string category, Marking m, ParseResult result)
        {
            foreach (string piece in category.Split('/'))
            {
                string item = piece.Trim();
                RegisterEntry entry = Register.Recognize(Register.NonIcDissemination, item);
                if (entry == null)
                {
                    result.Errors.Add("\"" + item + "\" is not a non-IC dissemination control marking of the Register.");
                }
                else
                {
                    AddKey(m.NonIcDissem, entry.Key, result);
                }
            }
            CheckRegisterOrder(m.NonIcDissem, Register.NonIcDissemination, result);
        }

        private static void AddKey(List<string> keys, string key, ParseResult result)
        {
            if (keys.Contains(key))
            {
                result.Errors.Add("\"" + key + "\" appears twice.");
                return;
            }
            keys.Add(key);
        }

        private static void CheckRegisterOrder(List<string> keys, RegisterEntry[] entries, ParseResult result)
        {
            for (int i = 1; i < keys.Count; i++)
            {
                if (Register.OrderOf(entries, keys[i]) < Register.OrderOf(entries, keys[i - 1]))
                {
                    result.Errors.Add("\"" + keys[i] + "\" must come before \"" + keys[i - 1] + "\": controls are listed in Register order.");
                }
            }
        }
    }
}
