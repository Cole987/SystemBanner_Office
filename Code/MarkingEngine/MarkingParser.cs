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
                result.Errors.Add(Text.T("The marking is empty.", "Le marquage est vide.", "Die Kennzeichnung ist leer."));
                return result;
            }
            if (value != value.ToUpperInvariant())
            {
                result.Errors.Add(Text.T("Markings must be in uppercase letters.", "Les marquages s'écrivent en majuscules.", "Kennzeichnungen werden in Großbuchstaben geschrieben."));
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
                result.Errors.Add(Text.T("\"" + categories[0] + "\" is not a classification marking.", "« " + categories[0] + " » n'est pas un marquage de classification.", "\"" + categories[0] + "\" ist keine Kennzeichnung eines Geheimhaltungsgrads."));
                return result;
            }
            int lastCategory = 0; //categories must come in Register order: SCI 1, SAP 2, AEA 3, FGI 4, dissemination 5, non-IC dissemination 6
            for (int i = 1; i < categories.Count; i++)
            {
                string category = categories[i].Trim();
                int kind;
                if (category.Length == 0)
                {
                    result.Errors.Add(Text.T("Empty category: no \"////\" placeholder is allowed for categories that do not apply.", "Catégorie vide : aucun « //// » n'est admis pour les catégories sans objet.", "Leere Kategorie: Für nicht zutreffende Kategorien ist kein Platzhalter \"////\" zulässig."));
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
                    result.Errors.Add(Text.T("\"" + category + "\" is out of order: categories go classification, SCI, SAP, AEA, FGI, dissemination controls, non-IC dissemination controls.", "« " + category + " » n'est pas à sa place : l'ordre des catégories est classification, SCI, SAP, AEA, FGI, contrôles de diffusion, contrôles de diffusion hors IC.", "\"" + category + "\" steht an falscher Stelle: Die Reihenfolge der Kategorien ist Geheimhaltungsgrad, SCI, SAP, AEA, FGI, Weitergabebeschränkungen, Weitergabebeschränkungen außerhalb der IC."));
                }
                else if (kind == lastCategory)
                {
                    result.Errors.Add(Text.T("\"" + category + "\" repeats a category: values of one category are separated by a single \"/\".", "« " + category + " » répète une catégorie : les valeurs d'une même catégorie sont séparées par un seul « / ».", "\"" + category + "\" wiederholt eine Kategorie: Werte einer Kategorie werden durch ein einzelnes \"/\" getrennt."));
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
                    result.Errors.Add(Text.T("\"" + text + "\" has no valid JOINT classification level.", "« " + text + " » n'a pas de niveau de classification JOINT valide.", "\"" + text + "\" hat keinen gültigen JOINT-Geheimhaltungsgrad."));
                    return false;
                }
                m.Level = level.Value;
                for (int i = start; i < words.Length; i++)
                {
                    m.JointOwners.Add(words[i]); //none in "(//JOINT S//REL)": the co-owners of the banner line
                }
                if (m.JointOwners.Count > 0 && !m.JointOwners.Contains("USA"))
                {
                    result.Errors.Add(Text.T("The co-owners of JOINT information include USA.", "Les copropriétaires d'une information JOINT incluent USA.", "Zu den Miteigentümern von JOINT-Informationen gehört USA."));
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
                                result.Errors.Add(Text.T("\"" + words[i] + "\" is not a trigraph or tetragraph code.", "« " + words[i] + " » n'est pas un code trigraphe ou tétragraphe.", "\"" + words[i] + "\" ist kein Trigraph- oder Tetragraph-Code."));
                            }
                        }
                        m.OwnerCode = owners;
                    }
                    return true;
                }
            }
            result.Errors.Add(Text.T("\"//" + text + "\" is not a non-US, NATO or JOINT classification marking.", "« //" + text + " » n'est pas un marquage de classification non américain, OTAN ou JOINT.", "\"//" + text + "\" ist keine nicht-amerikanische, NATO- oder JOINT-Kennzeichnung eines Geheimhaltungsgrads."));
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
                ControlSystem system = ParseHierarchy(item, result, Text.T("SCI control system", "système de contrôle SCI", "SCI-Kontrollsystem"), false);
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
                result.Errors.Add(Text.T("Empty " + what + ".", "Valeur vide : " + what + ".", "Leerer Wert: " + what + "."));
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
                result.Errors.Add(Text.T("\"" + name + "\": a " + what + " is joined to its compartments by a hyphen.", "« " + name + " » : un " + what + " est relié à ses compartiments par un trait d'union.", "\"" + name + "\": " + what + " und Compartments werden durch einen Bindestrich verbunden."));
            }
            ControlSystem system = new ControlSystem(name);
            if (rest.Length > 0)
            {
                foreach (string part in rest.Substring(1).Split('-'))
                {
                    string compartmentText = part.Trim();
                    if (compartmentText.Length == 0)
                    {
                        result.Errors.Add(Text.T("Empty compartment in \"" + trimmed + "\".", "Compartiment vide dans « " + trimmed + " ».", "Leeres Compartment in \"" + trimmed + "\"."));
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
                    result.Errors.Add(Text.T("The SAR- category indicator is not repeated for each program: \"" + category + "\".", "L'indicateur de catégorie SAR- ne se répète pas pour chaque programme : « " + category + " ».", "Das Kategoriekennzeichen SAR- wird nicht für jedes Programm wiederholt: \"" + category + "\"."));
                    programText = programText.Substring(programText.IndexOf('-') + 1);
                }
                ControlSystem program = ParseHierarchy(programText, result, Text.T("SAP program identifier", "identifiant de programme SAP", "SAP-Programmkennung"), true);
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
                            result.Errors.Add(Text.T("\"" + part + "\" is not a sub-marking of RESTRICTED DATA (CNWDI, SIGMA [#]).", "« " + part + " » n'est pas un sous-marquage de RESTRICTED DATA (CNWDI, SIGMA [#]).", "\"" + part + "\" ist keine Unterkennzeichnung von RESTRICTED DATA (CNWDI, SIGMA [#])."));
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
                            result.Errors.Add(Text.T("\"" + parts[i] + "\" is not a sub-marking of FORMERLY RESTRICTED DATA (SIGMA [#]).", "« " + parts[i] + " » n'est pas un sous-marquage de FORMERLY RESTRICTED DATA (SIGMA [#]).", "\"" + parts[i] + "\" ist keine Unterkennzeichnung von FORMERLY RESTRICTED DATA (SIGMA [#])."));
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
                    result.Errors.Add(Text.T("\"" + item + "\" is not an Atomic Energy Act marking.", "« " + item + " » n'est pas un marquage de l'Atomic Energy Act.", "\"" + item + "\" ist keine Kennzeichnung nach dem Atomic Energy Act."));
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
                    result.Errors.Add(Text.T("\"" + code + "\" is not a trigraph or tetragraph code in the FGI marking.", "« " + code + " » n'est pas un code trigraphe ou tétragraphe dans le marquage FGI.", "\"" + code + "\" ist kein Trigraph- oder Tetragraph-Code in der FGI-Kennzeichnung."));
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
                            result.Errors.Add(Text.T("\"" + trimmed + "\" is not a trigraph or tetragraph code in REL TO; codes are separated by \", \".", "« " + trimmed + " » n'est pas un code trigraphe ou tétragraphe dans REL TO ; les codes sont séparés par « , ».", "\"" + trimmed + "\" ist kein Trigraph- oder Tetragraph-Code in REL TO; Codes werden durch \", \" getrennt."));
                        }
                        m.RelTo.Add(trimmed);
                    }
                    if (m.RelTo.Count == 0 || m.RelTo[0] != "USA")
                    {
                        result.Errors.Add(Text.T("USA must come first in the REL TO list.", "USA vient en premier dans la liste REL TO.", "USA steht in der REL TO-Liste an erster Stelle."));
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
                        result.Errors.Add(Text.T("\"" + item + "\" is not a dissemination control marking of the Register.", "« " + item + " » n'est pas un marquage de contrôle de diffusion du Register.", "\"" + item + "\" ist keine Kennzeichnung einer Weitergabebeschränkung aus dem Register."));
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
                    result.Errors.Add(Text.T("\"" + item + "\" is not a non-IC dissemination control marking of the Register.", "« " + item + " » n'est pas un marquage de contrôle de diffusion hors IC du Register.", "\"" + item + "\" ist keine Kennzeichnung einer Weitergabebeschränkung außerhalb der IC aus dem Register."));
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
                result.Errors.Add(Text.T("\"" + key + "\" appears twice.", "« " + key + " » figure deux fois.", "\"" + key + "\" kommt zweimal vor."));
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
                    result.Errors.Add(Text.T("\"" + keys[i] + "\" must come before \"" + keys[i - 1] + "\": controls are listed in Register order.", "« " + keys[i] + " » doit précéder « " + keys[i - 1] + " » : les contrôles suivent l'ordre du Register.", "\"" + keys[i] + "\" muss vor \"" + keys[i - 1] + "\" stehen: Beschränkungen werden in der Reihenfolge des Register aufgeführt."));
                }
            }
        }
    }
}
