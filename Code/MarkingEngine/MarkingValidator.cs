//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;

namespace SystemBannerOffice.Markings
{
    //checks one portion mark or banner line against the templates of the Manual: the levels each marking allows,
    //what it requires, what it may not be combined with, and the length of compartment values
    public static class MarkingValidator
    {
        private const int TS = 1 << (int)Level.TopSecret, S = 1 << (int)Level.Secret, C = 1 << (int)Level.Confidential, U = 1 << (int)Level.Unclassified;
        private const int Classified = TS | S | C, Any = TS | S | C | U;

        //levels allowed for each dissemination and non-IC control, by Register key
        private static readonly Dictionary<string, int> ControlLevels = CreateControlLevels();

        private static Dictionary<string, int> CreateControlLevels()
        {
            Dictionary<string, int> levels = new Dictionary<string, int>();
            levels["RS"] = TS | S;
            levels["FOUO"] = U;
            levels["OC"] = Classified;
            levels["OC-USGOV"] = Classified;
            levels["IMC"] = TS | S;
            levels["NF"] = Any;
            levels["PR"] = Any;
            levels["REL"] = Any;
            levels["RELIDO"] = Any;
            levels["EYES"] = Classified;
            levels["DSEN"] = Any;
            levels["FISA"] = Any;
            levels["DISPLAYONLY"] = Any;
            levels["DS"] = U;
            levels["XD"] = Any;
            levels["ND"] = Any;
            levels["SBU"] = U;
            levels["SBU-NF"] = U;
            levels["LES"] = Any;
            levels["LES-NF"] = Any;
            levels["SSI"] = Any;
            return levels;
        }

        //SIGMA numbers in use; 1 to 5 and 9 to 13 are obsolete and must be converted when the information is reused
        public static readonly int[] ValidSigma = new int[] { 14, 15, 18, 20 };

        public static List<string> Validate(Marking m)
        {
            List<string> errors = new List<string>();
            //RESTRICTED is a non-US and NATO level only; JOINT levels follow the US ones
            if (m.Level == Level.Restricted && (m.Owner == Owner.UnitedStates || m.Owner == Owner.Joint))
            {
                errors.Add("RESTRICTED is not a US classification level.");
            }
            int level = 1 << (int)(m.Level == Level.Restricted ? Level.Confidential : m.Level);
            bool unclassified = m.Level == Level.Unclassified;
            if (unclassified && (m.Sci.Count > 0 || m.Sap.Count > 0 || m.SapProgramsHidden))
            {
                errors.Add("SCI and SAP markings only apply to classified information.");
            }
            CheckSci(m, level, errors);
            CheckAea(m, level, errors);
            CheckDissemination(m, level, errors);
            CheckJoint(m, errors);
            return errors;
        }

        private static string LevelsText(int levels)
        {
            List<string> names = new List<string>();
            if ((levels & TS) != 0) names.Add("TOP SECRET");
            if ((levels & S) != 0) names.Add("SECRET");
            if ((levels & C) != 0) names.Add("CONFIDENTIAL");
            if ((levels & U) != 0) names.Add("UNCLASSIFIED");
            return string.Join(", ", names.ToArray());
        }

        private static void CheckLevel(string marking, int allowed, int level, List<string> errors)
        {
            if ((allowed & level) == 0)
            {
                errors.Add(marking + " is only used with " + LevelsText(allowed) + ".");
            }
        }

        private static ControlSystem FindSystem(Marking m, string name)
        {
            foreach (ControlSystem system in m.Sci)
            {
                if (system.Name == name)
                {
                    return system;
                }
            }
            return null;
        }

        private static void CheckSci(Marking m, int level, List<string> errors)
        {
            bool orcon = m.HasDissem("OC");
            bool noforn = m.HasDissem("NF");
            bool usgov = m.HasDissem("OC-USGOV");
            foreach (ControlSystem system in m.Sci)
            {
                foreach (ControlSystem other in m.Sci)
                {
                    if (other != system && other.Name == system.Name)
                    {
                        errors.Add("SCI control system " + system.Name + " appears twice: list each value once.");
                    }
                }
                foreach (Compartment compartment in system.Compartments)
                {
                    if (compartment.Name == "ECI")
                    {
                        errors.Add("\"ECI\" is an SCI type indicator, not a compartment: write SI-ABC, not SI-ECI ABC.");
                    }
                }
            }
            if (m.Owner == Owner.Joint && FindSystem(m, "HCS") != null)
            {
                errors.Add("HCS markings may not be used on JOINT information.");
            }
            ControlSystem hcs = FindSystem(m, "HCS");
            if (hcs != null)
            {
                CheckLevel("HCS", Classified, level, errors);
                if (hcs.Compartments.Count == 0)
                {
                    errors.Add("HCS requires its compartment, HCS-O or HCS-P; HCS alone is a legacy marking to re-mark when reused.");
                }
                Compartment o = hcs.FindCompartment("O");
                Compartment p = hcs.FindCompartment("P");
                if (o != null)
                {
                    CheckLevel("HCS-O", TS | S, level, errors);
                    if (!orcon || !noforn) errors.Add("HCS-O requires ORCON and NOFORN.");
                    if (usgov) errors.Add("HCS-O may not be used with ORCON-USGOV.");
                }
                if (p != null)
                {
                    CheckLevel("HCS-P", TS | S, level, errors);
                    if (!noforn) errors.Add("HCS-P requires NOFORN.");
                    if (p.SubCompartments.Count > 0)
                    {
                        CheckLevel("HCS-P sub-compartments", TS, level, errors);
                        if (!orcon) errors.Add("HCS-P sub-compartments require ORCON.");
                        if (usgov) errors.Add("HCS-P sub-compartments may not be used with ORCON-USGOV.");
                        CheckLength(p.SubCompartments, 1, 6, false, "HCS-P sub-compartments", errors);
                    }
                }
            }
            ControlSystem rsv = FindSystem(m, "RSV");
            if (rsv != null)
            {
                CheckLevel("RESERVE", TS | S, level, errors);
                if (rsv.Compartments.Count == 0)
                {
                    errors.Add("RESERVE requires a compartment: RSV-[COMPARTMENT].");
                }
                foreach (Compartment compartment in rsv.Compartments)
                {
                    CheckLength(new List<string>(new string[] { compartment.Name }), 3, 3, false, "RESERVE compartments", errors);
                }
            }
            ControlSystem si = FindSystem(m, "SI");
            if (si != null)
            {
                CheckLevel("SI", Classified, level, errors);
                foreach (Compartment compartment in si.Compartments)
                {
                    CheckLevel("SI-" + compartment.Name, TS, level, errors);
                    if (compartment.Name == "G")
                    {
                        if (!orcon) errors.Add("SI-GAMMA requires ORCON.");
                        if (usgov) errors.Add("SI-GAMMA may not be used with ORCON-USGOV.");
                        CheckLength(compartment.SubCompartments, 4, 4, true, "GAMMA sub-compartments", errors);
                    }
                    else if (compartment.Name != "EU" && compartment.Name != "NK")
                    {
                        CheckLength(new List<string>(new string[] { compartment.Name }), 2, 3, true, "SI compartments", errors);
                    }
                }
            }
            ControlSystem tk = FindSystem(m, "TK");
            if (tk != null)
            {
                CheckLevel("TALENT KEYHOLE", TS | S, level, errors);
                foreach (Compartment compartment in tk.Compartments)
                {
                    CheckLevel("TK-" + compartment.Name, compartment.Name == "BLFH" ? TS : TS | S, level, errors);
                    if (!noforn) errors.Add("TK-" + compartment.Name + " requires NOFORN.");
                    CheckLength(compartment.SubCompartments, 1, 6, false, "TK-" + compartment.Name + " sub-compartments", errors);
                }
            }
        }

        private static void CheckLength(List<string> values, int min, int max, bool lettersOnly, string what, List<string> errors)
        {
            foreach (string value in values)
            {
                bool valid = value.Length >= min && value.Length <= max;
                foreach (char c in value)
                {
                    valid &= (c >= 'A' && c <= 'Z') || (!lettersOnly && c >= '0' && c <= '9');
                }
                if (!valid)
                {
                    string length = min == max ? min + "" : (min == 1 ? "up to " + max : min + " to " + max);
                    errors.Add(what + " are " + length + (lettersOnly ? " letters" : " letters or digits") + ": \"" + value + "\".");
                }
            }
        }

        private static void CheckAea(Marking m, int level, List<string> errors)
        {
            if (m.RestrictedData) CheckLevel("RESTRICTED DATA", Classified, level, errors);
            if (m.FormerlyRestrictedData) CheckLevel("FORMERLY RESTRICTED DATA", Classified, level, errors);
            if (m.Tfni) CheckLevel("TFNI", Classified, level, errors);
            if (m.Cnwdi)
            {
                if (!m.RestrictedData) errors.Add("CNWDI requires RESTRICTED DATA: RD-CNWDI.");
                CheckLevel("CNWDI", TS | S, level, errors);
            }
            if (m.RdSigma.Count > 0 || m.FrdSigma.Count > 0)
            {
                CheckLevel("SIGMA", TS | S, level, errors);
                foreach (int number in Join(m.RdSigma, m.FrdSigma))
                {
                    if (Array.IndexOf(ValidSigma, number) < 0)
                    {
                        errors.Add("SIGMA " + number + " is obsolete or not valid: SIGMA numbers are 14, 15, 18 and 20.");
                    }
                }
                if (m.HasDissem("REL"))
                {
                    errors.Add("SIGMA information may not share a portion with REL TO unless an equivalent release decision exists.");
                }
            }
            if (m.DodUcni) CheckLevel("DOD UCNI", U, level, errors);
            if (m.DoeUcni) CheckLevel("DOE UCNI", U, level, errors);
            if ((m.RestrictedData || m.FormerlyRestrictedData) && !m.HasDissem("NF"))
            {
                errors.Add("RD and FRD are used with NOFORN unless an Atomic Energy Act sharing agreement applies: check the release.");
            }
        }

        private static List<int> Join(List<int> a, List<int> b)
        {
            List<int> all = new List<int>(a);
            all.AddRange(b);
            return all;
        }

        private static void CheckDissemination(Marking m, int level, List<string> errors)
        {
            foreach (string key in m.Dissem)
            {
                RegisterEntry entry = Register.FindDissemination(key);
                int allowed;
                if (entry != null && ControlLevels.TryGetValue(key, out allowed))
                {
                    CheckLevel(entry.BannerText(true), allowed, level, errors);
                }
            }
            foreach (string key in m.NonIcDissem)
            {
                RegisterEntry entry = Register.FindNonIcDissemination(key);
                int allowed;
                if (entry != null && ControlLevels.TryGetValue(key, out allowed))
                {
                    CheckLevel(entry.BannerText(true), allowed, level, errors);
                }
            }
            bool nf = m.HasDissem("NF");
            if (nf && (m.HasDissem("REL") || m.HasDissem("RELIDO") || m.HasDissem("EYES") || m.HasDissem("DISPLAYONLY")))
            {
                errors.Add("NOFORN may not be used with REL TO, RELIDO, EYES ONLY or DISPLAY ONLY.");
            }
            if (m.HasDissem("REL") && m.HasDissem("EYES"))
            {
                errors.Add("REL TO may not be used with EYES ONLY.");
            }
            if (m.HasDissem("RELIDO") && m.HasDissem("DISPLAYONLY"))
            {
                errors.Add("RELIDO may not be used with DISPLAY ONLY.");
            }
            if (m.HasDissem("OC") && m.HasDissem("OC-USGOV"))
            {
                errors.Add("ORCON and ORCON-USGOV may not be used together: ORCON supersedes ORCON-USGOV.");
            }
            if (m.HasDissem("RELIDO") && (m.HasDissem("OC") || m.HasDissem("OC-USGOV")))
            {
                errors.Add("RELIDO may not be used with ORCON or ORCON-USGOV.");
            }
            if (m.HasNonIcDissem("XD") && m.HasNonIcDissem("ND"))
            {
                errors.Add("EXDIS and NODIS may not be used together: NODIS supersedes EXDIS.");
            }
            if ((m.HasNonIcDissem("XD") || m.HasNonIcDissem("ND")) && !nf)
            {
                errors.Add("EXDIS and NODIS require NOFORN.");
            }
            if (m.HasNonIcDissem("DS") && m.HasDissem("FOUO"))
            {
                errors.Add("LIMDIS information may not be combined with FOUO information in a portion.");
            }
            if (m.HasDissem("REL"))
            {
                bool other = false;
                foreach (string code in m.RelTo)
                {
                    other |= code != "USA";
                }
                if (m.RelTo.Count > 0 && !other)
                {
                    errors.Add("\"REL TO USA\" alone is not authorized: name at least one other country or organization.");
                }
            }
            if (m.HasDissem("EYES"))
            {
                foreach (string code in m.EyesOnly)
                {
                    if (code != "USA" && code != "AUS" && code != "CAN" && code != "GBR" && code != "NZL")
                    {
                        errors.Add("EYES ONLY is only for Second Party partners (AUS, CAN, GBR, NZL): \"" + code + "\".");
                    }
                }
            }
            if (m.HasDissem("DISPLAYONLY") && m.DisplayOnly.Contains("USA"))
            {
                errors.Add("The DISPLAY ONLY list names foreign countries only, not USA.");
            }
        }

        private static void CheckJoint(Marking m, List<string> errors)
        {
            if (m.Owner != Owner.Joint)
            {
                return;
            }
            if (m.HasDissem("NF"))
            {
                errors.Add("JOINT information may not be NOFORN.");
            }
            if (!m.HasDissem("REL"))
            {
                errors.Add("JOINT information requires REL TO, to at least all its co-owners.");
                return;
            }
            if (m.RelTo.Count == 0)
            {
                return; //"REL": the list of the banner line
            }
            List<string> releasable = RollUp.Expand(m.RelTo);
            foreach (string owner in m.JointOwners)
            {
                if (!releasable.Contains(owner))
                {
                    errors.Add("JOINT information must be releasable to its co-owner " + owner + ".");
                }
            }
        }
    }
}
