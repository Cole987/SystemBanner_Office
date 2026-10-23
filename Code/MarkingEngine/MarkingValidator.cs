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
                errors.Add(Text.T("RESTRICTED is not a US classification level.", "RESTRICTED n'est pas un niveau de classification américain."));
            }
            int level = 1 << (int)(m.Level == Level.Restricted ? Level.Confidential : m.Level);
            bool unclassified = m.Level == Level.Unclassified;
            if (unclassified && (m.Sci.Count > 0 || m.Sap.Count > 0 || m.SapProgramsHidden))
            {
                errors.Add(Text.T("SCI and SAP markings only apply to classified information.", "Les marquages SCI et SAP ne s'appliquent qu'aux informations classifiées."));
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
                errors.Add(Text.T(marking + " is only used with " + LevelsText(allowed) + ".", marking + " : uniquement avec " + LevelsText(allowed) + "."));
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
                        errors.Add(Text.T("SCI control system " + system.Name + " appears twice: list each value once.", "Le système de contrôle SCI " + system.Name + " figure deux fois : citer chaque valeur une seule fois."));
                    }
                }
                foreach (Compartment compartment in system.Compartments)
                {
                    if (compartment.Name == "ECI")
                    {
                        errors.Add(Text.T("\"ECI\" is an SCI type indicator, not a compartment: write SI-ABC, not SI-ECI ABC.", "« ECI » est un indicateur de type SCI, pas un compartiment : écrire SI-ABC, et non SI-ECI ABC."));
                    }
                }
            }
            if (m.Owner == Owner.Joint && FindSystem(m, "HCS") != null)
            {
                errors.Add(Text.T("HCS markings may not be used on JOINT information.", "Les marquages HCS ne s'emploient pas sur une information JOINT."));
            }
            ControlSystem hcs = FindSystem(m, "HCS");
            if (hcs != null)
            {
                CheckLevel("HCS", Classified, level, errors);
                if (hcs.Compartments.Count == 0)
                {
                    errors.Add(Text.T("HCS requires its compartment, HCS-O or HCS-P; HCS alone is a legacy marking to re-mark when reused.", "HCS exige son compartiment, HCS-O ou HCS-P ; HCS seul est un marquage ancien, à reprendre en cas de réutilisation."));
                }
                Compartment o = hcs.FindCompartment("O");
                Compartment p = hcs.FindCompartment("P");
                if (o != null)
                {
                    CheckLevel("HCS-O", TS | S, level, errors);
                    if (!orcon || !noforn) errors.Add(Text.T("HCS-O requires ORCON and NOFORN.", "HCS-O exige ORCON et NOFORN."));
                    if (usgov) errors.Add(Text.T("HCS-O may not be used with ORCON-USGOV.", "HCS-O ne s'emploie pas avec ORCON-USGOV."));
                }
                if (p != null)
                {
                    CheckLevel("HCS-P", TS | S, level, errors);
                    if (!noforn) errors.Add(Text.T("HCS-P requires NOFORN.", "HCS-P exige NOFORN."));
                    if (p.SubCompartments.Count > 0)
                    {
                        CheckLevel(Text.T("HCS-P sub-compartments", "Sous-compartiments HCS-P"), TS, level, errors);
                        if (!orcon) errors.Add(Text.T("HCS-P sub-compartments require ORCON.", "Les sous-compartiments HCS-P exigent ORCON."));
                        if (usgov) errors.Add(Text.T("HCS-P sub-compartments may not be used with ORCON-USGOV.", "Les sous-compartiments HCS-P ne s'emploient pas avec ORCON-USGOV."));
                        CheckLength(p.SubCompartments, 1, 6, false, Text.T("HCS-P sub-compartments", "Sous-compartiments HCS-P"), errors);
                    }
                }
            }
            ControlSystem rsv = FindSystem(m, "RSV");
            if (rsv != null)
            {
                CheckLevel("RESERVE", TS | S, level, errors);
                if (rsv.Compartments.Count == 0)
                {
                    errors.Add(Text.T("RESERVE requires a compartment: RSV-[COMPARTMENT].", "RESERVE exige un compartiment : RSV-[COMPARTIMENT]."));
                }
                foreach (Compartment compartment in rsv.Compartments)
                {
                    CheckLength(new List<string>(new string[] { compartment.Name }), 3, 3, false, Text.T("RESERVE compartments", "Compartiments RESERVE"), errors);
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
                        if (!orcon) errors.Add(Text.T("SI-GAMMA requires ORCON.", "SI-GAMMA exige ORCON."));
                        if (usgov) errors.Add(Text.T("SI-GAMMA may not be used with ORCON-USGOV.", "SI-GAMMA ne s'emploie pas avec ORCON-USGOV."));
                        CheckLength(compartment.SubCompartments, 4, 4, true, Text.T("GAMMA sub-compartments", "Sous-compartiments GAMMA"), errors);
                    }
                    else if (compartment.Name != "EU" && compartment.Name != "NK")
                    {
                        CheckLength(new List<string>(new string[] { compartment.Name }), 2, 3, true, Text.T("SI compartments", "Compartiments SI"), errors);
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
                    if (!noforn) errors.Add(Text.T("TK-" + compartment.Name + " requires NOFORN.", "TK-" + compartment.Name + " exige NOFORN."));
                    CheckLength(compartment.SubCompartments, 1, 6, false, Text.T("TK-" + compartment.Name + " sub-compartments", "Sous-compartiments TK-" + compartment.Name), errors);
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
                    string longueur = min == max ? min + "" : (min == 1 ? "jusqu'à " + max : "de " + min + " à " + max);
                    errors.Add(Text.T(what + " are " + length + (lettersOnly ? " letters" : " letters or digits") + ": \"" + value + "\".", what + " : " + longueur + (lettersOnly ? " lettres" : " lettres ou chiffres") + ", « " + value + " »."));
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
                if (!m.RestrictedData) errors.Add(Text.T("CNWDI requires RESTRICTED DATA: RD-CNWDI.", "CNWDI exige RESTRICTED DATA : RD-CNWDI."));
                CheckLevel("CNWDI", TS | S, level, errors);
            }
            if (m.RdSigma.Count > 0 || m.FrdSigma.Count > 0)
            {
                CheckLevel("SIGMA", TS | S, level, errors);
                foreach (int number in Join(m.RdSigma, m.FrdSigma))
                {
                    if (Array.IndexOf(ValidSigma, number) < 0)
                    {
                        errors.Add(Text.T("SIGMA " + number + " is obsolete or not valid: SIGMA numbers are 14, 15, 18 and 20.", "SIGMA " + number + " est obsolète ou non valide : les numéros SIGMA sont 14, 15, 18 et 20."));
                    }
                }
                if (m.HasDissem("REL"))
                {
                    errors.Add(Text.T("SIGMA information may not share a portion with REL TO unless an equivalent release decision exists.", "Une information SIGMA ne partage pas une portion avec REL TO, sauf décision de diffusion équivalente."));
                }
            }
            if (m.DodUcni) CheckLevel("DOD UCNI", U, level, errors);
            if (m.DoeUcni) CheckLevel("DOE UCNI", U, level, errors);
            if ((m.RestrictedData || m.FormerlyRestrictedData) && !m.HasDissem("NF"))
            {
                errors.Add(Text.T("RD and FRD are used with NOFORN unless an Atomic Energy Act sharing agreement applies: check the release.", "RD et FRD s'emploient avec NOFORN, sauf accord de partage au titre de l'Atomic Energy Act : vérifier la diffusion."));
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
                errors.Add(Text.T("NOFORN may not be used with REL TO, RELIDO, EYES ONLY or DISPLAY ONLY.", "NOFORN ne s'emploie pas avec REL TO, RELIDO, EYES ONLY ou DISPLAY ONLY."));
            }
            if (m.HasDissem("REL") && m.HasDissem("EYES"))
            {
                errors.Add(Text.T("REL TO may not be used with EYES ONLY.", "REL TO ne s'emploie pas avec EYES ONLY."));
            }
            if (m.HasDissem("RELIDO") && m.HasDissem("DISPLAYONLY"))
            {
                errors.Add(Text.T("RELIDO may not be used with DISPLAY ONLY.", "RELIDO ne s'emploie pas avec DISPLAY ONLY."));
            }
            if (m.HasDissem("OC") && m.HasDissem("OC-USGOV"))
            {
                errors.Add(Text.T("ORCON and ORCON-USGOV may not be used together: ORCON supersedes ORCON-USGOV.", "ORCON et ORCON-USGOV ne s'emploient pas ensemble : ORCON l'emporte sur ORCON-USGOV."));
            }
            if (m.HasDissem("RELIDO") && (m.HasDissem("OC") || m.HasDissem("OC-USGOV")))
            {
                errors.Add(Text.T("RELIDO may not be used with ORCON or ORCON-USGOV.", "RELIDO ne s'emploie pas avec ORCON ou ORCON-USGOV."));
            }
            if (m.HasNonIcDissem("XD") && m.HasNonIcDissem("ND"))
            {
                errors.Add(Text.T("EXDIS and NODIS may not be used together: NODIS supersedes EXDIS.", "EXDIS et NODIS ne s'emploient pas ensemble : NODIS l'emporte sur EXDIS."));
            }
            if ((m.HasNonIcDissem("XD") || m.HasNonIcDissem("ND")) && !nf)
            {
                errors.Add(Text.T("EXDIS and NODIS require NOFORN.", "EXDIS et NODIS exigent NOFORN."));
            }
            if (m.HasNonIcDissem("DS") && m.HasDissem("FOUO"))
            {
                errors.Add(Text.T("LIMDIS information may not be combined with FOUO information in a portion.", "Une information LIMDIS ne se combine pas avec une information FOUO dans une même portion."));
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
                    errors.Add(Text.T("\"REL TO USA\" alone is not authorized: name at least one other country or organization.", "« REL TO USA » seul n'est pas autorisé : citer au moins un autre pays ou une autre organisation."));
                }
            }
            if (m.HasDissem("EYES"))
            {
                foreach (string code in m.EyesOnly)
                {
                    if (code != "USA" && code != "AUS" && code != "CAN" && code != "GBR" && code != "NZL")
                    {
                        errors.Add(Text.T("EYES ONLY is only for Second Party partners (AUS, CAN, GBR, NZL): \"" + code + "\".", "EYES ONLY est réservé aux partenaires Second Party (AUS, CAN, GBR, NZL) : « " + code + " »."));
                    }
                }
            }
            if (m.HasDissem("DISPLAYONLY") && m.DisplayOnly.Contains("USA"))
            {
                errors.Add(Text.T("The DISPLAY ONLY list names foreign countries only, not USA.", "La liste DISPLAY ONLY ne cite que des pays étrangers, pas USA."));
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
                errors.Add(Text.T("JOINT information may not be NOFORN.", "Une information JOINT ne peut pas être NOFORN."));
            }
            if (!m.HasDissem("REL"))
            {
                errors.Add(Text.T("JOINT information requires REL TO, to at least all its co-owners.", "Une information JOINT exige REL TO, au moins vers tous ses copropriétaires."));
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
                    errors.Add(Text.T("JOINT information must be releasable to its co-owner " + owner + ".", "Une information JOINT doit être diffusable à son copropriétaire " + owner + "."));
                }
            }
        }
    }
}
