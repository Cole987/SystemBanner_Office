//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;

namespace SystemBannerOffice.Markings
{
    //one entry of the Register: the marking title spelled out in a banner line, its authorized banner abbreviation
    //(null when there is none) and its authorized portion mark
    public class RegisterEntry
    {
        public string Key;          //stable key used in Marking lists and in the document properties, the portion mark for most entries
        public string Title;
        public string Abbreviation;
        public string Portion;
        public RegisterEntry(string key, string title, string abbreviation, string portion)
        {
            Key = key;
            Title = title;
            Abbreviation = abbreviation;
            Portion = portion;
        }
        //banner form: the abbreviation when there is one and abbreviations are chosen, else the title
        public string BannerText(bool abbreviate)
        {
            return abbreviate && Abbreviation != null ? Abbreviation : Title;
        }
    }

    //an SCI control system or compartment published in the Register, with the length rules of what goes below it
    public class SciEntry
    {
        public string Title;
        public string Abbreviation;
        public bool RequiresCompartment;
        public int MaxChildLength;       //0 when the Register sets no length for the compartments or sub-compartments below
        public bool ChildLengthExact;    //true when the length is exact (GAMMA sub-compartments: 4 characters)
        public SciEntry[] Children;      //published compartments, null when the compartments are free values
        public SciEntry(string title, string abbreviation, bool requiresCompartment, int maxChildLength, bool childLengthExact, params SciEntry[] children)
        {
            Title = title;
            Abbreviation = abbreviation;
            RequiresCompartment = requiresCompartment;
            MaxChildLength = maxChildLength;
            ChildLengthExact = childLengthExact;
            Children = children.Length > 0 ? children : null;
        }
    }

    //Table 4 of the IC Markings System Register and Manual (31 December 2016), in Register order:
    //"Multiple dissemination controls must be listed in the order they appear in the Register"
    public static class Register
    {
        public static readonly RegisterEntry[] Classifications = new RegisterEntry[]
        {
            new RegisterEntry("TS", "TOP SECRET", null, "TS"),
            new RegisterEntry("S", "SECRET", null, "S"),
            new RegisterEntry("C", "CONFIDENTIAL", null, "C"),
            new RegisterEntry("R", "RESTRICTED", null, "R"), //non-US and NATO only
            new RegisterEntry("U", "UNCLASSIFIED", null, "U")
        };

        //published SCI control systems; unpublished ones are free alphanumeric values, sorted with these
        public static readonly SciEntry[] Sci = new SciEntry[]
        {
            new SciEntry("HCS", "HCS", true, 0, false,
                new SciEntry("O", "O", false, 0, false),
                new SciEntry("P", "P", false, 6, false)),
            new SciEntry("RESERVE", "RSV", true, 3, false),
            new SciEntry("SI", "SI", false, 0, false,
                new SciEntry("ECRU", "EU", false, 0, false),
                new SciEntry("GAMMA", "G", false, 4, true),
                new SciEntry("NONBOOK", "NK", false, 0, false)),
            new SciEntry("TALENT KEYHOLE", "TK", false, 0, false,
                new SciEntry("BLUEFISH", "BLFH", false, 6, false),
                new SciEntry("IDITAROD", "IDIT", false, 0, false),
                new SciEntry("KANDIK", "KAND", false, 0, false))
        };

        public static readonly RegisterEntry[] Dissemination = new RegisterEntry[]
        {
            new RegisterEntry("RS", "RISK SENSITIVE", "RSEN", "RS"),
            new RegisterEntry("FOUO", "FOR OFFICIAL USE ONLY", "FOUO", "FOUO"),
            new RegisterEntry("OC", "ORIGINATOR CONTROLLED", "ORCON", "OC"),
            new RegisterEntry("OC-USGOV", "ORIGINATOR CONTROLLED-USGOV", "ORCON-USGOV", "OC-USGOV"),
            new RegisterEntry("IMC", "CONTROLLED IMAGERY", "IMCON", "IMC"),
            new RegisterEntry("NF", "NOT RELEASABLE TO FOREIGN NATIONALS", "NOFORN", "NF"),
            new RegisterEntry("PR", "CAUTION-PROPRIETARY INFORMATION INVOLVED", "PROPIN", "PR"),
            new RegisterEntry("REL", "AUTHORIZED FOR RELEASE TO", "REL TO", "REL"),
            new RegisterEntry("RELIDO", "RELEASABLE BY INFORMATION DISCLOSURE OFFICIAL", "RELIDO", "RELIDO"),
            new RegisterEntry("EYES", "EYES ONLY", null, "EYES"), //NSA SIGINT reporting only, under a waiver
            new RegisterEntry("DSEN", "DEA SENSITIVE", null, "DSEN"),
            new RegisterEntry("FISA", "FOREIGN INTELLIGENCE SURVEILLANCE ACT", "FISA", "FISA"),
            new RegisterEntry("DISPLAYONLY", "DISPLAY ONLY", null, "DISPLAY ONLY")
        };

        public static readonly RegisterEntry[] NonIcDissemination = new RegisterEntry[]
        {
            new RegisterEntry("DS", "LIMITED DISTRIBUTION", "LIMDIS", "DS"),
            new RegisterEntry("XD", "EXCLUSIVE DISTRIBUTION", "EXDIS", "XD"),
            new RegisterEntry("ND", "NO DISTRIBUTION", "NODIS", "ND"),
            new RegisterEntry("SBU", "SENSITIVE BUT UNCLASSIFIED", "SBU", "SBU"),
            new RegisterEntry("SBU-NF", "SENSITIVE BUT UNCLASSIFIED NOFORN", "SBU NOFORN", "SBU-NF"),
            new RegisterEntry("LES", "LAW ENFORCEMENT SENSITIVE", "LES", "LES"),
            new RegisterEntry("LES-NF", "LAW ENFORCEMENT SENSITIVE NOFORN", "LES NOFORN", "LES-NF"),
            new RegisterEntry("SSI", "SENSITIVE SECURITY INFORMATION", "SSI", "SSI")
        };

        //NAC activities of NATO/NAC portion marks, "N[NAC Activity]S": the Manual only shows NIS for NATO ISAF SECRET,
        //the full list is in Appendix B, which is not published with it; unknown activities are written as they are
        private static readonly string[,] NacActivities = new string[,] { { "ISAF", "I" } };

        public static string NacPortionCode(string activity)
        {
            for (int i = 0; i < NacActivities.GetLength(0); i++)
            {
                if (NacActivities[i, 0] == activity) return NacActivities[i, 1];
            }
            return activity;
        }

        public static string NacActivity(string portionCode)
        {
            for (int i = 0; i < NacActivities.GetLength(0); i++)
            {
                if (NacActivities[i, 1] == portionCode) return NacActivities[i, 0];
            }
            return portionCode;
        }

        //NATO special access program markings, Appendix B
        public static readonly string[] NatoSap = new string[] { "ATOMAL", "BALK", "BOHEMIA" };

        public static RegisterEntry FindDissemination(string key)
        {
            return Find(Dissemination, key);
        }

        public static RegisterEntry FindNonIcDissemination(string key)
        {
            return Find(NonIcDissemination, key);
        }

        public static RegisterEntry Find(RegisterEntry[] entries, string key)
        {
            foreach (RegisterEntry entry in entries)
            {
                if (entry.Key == key)
                {
                    return entry;
                }
            }
            return null;
        }

        //an entry from any of its written forms (key, title, abbreviation, portion mark), null when the text is no such entry
        public static RegisterEntry Recognize(RegisterEntry[] entries, string text)
        {
            foreach (RegisterEntry entry in entries)
            {
                if (entry.Key == text || entry.Title == text || entry.Abbreviation == text || entry.Portion == text)
                {
                    return entry;
                }
            }
            return null;
        }

        public static int OrderOf(RegisterEntry[] entries, string key)
        {
            for (int i = 0; i < entries.Length; i++)
            {
                if (entries[i].Key == key)
                {
                    return i;
                }
            }
            return int.MaxValue; //unknown values go last
        }

        //sorts keys in Register order, as banner lines and portion marks require
        public static void SortByRegister(List<string> keys, RegisterEntry[] entries)
        {
            keys.Sort(delegate (string a, string b) { return OrderOf(entries, a).CompareTo(OrderOf(entries, b)); });
        }

        //a published SCI control system from its title or abbreviation, null for unpublished ones
        public static SciEntry FindSci(string name)
        {
            return FindSci(Sci, name);
        }

        public static SciEntry FindSci(SciEntry[] entries, string name)
        {
            if (entries == null)
            {
                return null;
            }
            foreach (SciEntry entry in entries)
            {
                if (entry.Title == name || entry.Abbreviation == name)
                {
                    return entry;
                }
            }
            return null;
        }
    }
}
