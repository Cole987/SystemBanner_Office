//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;

namespace SystemBannerOffice.Markings
{
    //sort orders set by the Register and Manual, A.6 Formatting
    public static class Sorting
    {
        //SCI and SAP values: ascending, all numbered values first, then alphabetic values
        public static int CompareAlphanumeric(string a, string b)
        {
            bool aNumber = a.Length > 0 && char.IsDigit(a[0]);
            bool bNumber = b.Length > 0 && char.IsDigit(b[0]);
            if (aNumber != bNumber)
            {
                return aNumber ? -1 : 1;
            }
            if (aNumber && IsAllDigits(a) && IsAllDigits(b))
            {
                //compare as numbers, without parsing so long values cannot overflow
                string ta = a.TrimStart('0');
                string tb = b.TrimStart('0');
                if (ta.Length != tb.Length)
                {
                    return ta.Length < tb.Length ? -1 : 1;
                }
                int digits = string.CompareOrdinal(ta, tb);
                if (digits != 0)
                {
                    return digits;
                }
            }
            return string.CompareOrdinal(a, b);
        }

        public static bool IsAllDigits(string value)
        {
            if (value.Length == 0)
            {
                return false;
            }
            foreach (char c in value)
            {
                if (!char.IsDigit(c))
                {
                    return false;
                }
            }
            return true;
        }

        public static void SortAlphanumeric(List<string> values)
        {
            values.Sort(CompareAlphanumeric);
        }

        //country and organization codes: trigraphs in ascending order, then tetragraphs in ascending order,
        //with USA first when usaFirst is set (REL TO, EYES ONLY)
        public static List<string> SortCountries(IEnumerable<string> codes, bool usaFirst)
        {
            List<string> result = new List<string>();
            foreach (string code in codes)
            {
                if (!result.Contains(code))
                {
                    result.Add(code);
                }
            }
            result.Sort(delegate (string a, string b)
            {
                if (usaFirst && a != b)
                {
                    if (a == "USA") return -1;
                    if (b == "USA") return 1;
                }
                bool aTrigraph = a.Length == 3;
                bool bTrigraph = b.Length == 3;
                if (aTrigraph != bTrigraph)
                {
                    return aTrigraph ? -1 : 1;
                }
                return string.CompareOrdinal(a, b);
            });
            return result;
        }
    }
}
