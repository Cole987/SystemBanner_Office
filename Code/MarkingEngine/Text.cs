//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Globalization;
using Microsoft.Win32;

namespace SystemBannerOffice.Markings
{
    //the add-in speaks English or French: every message is written in both, next to each other, T("...", "...");
    //the markings themselves (SECRET, TRÈS SECRET, GEHEIM, OFFICIAL-SENSITIVE...) are never translated
    public static class Text
    {
        public static bool French;

        public static string T(string english, string french)
        {
            return French ? french : english;
        }

        //OfficeLanguage = "fr" or "en" in HKLM\SOFTWARE\Policies\SystemBanner, then in HKCU\Software\SystemBanner;
        //without it, the language of the Office user interface (an LCID, 0 when unknown), then the language of Windows
        public static void Choose(int officeLanguageId)
        {
            string setting = Setting(Registry.LocalMachine, SystemSettings.PolicyKey) ?? Setting(Registry.CurrentUser, @"Software\SystemBanner");
            if (setting == "fr" || setting == "en")
            {
                French = setting == "fr";
            }
            else if (officeLanguageId > 0)
            {
                French = (officeLanguageId & 0x3FF) == 0x0C; //LANG_FRENCH, any country
            }
            else
            {
                French = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "fr";
            }
        }

        private static string Setting(RegistryKey root, string path)
        {
            try
            {
                using (RegistryKey key = root.OpenSubKey(path, false))
                {
                    string value = key == null ? null : key.GetValue("OfficeLanguage") as string;
                    return value == null ? null : value.Trim().ToLowerInvariant();
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
