//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Globalization;
using Microsoft.Win32;

namespace SystemBannerOffice.Markings
{
    public enum Language { English, French, German }

    //the add-in speaks English, French or German: every message is written in the three, next to each other, T("...", "...", "...");
    //the markings themselves (SECRET, TRÈS SECRET, GEHEIM, OFFICIAL-SENSITIVE...) are never translated
    public static class Text
    {
        public static Language Current;

        public static bool French
        {
            get { return Current == Language.French; }
            set { Current = value ? Language.French : Language.English; }
        }

        public static bool German
        {
            get { return Current == Language.German; }
            set { Current = value ? Language.German : Language.English; }
        }

        public static string T(string english, string french, string german)
        {
            return Current == Language.French ? french : (Current == Language.German ? german : english);
        }

        //a message not translated into German yet shows in English
        public static string T(string english, string french)
        {
            return Current == Language.French ? french : english;
        }

        //OfficeLanguage = "fr", "de" or "en" in HKLM\SOFTWARE\Policies\SystemBanner, then in HKCU\Software\SystemBanner;
        //without it, the language of the Office user interface (an LCID, 0 when unknown), then the language of Windows
        public static void Choose(int officeLanguageId)
        {
            string setting = Setting(Registry.LocalMachine, SystemSettings.PolicyKey) ?? Setting(Registry.CurrentUser, @"Software\SystemBanner");
            if (setting == "fr" || setting == "de" || setting == "en")
            {
                Current = setting == "fr" ? Language.French : (setting == "de" ? Language.German : Language.English);
            }
            else if (officeLanguageId > 0)
            {
                int primary = officeLanguageId & 0x3FF; //LANG_FRENCH 0x0C, LANG_GERMAN 0x07, any country
                Current = primary == 0x0C ? Language.French : (primary == 0x07 ? Language.German : Language.English);
            }
            else
            {
                string windows = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
                Current = windows == "fr" ? Language.French : (windows == "de" ? Language.German : Language.English);
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
