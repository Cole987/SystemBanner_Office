//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using Microsoft.Win32;

namespace SystemBannerOffice.Markings
{
    //the classification of the computer, read from the same Group Policy values as SystemBanner;
    //it is the ceiling: a document written on this computer cannot be marked above it
    public class SystemSettings
    {
        public const string PolicyKey = @"SOFTWARE\Policies\SystemBanner";
        public Preset Ceiling; //null when Simple is not set, 0 or unknown: no ceiling
        public string CustomText; //the Text value, shown for information only, custom texts have no level to compare
        //Office add-in values, in the same key as the banner's
        public bool RequireClassification; //OfficeRequireClassification = 1: no saving or printing of a document with no classification
        public bool SpellOut;              //OfficeSpellOut = 1: banner lines spell out control markings (TALENT KEYHOLE) instead of abbreviating them (TK)
        public bool AnalyticProducts;      //OfficeAnalyticProducts = 1: documents are disseminated analytic products, every classified portion needs FD&R
        public bool Toolbar;               //OfficeToolbar = 1: the Office 2000-2003 toolbar in Office 2007 and later too (Add-ins tab), besides the ribbon

        public bool IsUnitedStates
        {
            get { return Ceiling == null || Ceiling.IsUnitedStates; }
        }

        public static SystemSettings Read()
        {
            SystemSettings settings = new SystemSettings();
            try
            {
                //HKLM\SOFTWARE\Policies is shared by the 32 and 64 bit registry views, so a 64 bit Office 2010 reads the same values; .NET 2.0 has no registry views anyway
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(PolicyKey, false))
                {
                    if (key != null)
                    {
                        object simpleValue = key.GetValue("Simple");
                        if (simpleValue is int)
                        {
                            settings.Ceiling = Presets.Find((int)simpleValue); //null for 0 and unknown codes
                        }
                        settings.CustomText = key.GetValue("Text") as string;
                        settings.RequireClassification = IsOne(key.GetValue("OfficeRequireClassification"));
                        settings.SpellOut = IsOne(key.GetValue("OfficeSpellOut"));
                        settings.AnalyticProducts = IsOne(key.GetValue("OfficeAnalyticProducts"));
                        settings.Toolbar = IsOne(key.GetValue("OfficeToolbar"));
                    }
                }
            }
            catch (Exception)
            {
                //no access to the key: same as no setting, the add-in still marks documents, with no ceiling
            }
            try
            {
                //the toolbar is a display preference, a user may turn it on; the ceiling and the save check stay with the administrator
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\SystemBanner", false))
                {
                    if (key != null && IsOne(key.GetValue("OfficeToolbar")))
                    {
                        settings.Toolbar = true;
                    }
                }
            }
            catch (Exception)
            {
            }
            return settings;
        }

        private static bool IsOne(object value)
        {
            return value is int && (int)value == 1;
        }

        //rank of a US marking on the scale of presets 1 to 6, to compare it with the ceiling
        public static int UnitedStatesRank(Marking marking)
        {
            if (marking.Level == Level.TopSecret && marking.Sci.Count > 0) return 5;
            if (marking.Level == Level.TopSecret) return 4;
            if (marking.Level == Level.Secret) return 3;
            if (marking.Level == Level.Confidential || marking.Level == Level.Restricted) return 2;
            return 0; //unclassified, with or without controls (FOUO, LES...), fits on any computer, CUI and UNCLASSIFIED ones included
        }

        //null when the marking fits on this computer, else the reason it does not
        public string CheckCeiling(Marking marking)
        {
            if (Ceiling == null)
            {
                return null;
            }
            if (!Ceiling.IsUnitedStates)
            {
                return "This computer is set to the " + Ceiling.ClassificationSystem + " classification system (" + Ceiling.Text + "): IC markings are not available, use the " + Ceiling.ClassificationSystem + " levels.";
            }
            if (UnitedStatesRank(marking) > Ceiling.Rank)
            {
                return "This computer is set to " + Ceiling.Text + ": a document marked " + MarkingFormatter.Banner(marking) + " cannot be written on it.";
            }
            return null;
        }
    }
}
