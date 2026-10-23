//Copyright (c) 2023, Aaron Wawrzyniak. MIT License Applies.
//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;

namespace SystemBannerOffice.Markings
{
    //the stock classifications of SystemBanner, same codes, texts and colors as Program.presets in SystemBanner,
    //so a document is marked with the same values as the banner of the computer it is written on
    public class Preset
    {
        public int Code; //value of the Simple setting
        public string ClassificationSystem;
        public string Text;
        public int[] Colors; //background red, green, blue, then text red, green, blue, same order as the "0" to "5" registry values
        public int Rank; //order of the levels inside one system, 0 is the lowest
        public Preset(int code, string classificationSystem, string text, int rank, int[] colors)
        {
            Code = code;
            ClassificationSystem = classificationSystem;
            Text = text;
            Rank = rank;
            Colors = colors;
        }
        public bool IsUnitedStates
        {
            get { return Code >= 1 && Code <= 6; }
        }
    }

    public static class Presets
    {
        public const string UnitedStates = "United States";
        public static Preset[] All = new Preset[]
        {
            //United States: gsa form color info collected by Frank Caviggia @ https://github.com/fcaviggia/classification-banner
            new Preset(1, UnitedStates, "UNCLASSIFIED", 0, new int[6] { 0, 122, 51, 255, 255, 255 }),
            new Preset(2, UnitedStates, "CUI", 1, new int[6] { 48, 31, 85, 255, 255, 255 }),
            new Preset(3, UnitedStates, "CONFIDENTIAL", 2, new int[6] { 0, 51, 160, 255, 255, 255 }),
            new Preset(4, UnitedStates, "SECRET", 3, new int[6] { 200, 16, 46, 255, 255, 255 }),
            new Preset(5, UnitedStates, "TOP SECRET", 4, new int[6] { 255, 103, 31, 255, 255, 255 }),
            new Preset(6, UnitedStates, "TOP SECRET//SCI", 5, new int[6] { 247, 234, 72, 0, 0, 0 }),
            new Preset(11, "France", "DIFFUSION RESTREINTE", 0, new int[6] { 0, 0, 0, 255, 255, 255 }),
            new Preset(12, "France", "SECRET", 1, new int[6] { 200, 16, 46, 255, 255, 255 }),
            new Preset(13, "France", "TRÈS SECRET", 2, new int[6] { 200, 16, 46, 255, 255, 255 }),
            new Preset(21, "Germany", "VS-NUR FÜR DEN DIENSTGEBRAUCH", 0, new int[6] { 0, 0, 0, 255, 255, 255 }),
            new Preset(22, "Germany", "VS-VERTRAULICH", 1, new int[6] { 0, 51, 160, 255, 255, 255 }),
            new Preset(23, "Germany", "GEHEIM", 2, new int[6] { 200, 16, 46, 255, 255, 255 }),
            new Preset(24, "Germany", "STRENG GEHEIM", 3, new int[6] { 200, 16, 46, 247, 234, 72 }),
            new Preset(31, "United Kingdom", "OFFICIAL", 0, new int[6] { 0, 0, 0, 255, 255, 255 }),
            new Preset(32, "United Kingdom", "OFFICIAL-SENSITIVE", 1, new int[6] { 0, 0, 0, 255, 255, 255 }),
            new Preset(33, "United Kingdom", "SECRET", 2, new int[6] { 0, 0, 0, 255, 255, 255 }),
            new Preset(34, "United Kingdom", "TOP SECRET", 3, new int[6] { 0, 0, 0, 255, 255, 255 }),
            new Preset(41, "NATO", "NATO UNCLASSIFIED", 0, new int[6] { 0, 122, 51, 255, 255, 255 }),
            new Preset(42, "NATO", "NATO RESTRICTED", 1, new int[6] { 255, 103, 31, 255, 255, 255 }),
            new Preset(43, "NATO", "NATO CONFIDENTIAL", 2, new int[6] { 0, 51, 160, 255, 255, 255 }),
            new Preset(44, "NATO", "NATO SECRET", 3, new int[6] { 200, 16, 46, 255, 255, 255 }),
            new Preset(45, "NATO", "COSMIC TOP SECRET", 4, new int[6] { 247, 234, 72, 0, 0, 0 }),
            new Preset(51, "United Nations", "UN CONFIDENTIAL", 0, new int[6] { 0, 0, 0, 255, 255, 255 }),
            new Preset(52, "United Nations", "UN STRICTLY CONFIDENTIAL", 1, new int[6] { 0, 0, 0, 255, 255, 255 })
        };

        public static Preset Find(int code) //null when no stock classification has this code
        {
            foreach (Preset preset in All)
            {
                if (preset.Code == code)
                {
                    return preset;
                }
            }
            return null;
        }

        public static Preset Find(string classificationSystem, string text) //null when the system has no such level
        {
            foreach (Preset preset in All)
            {
                if (preset.ClassificationSystem == classificationSystem && preset.Text == text)
                {
                    return preset;
                }
            }
            return null;
        }

        //the levels of one system, lowest first, up to the ceiling when there is one
        public static Preset[] LevelsOf(string classificationSystem, Preset ceiling)
        {
            System.Collections.Generic.List<Preset> levels = new System.Collections.Generic.List<Preset>();
            foreach (Preset preset in All)
            {
                if (preset.ClassificationSystem == classificationSystem && (ceiling == null || preset.Rank <= ceiling.Rank))
                {
                    levels.Add(preset);
                }
            }
            return levels.ToArray();
        }

        //colors of an IC banner line, from its US classification and SCI, the same choice the banner makes for codes 1 to 6
        public static int[] ColorsFor(Marking marking)
        {
            Preset preset;
            if (marking.Level == Level.TopSecret && marking.Sci.Count > 0)
            {
                preset = Find(6);
            }
            else if (marking.Level == Level.TopSecret)
            {
                preset = Find(5);
            }
            else if (marking.Level == Level.Secret)
            {
                preset = Find(4);
            }
            else if (marking.Level == Level.Confidential || marking.Level == Level.Restricted)
            {
                preset = Find(3);
            }
            else if (marking.HasControls)
            {
                preset = Find(2); //unclassified with controls, purple like CUI
            }
            else
            {
                preset = Find(1);
            }
            return preset.Colors;
        }
    }
}
