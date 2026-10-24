//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice.Tests
{
    //VS-Anweisung of 13 March 2023: § 16, § 20, Anlage IV
    [TestClass]
    public static class GermanTests
    {
        private static GermanMarking P(string mark)
        {
            GermanMarking m = GermanParser.ParsePortion(mark);
            Assert.IsNotNull(m, "\"" + mark + "\" should be a part mark");
            return m;
        }

        public static void StampsWrittenInFull()
        {
            Assert.AreEqual("STRENG GEHEIM", GermanFormatter.Stamp(GermanLevel.StrengGeheim));
            Assert.AreEqual("VS-NUR FÜR DEN DIENSTGEBRAUCH", GermanFormatter.Stamp(GermanLevel.VsNfd));
            Assert.IsTrue(P("[Geh.]").AmtlichGeheimGehalten && P("[VS-Vertr.]").AmtlichGeheimGehalten, "amtlich geheim gehalten from VS-VERTRAULICH up");
            Assert.IsTrue(!P("[VS-NfD]").AmtlichGeheimGehalten, "VS-NfD: the level alone");
            Assert.IsTrue(GermanFormatter.AtBottom(GermanLevel.Geheim) && !GermanFormatter.AtBottom(GermanLevel.VsVertraulich), "top and bottom for GEHEIM, top only for VS-VERTRAULICH");
        }

        public static void PartMarks()
        {
            foreach (string mark in new string[] { "[Str. Geh.]", "[Geh.]", "[VS-Vertr.]", "[VS-NfD]", "[offen]" })
            {
                Assert.AreEqual(mark, GermanFormatter.Portion(P(mark).Level));
            }
            Assert.IsNull(GermanParser.ParsePortion("[S]"));
        }

        public static void HighestLevel()
        {
            List<GermanMarking> parts = new List<GermanMarking>();
            parts.Add(P("[offen]"));
            parts.Add(P("[VS-NfD]"));
            parts.Add(P("[Geh.]"));
            Assert.AreEqual(GermanLevel.Geheim, GermanRollUp.Compute(parts).Level);
        }

        public static void FirstPage()
        {
            GermanIdentification id = new GermanIdentification();
            id.Datum = new DateTime(2026, 10, 24);
            GermanMarking geheim = P("[Geh.]");
            Assert.AreEqual(4, id.Validate(geheim).Count); //issuer, end of classification, reference, copy
            id.Herausgeber = "Bundesministerium X";
            id.EinstufungEndet = 2056;
            id.Kennzeichen = "Tgb.-Nr. 42/26";
            id.Ausfertigung = 2;
            id.Empfaenger = "Referat Y";
            Assert.AreEqual(0, id.Validate(geheim).Count);
            id.EinstufungEndet = 2057;
            Assert.IsTrue(id.Validate(geheim).Count == 1, "30 years at most");
            string[] lines = new GermanIdentification().Lines(geheim);
            Assert.AreEqual("{NUMPAGES} Seiten", lines[lines.Length - 1]);
            id.EinstufungEndet = 2036;
            lines = id.Lines(geheim);
            Assert.AreEqual("Tgb.-Nr. 42/26 Geh.", lines[2]);
            Assert.AreEqual("Die Einstufung endet mit Ablauf des Jahres 2036.", lines[4]);
            GermanIdentification nfd = new GermanIdentification();
            nfd.Herausgeber = "Behörde";
            Assert.AreEqual(0, nfd.Validate(P("[VS-NfD]")).Count); //30 years by default, no copy number
        }

        public static void Ceiling()
        {
            SystemSettings settings = new SystemSettings();
            Assert.IsTrue(GermanSettings.Available(settings), "no Simple value: German markings offered");
            settings.Ceiling = Presets.Find(23); //GEHEIM
            Assert.IsNull(GermanSettings.CheckCeiling(settings, P("[Geh.]")));
            Assert.IsNotNull(GermanSettings.CheckCeiling(settings, P("[Str. Geh.]")), "STRENG GEHEIM above a GEHEIM computer");
            settings.Ceiling = Presets.Find(12);
            Assert.IsTrue(!GermanSettings.Available(settings), "French computer: no German markings");
        }
    }

    [TestClass]
    public static class LanguageTests
    {
        public static void EnglishAndFrench()
        {
            bool saved = Text.French;
            try
            {
                Text.French = false;
                string english = MarkingValidator.Validate(MarkingParser.Parse("SECRET//FOUO").Marking)[0];
                Text.French = true;
                string french = MarkingValidator.Validate(MarkingParser.Parse("SECRET//FOUO").Marking)[0];
                Assert.IsTrue(english != french && english.Contains("FOUO") && french.Contains("FOUO"), "the same rule in two languages, the marking untranslated: " + english + " / " + french);
                //without an OfficeLanguage setting, the language of the Office user interface decides
                Text.Choose(0x040C); //French (France)
                Assert.IsTrue(Text.French, "a French Office speaks French (unless OfficeLanguage is set on this computer)");
                Text.Choose(0x0C0C); //French (Canada)
                Assert.IsTrue(Text.French, "any French Office speaks French");
                Text.Choose(0x0409); //English (United States)
                Assert.IsTrue(!Text.French, "an English Office speaks English (unless OfficeLanguage is set on this computer)");
            }
            finally
            {
                Text.French = saved;
            }
        }

        public static void German()
        {
            Language saved = Text.Current;
            try
            {
                Text.Current = Language.German;
                string german = MarkingValidator.Validate(MarkingParser.Parse("SECRET//FOUO").Marking)[0];
                Text.Current = Language.English;
                string english = MarkingValidator.Validate(MarkingParser.Parse("SECRET//FOUO").Marking)[0];
                Assert.IsTrue(english != german && german.Contains("FOUO"), "the same rule in German, the marking untranslated: " + german);
                Text.Current = Language.German;
                Assert.AreEqual("Ja", Text.T("Yes", "Oui", "Ja"));
                Assert.IsTrue(Text.T("Yes", "Oui") == "Yes", "a message without German text shows in English");
                Text.Choose(0x0407); //German (Germany)
                Assert.IsTrue(Text.German, "a German Office speaks German (unless OfficeLanguage is set on this computer)");
                Text.Choose(0x0C07); //German (Austria)
                Assert.IsTrue(Text.German, "any German Office speaks German");
            }
            finally
            {
                Text.Current = saved;
            }
        }
    }

    //Government Security Classifications Policy
    [TestClass]
    public static class BritishTests
    {
        public static void Classifications()
        {
            UkMarking m = new UkMarking();
            Assert.AreEqual("OFFICIAL", m.ToString());
            m.Sensitive = true;
            Assert.AreEqual("OFFICIAL-SENSITIVE", m.ToString());
            m.Descriptors.Add("PERSONAL DATA");
            m.HandlingInstruction = UkRegister.RecipientsOnly;
            Assert.AreEqual("OFFICIAL-SENSITIVE RECIPIENTS ONLY PERSONAL DATA", m.ToString());
            Assert.AreEqual(0, UkValidator.Validate(m).Count);
        }

        public static void NationalCaveats()
        {
            UkMarking m = new UkMarking();
            m.Level = UkLevel.Secret;
            m.Caveats.Add("UK/US EYES ONLY");
            m.Caveats.Add("FIVE EYES ONLY");
            Assert.AreEqual("SECRET FIVE EYES ONLY UK/US EYES ONLY", m.ToString()); //alphabetical order
            Assert.AreEqual(0, UkValidator.Validate(m).Count);
            m.Level = UkLevel.Official;
            Assert.IsTrue(UkValidator.Validate(m).Count == 1, "national caveats on SECRET and TOP SECRET only");
        }

        public static void AllowedTiers()
        {
            UkMarking m = new UkMarking();
            m.HandlingInstruction = UkRegister.RecipientsOnly;
            Assert.IsTrue(UkValidator.Validate(m).Count == 1, "RECIPIENTS ONLY needs OFFICIAL-SENSITIVE or above");
            m.HandlingInstruction = UkRegister.ForPublicRelease;
            m.Sensitive = true;
            Assert.IsTrue(UkValidator.Validate(m).Count == 1, "FOR PUBLIC RELEASE is OFFICIAL only, never with -SENSITIVE");
            m = new UkMarking();
            m.Level = UkLevel.TopSecret;
            m.HandlingInstruction = UkRegister.HmgUseOnly;
            m.Descriptors.Add("LEGAL");
            Assert.IsTrue(UkValidator.Validate(m).Count == 2, "HMG USE ONLY up to SECRET, LEGAL up to OFFICIAL-SENSITIVE");
            m = new UkMarking();
            m.Level = UkLevel.Secret;
            m.Sensitive = true;
            Assert.IsTrue(UkValidator.Validate(m).Count == 1, "-SENSITIVE marks OFFICIAL only");
        }

        public static void PrefixAndCodeword()
        {
            UkMarking m = new UkMarking();
            m.Level = UkLevel.Secret;
            m.UkPrefix = true;
            m.RelEu = true;
            m.Codeword = "TALISMAN";
            m.HandlingInstruction = UkRegister.OrganisationUseOnly;
            m.Organisation = "Cabinet Office";
            Assert.AreEqual("UK SECRET REL EU CABINET OFFICE USE ONLY TALISMAN", m.ToString());
            Assert.AreEqual(0, UkValidator.Validate(m).Count);
            m.Codeword = "Two words";
            Assert.IsTrue(UkValidator.Validate(m).Count == 1, "a codeword is one word in capitals");
        }

        public static void Ceiling()
        {
            SystemSettings settings = new SystemSettings();
            settings.Ceiling = Presets.Find(32); //OFFICIAL-SENSITIVE
            UkMarking m = new UkMarking();
            m.Sensitive = true;
            Assert.IsNull(UkSettings.CheckCeiling(settings, m));
            m.Sensitive = false;
            m.Level = UkLevel.Secret;
            Assert.IsNotNull(UkSettings.CheckCeiling(settings, m), "SECRET above an OFFICIAL-SENSITIVE computer");
            settings.Ceiling = Presets.Find(4);
            Assert.IsTrue(!UkSettings.Available(settings), "US computer: no British markings");
        }
    }
}
