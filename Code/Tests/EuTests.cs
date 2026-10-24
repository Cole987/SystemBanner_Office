//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice.Tests
{
    //Council Decision 2013/488/EU (CSR) and Commission Decision 2015/444 (CSD), see Docs\EU-marking-rules.md
    [TestClass]
    public static class EuTests
    {
        public static void Levels()
        {
            //CSR Art. 2(2): the bilingual string is the marking, French part first
            Assert.AreEqual("TRÈS SECRET UE/EU TOP SECRET", EuFormatter.Marking(EuLevel.TresSecret));
            Assert.AreEqual("SECRET UE/EU SECRET", EuFormatter.Marking(EuLevel.Secret));
            Assert.AreEqual("CONFIDENTIEL UE/EU CONFIDENTIAL", EuFormatter.Marking(EuLevel.Confidentiel));
            Assert.AreEqual("RESTREINT UE/EU RESTRICTED", EuFormatter.Marking(EuLevel.Restreint));
            //the English text of the CSD prints TRES without the accent
            Assert.AreEqual(EuLevel.TresSecret, EuParser.ParseMarking("TRES SECRET UE/EU TOP SECRET").Level);
            Assert.IsNull(EuParser.ParseMarking("SECRET UE"));
        }

        public static void ParagraphMarks()
        {
            //CSR Annex III pt 12: TS-UE/EU-TS, S-UE/EU-S, C-UE/EU-C, R-UE/EU-R
            foreach (EuLevel level in new EuLevel[] { EuLevel.Restreint, EuLevel.Confidentiel, EuLevel.Secret, EuLevel.TresSecret })
            {
                string mark = EuFormatter.Portion(level);
                Assert.AreEqual(level, EuParser.ParsePortion(mark).Level);
            }
            Assert.AreEqual(EuLevel.Secret, EuParser.ParsePortion("(S-UE/EU-S) text").Level);
            Assert.IsNull(EuParser.ParsePortion("[S-UE/EU-TS]")); //the two halves of the abbreviation must match
            Assert.IsNull(EuParser.ParsePortion("[S]"));
        }

        public static void HighestLevel()
        {
            List<EuMarking> parts = new List<EuMarking>();
            parts.Add(EuParser.ParsePortion("[R-UE/EU-R]"));
            parts.Add(EuParser.ParsePortion("[C-UE/EU-C]"));
            parts.Add(EuParser.ParsePortion("[R-UE/EU-R]"));
            Assert.AreEqual(EuLevel.Confidentiel, EuRollUp.Compute(parts).Level);
        }

        public static void ReferenceAndCopy()
        {
            EuMarking m = new EuMarking();
            m.Level = EuLevel.Confidentiel;
            Assert.IsTrue(EuValidator.Validate(m).Count == 1, "a reference number on every level (CSR Annex III pt 13(c))");
            m.Reference = "EEAS(2026) 42";
            m.Copy = 3;
            Assert.AreEqual(0, EuValidator.Validate(m).Count);
            Assert.AreEqual("EEAS(2026) 42", EuFormatter.Footer(m)); //copy numbers from SECRET UE/EU SECRET up (pt 13(e))
            m.Level = EuLevel.Secret;
            Assert.AreEqual("EEAS(2026) 42 - Ex./Copy 3", EuFormatter.Footer(m));
        }

        public static void Ceiling()
        {
            SystemSettings settings = new SystemSettings();
            Assert.IsTrue(EuSettings.Available(settings), "no Simple value: EU markings offered");
            settings.Ceiling = Presets.Find(62); //CONFIDENTIEL UE/EU CONFIDENTIAL
            Assert.AreEqual("European Union", settings.Ceiling.ClassificationSystem);
            EuMarking m = new EuMarking();
            m.Level = EuLevel.Confidentiel;
            Assert.IsNull(EuSettings.CheckCeiling(settings, m));
            m.Level = EuLevel.Secret;
            Assert.IsNotNull(EuSettings.CheckCeiling(settings, m), "SECRET UE above a CONFIDENTIEL UE computer");
            Assert.IsTrue(!settings.IcAvailable && !settings.FrenchAvailable && !GermanSettings.Available(settings) && !UkSettings.Available(settings), "an EU computer offers the EU markings only");
            settings.Ceiling = Presets.Find(12);
            Assert.IsTrue(!EuSettings.Available(settings), "French computer: no EU markings");
        }
    }
}
