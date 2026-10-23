//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice.Tests
{
    //IGI 1300 du 9 août 2021: 7.1 (classification and marking), 7.6.1 (deadline), annexes 1, 36 and 37
    [TestClass]
    public static class FrenchTests
    {
        private static FrenchMarking P(string mark)
        {
            FrenchMarking m = FrenchParser.ParsePortion(mark);
            Assert.IsNotNull(m, "\"" + mark + "\" should be a paragraph mark");
            return m;
        }

        private static FrenchMarking Roll(params string[] marks)
        {
            List<FrenchMarking> portions = new List<FrenchMarking>();
            foreach (string mark in marks)
            {
                portions.Add(P(mark));
            }
            FrenchRollUpResult result = FrenchRollUp.Compute(portions);
            Assert.IsTrue(result.Errors.Count == 0, string.Join(" ", result.Errors.ToArray()));
            return result.Document;
        }

        public static void ParagraphMarks()
        {
            foreach (string mark in new string[] { "[NP]", "[DR]", "[S]", "[TS]", "[DR SF]", "[S SF]", "[TS SF]" })
            {
                Assert.AreEqual(mark, FrenchFormatter.Portion(P(mark)));
            }
            Assert.IsNull(FrenchParser.ParsePortion("[CD]"));        //Confidentiel Défense no longer exists since 1 July 2021
            Assert.IsNull(FrenchParser.ParsePortion("[see annex]"));
            Assert.IsNull(FrenchParser.ParsePortion("(S)"));         //IC portion marks are another system
        }

        public static void StampsInFullLetters()
        {
            Assert.AreEqual("TRÈS SECRET", FrenchFormatter.Stamp(FrenchLevel.TresSecret));
            Assert.AreEqual("SECRET", FrenchFormatter.Stamp(FrenchLevel.Secret));
            Assert.AreEqual("DIFFUSION RESTREINTE", FrenchFormatter.Stamp(FrenchLevel.DiffusionRestreinte));
            Assert.AreEqual("", FrenchFormatter.Stamp(FrenchLevel.NonProtege));
            //Diffusion Restreinte is stamped at the top of the page only (annexe 1 § 2), classification stamps at the top and bottom
            Assert.AreEqual("", FrenchFormatter.BottomStamp(P("[DR]")));
            Assert.AreEqual("SECRET", FrenchFormatter.BottomStamp(P("[S SF]")));
        }

        //annexe 36: [NP] title and paragraphs, Secret Spécial France paragraphs, a [DR] paragraph: the document is SECRET, SPÉCIAL FRANCE
        public static void Annexe36Example()
        {
            FrenchMarking document = Roll("[NP]", "[NP]", "[S SF]", "[NP]", "[S SF]", "[DR]", "[S SF]");
            Assert.AreEqual(FrenchLevel.Secret, document.Level);
            Assert.IsTrue(document.SpecialFrance, "one Spécial France part makes the document Spécial France");
        }

        public static void HighestLevel()
        {
            Assert.AreEqual(FrenchLevel.TresSecret, Roll("[S]", "[TS]", "[DR]").Level);
            Assert.AreEqual(FrenchLevel.DiffusionRestreinte, Roll("[NP]", "[DR]").Level);
            Assert.AreEqual(FrenchLevel.NonProtege, Roll("[NP]").Level);
            Assert.IsTrue(!Roll("[S]", "[DR]").SpecialFrance, "no Spécial France part");
        }

        private static FrenchIdentification Complete()
        {
            FrenchIdentification id = new FrenchIdentification();
            id.AutoriteEmettrice = "Ministère X, direction Y";
            id.Auteur = "Bureau Z";
            id.NumeroEnregistrement = "2026-001";
            id.DateEmission = new DateTime(2026, 10, 23);
            id.Echeance = new DateTime(2036, 10, 23);
            return id;
        }

        public static void IdentificationOfAClassifiedDocument()
        {
            FrenchMarking secret = P("[S]");
            Assert.AreEqual(0, Complete().Validate(secret).Count);
            FrenchIdentification id = Complete();
            id.Echeance = null;
            Assert.IsTrue(id.Validate(secret).Count == 1, "a classified document carries its deadline");
            id.Reevaluation = new DateTime(2046, 10, 23);
            Assert.AreEqual(0, id.Validate(secret).Count);
            id.Reevaluation = new DateTime(2046, 10, 24);
            Assert.IsTrue(id.Validate(secret).Count == 1, "re-evaluation within twenty years");
            id.Reevaluation = null;
            id.SansEcheance = true;
            Assert.AreEqual(0, id.Validate(secret).Count);
            id = Complete();
            id.Echeance = new DateTime(2076, 10, 23);
            Assert.IsTrue(id.Validate(secret).Count == 1, "deadline before the fifty years of communicability");
            id = Complete();
            id.Reevaluation = new DateTime(2030, 1, 1);
            Assert.IsTrue(id.Validate(secret).Count == 1, "one option only");
        }

        public static void TresSecretCopies()
        {
            FrenchMarking ts = P("[TS]");
            FrenchIdentification id = Complete();
            Assert.IsTrue(id.Validate(ts).Count == 2, "copy number and reference are required");
            id.Exemplaire = 2;
            id.NombreExemplaires = 5;
            id.Reference = "TS-2026-001";
            Assert.AreEqual(0, id.Validate(ts).Count);
            id.Exemplaire = 6;
            Assert.IsTrue(id.Validate(ts).Count == 1, "copy number within the total");
            string[] lines = Complete().Lines(ts);
            Assert.AreEqual("le 23 octobre 2026", lines[0]);
            Assert.AreEqual("Déclassifié le 23 octobre 2036", lines[4]);
        }

        public static void DiffusionRestreinte()
        {
            FrenchMarking dr = P("[DR SF]");
            FrenchIdentification id = new FrenchIdentification();
            Assert.IsTrue(id.Validate(dr).Count == 2, "issuing authority and registration number");
            id.AutoriteEmettrice = "Préfecture";
            id.NumeroEnregistrement = "42";
            Assert.AreEqual(0, id.Validate(dr).Count); //no deadline for a protection mention
        }

        public static void SpecialFranceNeedsProtection()
        {
            FrenchMarking np = new FrenchMarking();
            np.SpecialFrance = true;
            Assert.IsTrue(FrenchValidator.Validate(np).Count == 1, "Spécial France on non protected information");
            Assert.AreEqual(0, FrenchValidator.Validate(P("[DR SF]")).Count);
        }

        public static void Ceiling()
        {
            SystemSettings settings = new SystemSettings();
            Assert.IsTrue(settings.FrenchAvailable && settings.IcAvailable, "no Simple value: both systems");
            settings.Ceiling = Presets.Find(12); //France SECRET
            Assert.IsTrue(settings.FrenchAvailable && !settings.IcAvailable, "French computer: French markings only");
            Assert.IsNull(settings.CheckCeiling(P("[S SF]")));
            Assert.IsNull(settings.CheckCeiling(P("[DR]")));
            Assert.IsNotNull(settings.CheckCeiling(P("[TS]")), "TRÈS SECRET is above a SECRET computer");
            settings.Ceiling = Presets.Find(11); //France DIFFUSION RESTREINTE
            Assert.IsNotNull(settings.CheckCeiling(P("[S]")), "SECRET is above a Diffusion Restreinte computer");
            settings.Ceiling = Presets.Find(4); //United States SECRET
            Assert.IsTrue(!settings.FrenchAvailable && settings.IcAvailable, "US computer: IC markings only");
            Assert.IsNotNull(settings.CheckCeiling(P("[DR]")), "no French markings on a US computer");
            settings.Ceiling = Presets.Find(23); //Germany
            Assert.IsTrue(!settings.FrenchAvailable && !settings.IcAvailable, "German computer: neither");
        }
    }
}
