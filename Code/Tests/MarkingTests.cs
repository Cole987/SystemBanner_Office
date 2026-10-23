//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice.Tests
{
    //examples of the IC Markings System Register and Manual (31 December 2016): section J and the notional pages of the marking templates
    internal static class Helpers
    {
        public static Marking Parse(string text)
        {
            ParseResult result = MarkingParser.Parse(text);
            Assert.IsTrue(result.Success, "\"" + text + "\" does not parse: " + string.Join(" ", result.Errors.ToArray()));
            return result.Marking;
        }

        //banner and portion written back exactly as given
        public static void RoundTrip(string banner, string portion)
        {
            Assert.AreEqual(banner, MarkingFormatter.Banner(Parse(banner)));
            Assert.AreEqual(portion, "(" + MarkingFormatter.Portion(Parse(portion), Parse(banner)) + ")");
        }

        public static string RollWith(string previousBanner, params string[] portions)
        {
            List<Marking> markings = new List<Marking>();
            foreach (string portion in portions)
            {
                markings.Add(Parse(portion));
            }
            RollUpOptions options = new RollUpOptions();
            if (previousBanner != null)
            {
                options.PreviousBanner = Parse(previousBanner);
            }
            RollUpResult result = RollUp.Compute(markings, options);
            Assert.IsTrue(result.Errors.Count == 0, "roll-up errors: " + string.Join(" ", result.Errors.ToArray()));
            return MarkingFormatter.Banner(result.Banner);
        }

        public static string Roll(params string[] portions)
        {
            return RollWith(null, portions);
        }

        public static void Invalid(string text)
        {
            Assert.IsTrue(!MarkingParser.Parse(text).Success, "\"" + text + "\" should not parse");
        }
    }

    [TestClass]
    public static class SectionJExamples
    {
        public static void Basic()
        {
            Helpers.RoundTrip("SECRET//REL TO USA, FVEY/RELIDO", "(S//REL/RELIDO)");
            Assert.AreEqual("C//REL TO USA, FVEY/RELIDO", MarkingFormatter.Portion(Helpers.Parse("(C//REL TO USA, FVEY/RELIDO)")));
        }

        public static void MultipleSciControlSystems()
        {
            Helpers.RoundTrip("TOP SECRET//SI-G/TK//RSEN/ORCON/NOFORN", "(TS//SI-G/TK//RS/OC/NF)");
            Assert.AreEqual("TOP SECRET//SI-GAMMA/TALENT KEYHOLE//RISK SENSITIVE/ORIGINATOR CONTROLLED/NOT RELEASABLE TO FOREIGN NATIONALS",
                MarkingFormatter.Banner(Helpers.Parse("TOP SECRET//SI-GAMMA/TALENT KEYHOLE//RISK SENSITIVE/ORIGINATOR CONTROLLED/NOFORN"), false));
        }

        public static void MultipleSciCompartments()
        {
            Helpers.RoundTrip("TOP SECRET//SI-ABC-DEF//ORCON/NOFORN", "(TS//SI-ABC-DEF//OC/NF)");
        }

        public static void MultipleSciSubCompartments()
        {
            Helpers.RoundTrip("TOP SECRET//SI-G ABCD EFGH-XYZ//ORCON/NOFORN", "(TS//SI-G ABCD EFGH-XYZ//OC/NF)");
        }

        public static void UnpublishedSci()
        {
            Helpers.RoundTrip("TOP SECRET//ANB/SI/TK/XNB//NOFORN", "(TS//ANB/SI/TK/XNB//NF)");
        }

        public static void SciAndSap()
        {
            Helpers.RoundTrip("TOP SECRET//HCS-P/LMN//SAR-BP//NOFORN", "(TS//HCS-P/LMN//SAR-BP//NF)");
        }

        public static void MultipleSap()
        {
            Helpers.RoundTrip("TOP SECRET//SAR-BP-123/CA-XYZ YYY//NOFORN", "(TS//SAR-BP-123/CA-XYZ YYY//NF)");
            Marking spelledOut = Helpers.Parse("TOP SECRET//SPECIAL ACCESS REQUIRED-BUTTER POPCORN-123/CANDY APPLE-XYZ YYY//NOFORN");
            Assert.AreEqual("TOP SECRET//SPECIAL ACCESS REQUIRED-BUTTER POPCORN-123/CANDY APPLE-XYZ YYY//NOT RELEASABLE TO FOREIGN NATIONALS", MarkingFormatter.Banner(spelledOut, false));
        }

        public static void AtomicEnergyAct()
        {
            Helpers.RoundTrip("TOP SECRET//RD-CNWDI//NOFORN", "(TS//RD-CNWDI//NF)");
            Helpers.RoundTrip("SECRET//FRD-SIGMA 14 18//REL TO USA, ACGU", "(S//FRD-SG 14 18//REL)");
        }

        public static void NonUs()
        {
            Helpers.RoundTrip("//COSMIC TOP SECRET//BOHEMIA", "(//CTS//BOHEMIA)");
            Helpers.RoundTrip("//DEU SECRET//REL TO USA, DEU", "(//DEU S//REL)");
            Assert.AreEqual("//DEU S//REL TO USA, DEU", MarkingFormatter.Portion(Helpers.Parse("(//DEU S//REL TO USA, DEU)")));
            Helpers.RoundTrip("//NATO SECRET//ATOMAL//ORCON", "(//NS//ATOMAL//OC)");
        }

        public static void Joint()
        {
            Helpers.RoundTrip("//JOINT SECRET CAN GBR USA//REL TO USA, CAN, GBR", "(//JOINT S//REL)");
        }

        public static void ForeignGovernmentInformation()
        {
            Helpers.RoundTrip("TOP SECRET//FGI DEU GBR//REL TO USA, DEU, GBR", "(TS//FGI DEU GBR//REL)");
            Assert.AreEqual("(TS//FGI DEU GBR//REL TO USA, DEU, GBR)", MarkingFormatter.PortionWithParentheses(Helpers.Parse("(TS//FGI DEU GBR//REL TO USA, DEU, GBR)")));
            Helpers.RoundTrip("SECRET//TK//FGI//NOFORN", "(S//TK//FGI//NF)");
            Assert.AreEqual("//FGI S//NF", MarkingFormatter.Portion(Helpers.Parse("(//FGI S//NF)")));
        }

        public static void Dissemination()
        {
            Helpers.RoundTrip("SECRET//REL TO USA, DEU/RELIDO", "(S//REL/RELIDO)");
            Helpers.RoundTrip("SECRET", "(S)");
            Helpers.RoundTrip("SECRET//NOFORN/PROPIN", "(S//NF/PR)");
            Helpers.RoundTrip("UNCLASSIFIED//REL TO USA, FVEY", "(U//REL)");
            Helpers.RoundTrip("UNCLASSIFIED//FOUO/NOFORN", "(U//FOUO/NF)");
            Helpers.RoundTrip("UNCLASSIFIED//SSI", "(U//SSI)");
        }
    }

    [TestClass]
    public static class Syntax
    {
        public static void SortsSciNumbersFirst()
        {
            Marking m = Helpers.Parse("TOP SECRET//SI-MMM AACD-G DEFG ABCD/123//ORCON/NOFORN");
            Assert.AreEqual("TOP SECRET//123/SI-G ABCD DEFG-MMM AACD//ORCON/NOFORN", MarkingFormatter.Banner(m));
        }

        public static void SortsSapPrograms()
        {
            Assert.AreEqual("SECRET//SAR-ABC-DEF 123/SDA-121//NOFORN", MarkingFormatter.Banner(Helpers.Parse("SECRET//SAR-SDA-121/ABC-DEF 123//NOFORN")));
        }

        public static void SortsCountries()
        {
            Assert.AreEqual("SECRET//REL TO USA, GBR, JPN, ISAF, NATO", MarkingFormatter.Banner(Helpers.Parse("SECRET//REL TO USA, NATO, JPN, ISAF, GBR")));
            Assert.AreEqual("SECRET//FGI GBR JPN NATO//REL TO USA, GBR, JPN, NATO", MarkingFormatter.Banner(Helpers.Parse("SECRET//FGI NATO JPN GBR//REL TO USA, GBR, JPN, NATO")));
        }

        public static void SensitiveButUnclassifiedNoforn()
        {
            Helpers.RoundTrip("UNCLASSIFIED//SBU NOFORN", "(U//SBU-NF)");
            Helpers.RoundTrip("UNCLASSIFIED//LES NOFORN", "(U//LES-NF)");
        }

        public static void AeaWithAtomal()
        {
            Assert.AreEqual("SECRET//RD/ATOMAL//FGI NATO//NOFORN", MarkingFormatter.Banner(Helpers.Parse("SECRET//RD/ATOMAL//FGI NATO//NOFORN")));
        }

        public static void SigmaNumbers()
        {
            Helpers.RoundTrip("SECRET//RD-SIGMA 18 20", "(S//RD-SG 18 20)");
            Assert.AreEqual("SECRET//RESTRICTED DATA-SIGMA 20//NOT RELEASABLE TO FOREIGN NATIONALS", MarkingFormatter.Banner(Helpers.Parse("(S//RD-SG 20//NF)"), false));
        }

        public static void EyesOnly()
        {
            Helpers.RoundTrip("SECRET//USA/CAN/GBR EYES ONLY", "(S//EYES)");
        }

        public static void DisplayOnly()
        {
            Helpers.RoundTrip("CONFIDENTIAL//DISPLAY ONLY AFG, IRQ", "(C//DISPLAY ONLY AFG, IRQ)");
        }

        public static void NatoNacActivity()
        {
            Marking m = Helpers.Parse("(//NIS//REL TO USA, ISAF, NATO)");
            Assert.AreEqual(Owner.Nato, m.Owner);
            Assert.AreEqual("ISAF", m.NatoActivity);
            Assert.AreEqual("//NATO ISAF SECRET//REL TO USA, ISAF, NATO", MarkingFormatter.Banner(m));
        }

        public static void RejectsBadSyntax()
        {
            Helpers.Invalid("secret");
            Helpers.Invalid("SECRET//NOFORN//SI");          //categories out of order
            Helpers.Invalid("SECRET//PROPIN/NOFORN");       //controls out of Register order
            Helpers.Invalid("SECRET////NOFORN");            //no placeholder
            Helpers.Invalid("RESTRICTED");                  //not a US level
            Helpers.Invalid("SECRET//REL TO GBR, USA");     //USA first
            Helpers.Invalid("SECRET//SAR-ABC/SAR-DEF");     //indicator not repeated
        }
    }

    [TestClass]
    public static class RollUpExamples
    {
        public static void Classification()
        {
            Assert.AreEqual("TOP SECRET//NOFORN", Helpers.Roll("(TS//NF)"));
            Assert.AreEqual("SECRET//RELIDO", Helpers.Roll("(S//RELIDO)"));
            Assert.AreEqual("CONFIDENTIAL", Helpers.Roll("(C)"));
            Assert.AreEqual("UNCLASSIFIED", Helpers.Roll("(U)"));
            Assert.AreEqual("SECRET//REL TO USA, FVEY", Helpers.RollWith("SECRET//REL TO USA, FVEY", "(S//REL)", "(U)"));
        }

        public static void Unclassified()
        {
            Assert.AreEqual("SECRET//REL TO USA, JPN", Helpers.Roll("(U//FOUO/REL TO USA, JPN)", "(S//REL TO USA, JPN)"));
            Assert.AreEqual("SECRET//NOFORN/FISA", Helpers.Roll("(U//FISA)", "(S//REL TO USA, GBR)"));
            Assert.AreEqual("UNCLASSIFIED//NOFORN//SBU", Helpers.Roll("(U//SBU-NF)", "(U//REL TO USA, FVEY)"));
            Assert.AreEqual("UNCLASSIFIED//REL TO USA, FVEY", Helpers.Roll("(U)", "(U//REL TO USA, FVEY)"));
        }

        public static void Joint()
        {
            Assert.AreEqual("//JOINT SECRET CAN GBR USA//REL TO USA, CAN, GBR",
                Helpers.RollWith("//JOINT SECRET CAN GBR USA//REL TO USA, CAN, GBR", "(//JOINT S//REL)", "(//JOINT S CAN GBR USA//REL TO USA, FVEY)"));
            Assert.AreEqual("//JOINT SECRET GBR USA//FGI NZL//REL TO USA, GBR",
                Helpers.RollWith("//JOINT SECRET GBR USA//REL TO USA, GBR", "(//JOINT S//REL)", "(//JOINT S GBR USA//REL TO USA, FVEY)", "(//JOINT S GBR USA//FGI NZL//REL TO USA, FVEY)"));
        }

        public static void JointInUsDocument()
        {
            //the Manual prints "SECRET//FGI ISAF NATO//REL TO USA, GBR"; its own rule adds the JOINT co-owners to the FGI list
            Assert.AreEqual("SECRET//FGI GBR ISAF NATO//REL TO USA, GBR",
                Helpers.RollWith("SECRET//REL TO USA, GBR", "(S//REL)", "(//JOINT S GBR USA//REL TO USA, FVEY)", "(//NIS//REL TO USA, GBR, ISAF, NATO)"));
        }

        public static void Sci()
        {
            Assert.AreEqual("TOP SECRET//RSV-123-ABC//NOFORN", Helpers.Roll("(TS//RSV-ABC//NF)", "(TS//RSV-123//NF)"));
            Assert.AreEqual("TOP SECRET//TK-BLFH ABCD-IDIT//NOFORN", Helpers.Roll("(TS//TK-BLFH ABCD//NF)", "(TS//TK-IDIT//NF)"));
            Assert.AreEqual("TOP SECRET//TK-IDIT-KAND ABCD//NOFORN", Helpers.Roll("(TS//TK-KAND ABCD//NF)", "(TS//TK-IDIT//NF)"));
            Assert.AreEqual("SECRET//SI//REL TO USA, FVEY", Helpers.Roll("(S//SI//REL TO USA, FVEY)"));
            Assert.AreEqual("SECRET//TK//RELIDO", Helpers.Roll("(S//TK//RELIDO)"));
        }

        public static void AtomicEnergyAct()
        {
            Assert.AreEqual("SECRET//RD//NOFORN", Helpers.Roll("(S//RD//NF)", "(S//NF)"));
            Assert.AreEqual("SECRET//RD-SIGMA 14 20//NOFORN", Helpers.Roll("(S//RD-SG 20//NF)", "(S//FRD-SG 14//NF)"));
            Assert.AreEqual("SECRET//RD//NOFORN", Helpers.Roll("(S//TFNI//NF)", "(S//RD//NF)"));
            Assert.AreEqual("SECRET//TFNI//REL TO USA, ACGU", Helpers.Roll("(S//TFNI//REL TO USA, ACGU)", "(S//REL TO USA, ACGU)"));
            Assert.AreEqual("UNCLASSIFIED//DOD UCNI", Helpers.Roll("(U//DCNI)"));
        }

        public static void ForeignGovernmentInformation()
        {
            Assert.AreEqual("TOP SECRET//FGI CAN DEU//REL TO USA, CAN, DEU",
                Helpers.Roll("(TS//REL TO USA, CAN, DEU)", "(TS//FGI DEU//REL TO USA, CAN, DEU)", "(//CAN S//REL TO USA, CAN, DEU)"));
            Assert.AreEqual("TOP SECRET//BOHEMIA//FGI AUS CAN DEU NATO//NOFORN",
                Helpers.Roll("(TS//RELIDO)", "(//CAN DEU S//REL TO USA, CAN, DEU)", "(//AUS S//REL TO USA, AUS)", "(//CTS//BOHEMIA//REL TO USA, NATO)"));
            Assert.AreEqual("SECRET//FGI//NOFORN",
                Helpers.Roll("(S//RELIDO)", "(//DEU S//NF)", "(//DEU C//REL TO USA, CAN, DEU)", "(//FGI S//NF)"));
            Assert.AreEqual("TOP SECRET//FGI CAN DEU//NOFORN",
                Helpers.Roll("(S//REL TO USA, AUS)", "(//CAN S//REL TO USA, AUS, CAN, GBR)", "(//DEU TS//NF)"));
            Assert.AreEqual("SECRET//FGI CAN GBR//REL TO USA, CAN, GBR",
                Helpers.Roll("(S//FGI CAN//REL TO USA, CAN, GBR)", "(//GBR S//REL TO USA, CAN, GBR)"));
        }

        public static void Dissemination()
        {
            Assert.AreEqual("TOP SECRET//TK//RSEN/REL TO USA, ACGU", Helpers.Roll("(TS//TK//RS/REL TO USA, ACGU)"));
            Assert.AreEqual("UNCLASSIFIED//FOUO", Helpers.Roll("(U//FOUO)"));
            Assert.AreEqual("TOP SECRET//ORCON/NOFORN", Helpers.Roll("(TS//OC/NF)"));
            Assert.AreEqual("TOP SECRET//ORCON-USGOV/NOFORN", Helpers.Roll("(TS//OC-USGOV/NF)", "(C//NF)"));
            Assert.AreEqual("TOP SECRET//ORCON/NOFORN", Helpers.Roll("(TS//OC-USGOV/NF)", "(S//OC/NF)"));
            Assert.AreEqual("SECRET//IMCON/REL TO USA, AUS, CAN, GBR, NZL", Helpers.Roll("(S//IMC/REL TO USA, AUS, CAN, GBR, NZL)"));
            Assert.AreEqual("TOP SECRET//IMCON/NOFORN", Helpers.Roll("(S//IMC/REL TO USA, AUS, CAN, GBR, NZL)", "(TS//NF)"));
            Assert.AreEqual("SECRET//NOFORN/PROPIN", Helpers.Roll("(S//NF/PR)"));
            Assert.AreEqual("UNCLASSIFIED//DEA SENSITIVE", Helpers.Roll("(U//DSEN)"));
            Assert.AreEqual("SECRET//NOFORN/DEA SENSITIVE", Helpers.Roll("(U//DSEN)", "(S//NF/DSEN)"));
            Assert.AreEqual("TOP SECRET//NOFORN/FISA", Helpers.Roll("(TS//NF/FISA)"));
        }

        public static void Noforn()
        {
            Assert.AreEqual("TOP SECRET//NOFORN", Helpers.Roll("(TS//NF)", "(S//REL TO USA, JPN)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//REL TO USA, JPN)", "(C//RELIDO)"));
            Assert.AreEqual("SECRET//NOFORN//LES", Helpers.Roll("(S//REL TO USA, JPN)", "(U//LES-NF)"));
        }

        public static void RelTo()
        {
            Assert.AreEqual("TOP SECRET//REL TO USA, EGY, ISR", Helpers.RollWith("TOP SECRET//REL TO USA, EGY, ISR", "(TS//REL)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//REL TO USA, FVEY)", "(C)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//REL TO USA, FVEY)", "(C//REL TO USA, ISAF)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//REL TO USA, AUS)", "(C//RELIDO)"));
            Assert.AreEqual("SECRET//REL TO USA, NZL", Helpers.RollWith("SECRET//REL TO USA, NZL", "(S//REL TO USA, JPN, NZL)", "(S//REL)"));
            Assert.AreEqual("SECRET//REL TO USA, AUS, CAN, GBR", Helpers.Roll("(S//REL TO USA, FVEY)", "(S//REL TO USA, ACGU)"));
        }

        public static void Relido()
        {
            Assert.AreEqual("TOP SECRET//TK//RELIDO", Helpers.Roll("(TS//TK//RELIDO)"));
            Assert.AreEqual("SECRET//RELIDO", Helpers.Roll("(S//RELIDO)", "(S//REL TO USA, AUS, CAN/RELIDO)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//RELIDO)", "(S//REL TO USA, AUS, CAN)"));
        }

        public static void EyesOnly()
        {
            Assert.AreEqual("TOP SECRET//USA/CAN/GBR EYES ONLY", Helpers.RollWith("TOP SECRET//USA/CAN/GBR EYES ONLY", "(TS//EYES)"));
        }

        public static void DisplayOnly()
        {
            Assert.AreEqual("SECRET//DISPLAY ONLY AFG", Helpers.Roll("(S//DISPLAY ONLY AFG)", "(S//DISPLAY ONLY AFG)", "(S//DISPLAY ONLY AFG)"));
            Assert.AreEqual("SECRET//DISPLAY ONLY AFG", Helpers.Roll("(S//DISPLAY ONLY AFG)", "(S//DISPLAY ONLY AFG, IRQ)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//DISPLAY ONLY AFG)", "(S//DISPLAY ONLY IRQ)"));
            Assert.AreEqual("SECRET//DISPLAY ONLY IRQ", Helpers.Roll("(S//REL TO USA, IRQ)", "(S//DISPLAY ONLY IRQ)"));
            Assert.AreEqual("SECRET//REL TO USA, IRQ/DISPLAY ONLY AFG", Helpers.Roll("(S//REL TO USA, IRQ/DISPLAY ONLY AFG)", "(S//REL TO USA, IRQ/DISPLAY ONLY AFG)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//REL TO USA, AFG/DISPLAY ONLY IRQ)", "(S//REL TO USA, GBR/DISPLAY ONLY PAK)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(S//DISPLAY ONLY IRQ)", "(S//RELIDO)"));
        }

        public static void NonIcDissemination()
        {
            Assert.AreEqual("UNCLASSIFIED//LIMDIS", Helpers.Roll("(U//DS)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(U//DS)", "(S//NF)"));
            Assert.AreEqual("SECRET//NOFORN//EXDIS", Helpers.Roll("(S//NF//XD)"));
            Assert.AreEqual("SECRET//NOFORN//NODIS", Helpers.Roll("(S//NF//ND)"));
            Assert.AreEqual("UNCLASSIFIED//SBU/SSI", Helpers.Roll("(U//SBU)", "(U//SBU/SSI)", "(U//FOUO)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(U//SBU)", "(S//NF)"));
            Assert.AreEqual("UNCLASSIFIED//SBU NOFORN", Helpers.Roll("(U//SBU-NF)", "(U//FOUO)"));
            Assert.AreEqual("SECRET//NOFORN", Helpers.Roll("(U//SBU-NF)", "(U//FOUO)", "(S//REL TO USA, AUS)"));
            Assert.AreEqual("UNCLASSIFIED//NOFORN//SBU", Helpers.Roll("(U//SBU-NF)", "(U//FOUO/NF)"));
            Assert.AreEqual("UNCLASSIFIED//LES", Helpers.Roll("(U//LES)", "(U)"));
            Assert.AreEqual("UNCLASSIFIED//LES", Helpers.Roll("(U//LES)", "(U//FOUO)"));
            Assert.AreEqual("SECRET//REL TO USA, FVEY//LES", Helpers.Roll("(S//REL TO USA, FVEY)", "(U//REL TO USA, FVEY//LES)"));
            Assert.AreEqual("SECRET//NOFORN//LES", Helpers.Roll("(S//NF)", "(U//LES)"));
            Assert.AreEqual("UNCLASSIFIED//LES NOFORN", Helpers.Roll("(U//LES-NF)", "(U)"));
            Assert.AreEqual("SECRET//NOFORN//LES", Helpers.Roll("(S//REL TO USA, FVEY)", "(U//LES-NF)"));
            Assert.AreEqual("UNCLASSIFIED//NOFORN//SSI", Helpers.Roll("(U//NF//SSI)"));
            Assert.AreEqual("SECRET//REL TO USA, ACGU//SSI", Helpers.Roll("(S//REL TO USA, ACGU)", "(S//REL TO USA, ACGU)", "(U//REL TO USA, ACGU//SSI)"));
        }
    }
}

namespace SystemBannerOffice.Tests
{
    [TestClass]
    public static class ValidatorTests
    {
        private static void Valid(string text)
        {
            List<string> errors = MarkingValidator.Validate(Helpers.Parse(text));
            Assert.IsTrue(errors.Count == 0, text + ": " + string.Join(" ", errors.ToArray()));
        }

        private static void Invalid(string text)
        {
            Assert.IsTrue(MarkingValidator.Validate(Helpers.Parse(text)).Count > 0, text + " should be refused");
        }

        public static void ManualExamplesAreValid()
        {
            Valid("SECRET//HCS-O//ORCON/NOFORN");
            Valid("TOP SECRET//HCS-P JJJ XYZ//ORCON/NOFORN");
            Valid("TOP SECRET//RSV-123-ABC//NOFORN");
            Valid("TOP SECRET//SI-G ABCD EFGH//ORCON/NOFORN");
            Valid("TOP SECRET//TK-BLFH ABCD//NOFORN");
            Valid("SECRET//RD-SIGMA 18 20//NOFORN");
            Valid("UNCLASSIFIED//DOD UCNI");
            Valid("//JOINT SECRET CAN GBR USA//REL TO USA, CAN, GBR");
            Valid("SECRET//NOFORN//EXDIS");
            Valid("SECRET//REL TO USA, IRQ/DISPLAY ONLY AFG");
        }

        public static void RefusesWrongCombinations()
        {
            Invalid("SECRET//FOUO");                        //FOUO is unclassified only
            Invalid("UNCLASSIFIED//SI");                    //SCI is classified only
            Invalid("SECRET//HCS-O//NOFORN");               //HCS-O requires ORCON
            Invalid("TOP SECRET//RSV//NOFORN");             //RESERVE requires a compartment
            Invalid("TOP SECRET//RSV-ABCD//NOFORN");        //RESERVE compartments are 3 characters
            Invalid("TOP SECRET//SI-G ABC//ORCON/NOFORN");  //GAMMA sub-compartments are 4 letters
            Invalid("SECRET//SI-G//ORCON/NOFORN");          //GAMMA is TOP SECRET only
            Invalid("SECRET//TK-IDIT");                     //TK compartments require NOFORN
            Invalid("SECRET//RD-SIGMA 3//NOFORN");          //obsolete SIGMA
            Marking cnwdi = Helpers.Parse("SECRET//FRD//NOFORN");
            cnwdi.Cnwdi = true;                             //CNWDI requires RD
            Assert.IsTrue(MarkingValidator.Validate(cnwdi).Count > 0, "CNWDI without RD should be refused");
            Invalid("SECRET//ORCON/RELIDO");
            Invalid("SECRET//NOFORN/RELIDO");
            Invalid("SECRET//REL TO USA");
            Invalid("SECRET//REL TO USA, GBR//EXDIS");       //EXDIS requires NOFORN
            Invalid("//JOINT SECRET CAN GBR USA//REL TO USA, CAN");
            Invalid("//JOINT SECRET CAN GBR USA//NOFORN");
            Invalid("SECRET//USA/DEU EYES ONLY");
        }

        public static void AuthorityBlock()
        {
            Marking banner = Helpers.Parse("SECRET//NOFORN");
            AuthorityBlock block = new AuthorityBlock();
            block.ClassifiedBy = "J. Doe, Analyst";
            block.DerivedFrom = "Agency X Classification Guide, 20250101";
            block.DeclassifyOn = DateTime.Today.AddYears(10).ToString("yyyyMMdd");
            Assert.AreEqual(0, block.Validate(banner).Count);
            block.DeclassifyOn = "OADR";
            Assert.AreEqual(1, block.Validate(banner).Count);
            block.DeclassifyOn = "25X1-human";
            Assert.IsTrue(block.Validate(banner)[0].Contains("50X1-HUM"), "25X1-human must point to 50X1-HUM");
            Marking rd = Helpers.Parse("SECRET//RD//NOFORN");
            block.DeclassifyOn = "20400101";
            Assert.IsTrue(block.Validate(rd)[0].Contains("N/A to RD portions. See source list for NSI portions."), "RD needs the N/A line");
            block.Original = true;
            block.Reason = "1.4(c)";
            block.DeclassifyOn = DateTime.Today.AddYears(30).ToString("yyyyMMdd");
            Assert.AreEqual(1, block.Validate(banner).Count); //more than 25 years
        }

        public static void MostRestrictiveDeclassification()
        {
            Assert.AreEqual("50X1-HUM", DeclassifyValue.MostRestrictive(new string[] { "20300101", "50X2-WMD", "50X1-HUM" }, null));
            Assert.AreEqual("25X3, 20600101", DeclassifyValue.MostRestrictive(new string[] { "25X3, 20600101", "25X1, 20500101", "20400101" }, null));
            Assert.AreEqual("25X1, 20500101", DeclassifyValue.MostRestrictive(new string[] { "25X4, 20500101", "25X1, 20500101" }, null));
            Assert.AreEqual("25X2, 20600315", DeclassifyValue.MostRestrictive(new string[] { "25X2", "20300101" }, new DateTime[] { new DateTime(2010, 3, 15), DateTime.Today }));
            Assert.AreEqual("N/A to RD and NATO portions. See source list for NSI portions.",
                DeclassifyValue.MostRestrictive(new string[] { "N/A to RD portions. See source list for NSI portions.", "N/A to NATO portions. See source list for NSI portions.", "50X1-HUM" }, null));
        }

        public static void Ceiling()
        {
            SystemSettings settings = new SystemSettings();
            settings.Ceiling = Presets.Find(4); //SECRET computer
            Assert.IsNull(settings.CheckCeiling(Helpers.Parse("SECRET//NOFORN")));
            Assert.IsNotNull(settings.CheckCeiling(Helpers.Parse("TOP SECRET//NOFORN")), "TOP SECRET is above a SECRET computer");
            settings.Ceiling = Presets.Find(5);
            Assert.IsNotNull(settings.CheckCeiling(Helpers.Parse("TOP SECRET//SI//NOFORN")), "SCI needs a TOP SECRET//SCI computer");
            settings.Ceiling = Presets.Find(12); //France
            Assert.IsNotNull(settings.CheckCeiling(Helpers.Parse("SECRET")), "IC markings are not offered on a French computer");
        }
    }
}
