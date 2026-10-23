//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Text;

namespace SystemBannerOffice.Markings
{
    //United Kingdom, Government Security Classifications Policy (GSCP), Cabinet Office: three tiers,
    //and additional markings (the -SENSITIVE marking, handling instructions, descriptors, codewords, prefixes, national caveats)
    public enum UkLevel
    {
        Official = 0,
        Secret = 1,
        TopSecret = 2
    }

    public class UkMarking
    {
        public UkLevel Level = UkLevel.Official;
        public bool Sensitive;                                  //OFFICIAL-SENSITIVE: OFFICIAL only, "immediately after the OFFICIAL classification" (paragraph 25)
        public bool UkPrefix;                                   //assets sent to foreign governments or international organisations
        public bool RelEu;                                      //UK classified information provided to an EU institution, with the UK prefix
        public List<string> Caveats = new List<string>();       //national caveats, SECRET and TOP SECRET only, in alphabetical order
        public string HandlingInstruction;                      //one of UkRegister.HandlingInstructions, "[ORGANISATION] USE ONLY" with its name
        public string Organisation;                             //for "[ORGANISATION] USE ONLY"
        public string Embargo;                                  //EMBARGOED: the time and date, and the handling after it
        public string Codeword;                                 //a single word in capitals
        public List<string> Descriptors = new List<string>();

        public UkMarking Clone()
        {
            UkMarking clone = (UkMarking)MemberwiseClone();
            clone.Caveats = new List<string>(Caveats);
            clone.Descriptors = new List<string>(Descriptors);
            return clone;
        }

        public override string ToString()
        {
            return UkFormatter.Marking(this);
        }
    }

    public static class UkRegister
    {
        public const string RecipientsOnly = "RECIPIENTS ONLY";
        public const string ForPublicRelease = "FOR PUBLIC RELEASE";
        public const string OrganisationUseOnly = "USE ONLY"; //written "[ORGANISATION NAME] USE ONLY"
        public const string HmgUseOnly = "HMG USE ONLY";
        public const string Embargoed = "EMBARGOED";

        public static readonly string[] HandlingInstructions = new string[] { RecipientsOnly, ForPublicRelease, OrganisationUseOnly, HmgUseOnly, Embargoed };
        public static readonly string[] Descriptors = new string[] { "PERSONAL DATA", "LEGAL PROFESSIONAL PRIVILEGE (LPP)", "LEGAL", "MARKET SENSITIVE", "COMMERCIAL", "HR / MANAGEMENT" };
        public static readonly string[] Caveats = new string[] { "UK EYES ONLY", "UK/US EYES ONLY", "FIVE EYES ONLY" };

        //the tiers each additional marking "can be applied to"; 0 OFFICIAL, 1 OFFICIAL-SENSITIVE, 2 SECRET, 3 TOP SECRET
        public static bool Allowed(string marking, int tier)
        {
            switch (marking)
            {
                case RecipientsOnly: return tier >= 1;
                case ForPublicRelease: return tier == 0;
                case OrganisationUseOnly: return true;
                case HmgUseOnly: return tier <= 2;
                case Embargoed: return tier <= 2;
                case "LEGAL": return tier <= 1;
                default: return true; //the other descriptors: OFFICIAL to TOP SECRET
            }
        }

        public static int Tier(UkMarking m)
        {
            if (m.Level == UkLevel.TopSecret) return 3;
            if (m.Level == UkLevel.Secret) return 2;
            return m.Sensitive ? 1 : 0;
        }
    }

    public static class UkFormatter
    {
        public static string Classification(UkMarking m)
        {
            switch (m.Level)
            {
                case UkLevel.TopSecret: return "TOP SECRET";
                case UkLevel.Secret: return "SECRET";
                default: return m.Sensitive ? "OFFICIAL-SENSITIVE" : "OFFICIAL";
            }
        }

        public static string HandlingInstruction(UkMarking m)
        {
            if (string.IsNullOrEmpty(m.HandlingInstruction)) return "";
            if (m.HandlingInstruction == UkRegister.OrganisationUseOnly)
            {
                return (string.IsNullOrEmpty(m.Organisation) ? "[ORGANISATION]" : m.Organisation.ToUpperInvariant()) + " USE ONLY";
            }
            return m.HandlingInstruction;
        }

        //[UK] CLASSIFICATION[-SENSITIVE] [REL EU] [NATIONAL CAVEATS] [HANDLING INSTRUCTION] [CODEWORD] [DESCRIPTORS]:
        //-SENSITIVE comes right after OFFICIAL, handling instructions before descriptors and codewords (paragraphs 25, descriptors, codewords);
        //the place of national caveats and of REL EU, shown in pictures of the GSCP, is a choice of this project
        public static string Marking(UkMarking m)
        {
            List<string> parts = new List<string>();
            if (m.UkPrefix || m.RelEu) parts.Add("UK");
            parts.Add(Classification(m));
            if (m.RelEu) parts.Add("REL EU");
            parts.AddRange(SortedCaveats(m));
            string handling = HandlingInstruction(m);
            if (handling.Length > 0) parts.Add(handling);
            if (!string.IsNullOrEmpty(m.Codeword)) parts.Add(m.Codeword.ToUpperInvariant());
            foreach (string descriptor in UkRegister.Descriptors)
            {
                if (m.Descriptors.Contains(descriptor)) parts.Add(descriptor);
            }
            return string.Join(" ", parts.ToArray());
        }

        //"If multiple National caveats are applied, they must be placed in alphabetical order"
        public static List<string> SortedCaveats(UkMarking m)
        {
            List<string> caveats = new List<string>(m.Caveats);
            caveats.Sort(string.CompareOrdinal);
            return caveats;
        }
    }

    public static class UkValidator
    {
        public static List<string> Validate(UkMarking m)
        {
            List<string> errors = new List<string>();
            int tier = UkRegister.Tier(m);
            if (m.Sensitive && m.Level != UkLevel.Official)
            {
                errors.Add(Text.T("-SENSITIVE marks OFFICIAL information only.", "-SENSITIVE ne s'applique qu'à de l'information OFFICIAL."));
            }
            if (m.Caveats.Count > 0 && m.Level == UkLevel.Official)
            {
                errors.Add(Text.T("National caveats (UK EYES ONLY...) can only be applied to SECRET and TOP SECRET.", "Les caveats nationaux (UK EYES ONLY...) ne s'appliquent qu'à SECRET et TOP SECRET."));
            }
            if (!string.IsNullOrEmpty(m.HandlingInstruction))
            {
                if (!UkRegister.Allowed(m.HandlingInstruction, tier))
                {
                    errors.Add(Text.T(UkFormatter.HandlingInstruction(m) + " cannot be applied to " + UkFormatter.Classification(m) + ".", UkFormatter.HandlingInstruction(m) + " ne s'applique pas à " + UkFormatter.Classification(m) + "."));
                }
                if (m.HandlingInstruction == UkRegister.OrganisationUseOnly && string.IsNullOrEmpty(m.Organisation))
                {
                    errors.Add(Text.T("[ORGANISATION] USE ONLY names the organisation(s).", "[ORGANISATION] USE ONLY nomme la ou les organisations."));
                }
                if (m.HandlingInstruction == UkRegister.Embargoed && string.IsNullOrEmpty(m.Embargo))
                {
                    errors.Add(Text.T("EMBARGOED: the time and date, and the handling after it, are shown as close as possible to the classification.", "EMBARGOED : l'heure et la date, et le traitement ensuite, figurent au plus près de la classification."));
                }
            }
            foreach (string descriptor in m.Descriptors)
            {
                if (!UkRegister.Allowed(descriptor, tier))
                {
                    errors.Add(Text.T(descriptor + " cannot be applied to " + UkFormatter.Classification(m) + ".", descriptor + " ne s'applique pas à " + UkFormatter.Classification(m) + "."));
                }
            }
            if (!string.IsNullOrEmpty(m.Codeword))
            {
                foreach (char c in m.Codeword)
                {
                    if (c < 'A' || c > 'Z')
                    {
                        errors.Add(Text.T("A codeword is a single word in capital letters.", "Un codeword est un seul mot en majuscules."));
                        break;
                    }
                }
            }
            if (m.RelEu && !m.UkPrefix)
            {
                errors.Add(Text.T("REL EU goes with the UK prefix.", "REL EU s'accompagne du préfixe UK."));
            }
            return errors;
        }
    }

    public static class UkSettings
    {
        public const string UnitedKingdom = "United Kingdom";

        //British markings: no Simple value, or a British one (31 to 34)
        public static bool Available(SystemSettings settings)
        {
            return settings.Ceiling == null || settings.Ceiling.ClassificationSystem == UnitedKingdom;
        }

        public static string CheckCeiling(SystemSettings settings, UkMarking marking)
        {
            if (!Available(settings))
            {
                return Text.T("This computer is set to the " + settings.Ceiling.ClassificationSystem + " system (" + settings.Ceiling.Text + "): British markings are not offered.",
                    "Ce poste est réglé sur le système " + settings.Ceiling.ClassificationSystem + " (" + settings.Ceiling.Text + ") : les marquages britanniques n'y sont pas proposés.");
            }
            //presets 31 to 34: OFFICIAL rank 0, OFFICIAL-SENSITIVE 1, SECRET 2, TOP SECRET 3
            if (settings.Ceiling != null && UkRegister.Tier(marking) > settings.Ceiling.Rank)
            {
                return Text.T("This computer is set to " + settings.Ceiling.Text + ": a " + UkFormatter.Classification(marking) + " document cannot be written on it.",
                    "Ce poste est réglé au niveau " + settings.Ceiling.Text + " : un document " + UkFormatter.Classification(marking) + " ne peut pas y être rédigé.");
            }
            return null;
        }
    }
}
