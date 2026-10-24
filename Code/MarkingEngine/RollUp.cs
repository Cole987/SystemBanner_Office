//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;

namespace SystemBannerOffice.Markings
{
    public class RollUpOptions
    {
        //disseminated analytic product: every classified portion must carry an explicit FD&R marking (ICD 710 E.5)
        public bool DisseminatedAnalyticProduct;
        //classified portions with no FD&R created before 28 June 2010 are handled as NOFORN, after it as RELIDO (Table 2)
        public DateTime DateOfOrigin = DateTime.Today;
        //the banner line already on the document: "REL", "EYES" and "(//JOINT S//REL)" portions refer to its lists
        public Marking PreviousBanner;
    }

    public class RollUpResult
    {
        public Marking Banner;
        public List<string> Errors = new List<string>();
        public List<string> Notes = new List<string>(); //choices the roll-up made where the Register leaves room, shown to the classifier
    }

    //builds the banner line from the portion marks, Register and Manual section D.2 and Table 3, and the roll-up rules of each marking template
    public static class RollUp
    {
        public static readonly DateTime RelidoDefaultDate = new DateTime(2010, 6, 28);

        //tetragraphs Table 3 rule 23 allows to expand to find the countries common to several REL TO lists;
        //other tetragraphs (NATO, ISAF...) are only common to lists that name them
        public static readonly Dictionary<string, string[]> Expansions = CreateExpansions();

        private static Dictionary<string, string[]> CreateExpansions()
        {
            Dictionary<string, string[]> expansions = new Dictionary<string, string[]>();
            expansions["FVEY"] = new string[] { "USA", "AUS", "CAN", "GBR", "NZL" };
            expansions["ACGU"] = new string[] { "USA", "AUS", "CAN", "GBR" };
            expansions["TEYE"] = new string[] { "USA", "AUS", "GBR" };
            return expansions;
        }

        public static RollUpResult Compute(IList<Marking> portions, RollUpOptions options)
        {
            if (options == null)
            {
                options = new RollUpOptions();
            }
            RollUpResult result = new RollUpResult();
            Marking banner = new Marking();
            result.Banner = banner;
            if (portions.Count == 0)
            {
                result.Errors.Add(Text.T("The document has no portion marks to roll up.", "Le document ne comporte aucune marque de portion à consolider.", "Das Dokument enthält keine Teil-Kennzeichnungen, die zusammengeführt werden können."));
                return result;
            }
            List<Marking> resolved = Resolve(portions, options.PreviousBanner, result);
            Ownership(resolved, banner, result);
            Controls(resolved, banner);
            Fgi(resolved, banner);
            bool classified = banner.Level > Level.Unclassified;
            ForeignDisclosure fdr = ComputeForeignDisclosure(resolved, classified, options, result);
            NonFdrControls(resolved, banner, classified, fdr);
            ApplyForeignDisclosure(banner, fdr);
            if (banner.Owner == Owner.Joint && !banner.HasDissem("REL"))
            {
                result.Errors.Add(Text.T("JOINT information must be releasable to all its co-owners (REL TO), it cannot roll up to " + (banner.HasDissem("NF") ? "NOFORN" : "no FD&R marking") + ".", "Une information JOINT doit être diffusable à tous ses copropriétaires (REL TO) : elle ne peut pas aboutir à " + (banner.HasDissem("NF") ? "NOFORN" : "une absence de marquage FD&R") + ".", "JOINT-Informationen müssen für alle Miteigentümer freigegeben sein (REL TO), sie können nicht zu " + (banner.HasDissem("NF") ? "NOFORN" : "keiner FD&R-Kennzeichnung") + " zusammengeführt werden."));
            }
            return result;
        }

        //copies the portions, filling "REL", "EYES" and JOINT portions with no co-owners from the banner already on the document
        private static List<Marking> Resolve(IList<Marking> portions, Marking previous, RollUpResult result)
        {
            List<Marking> resolved = new List<Marking>();
            bool missing = false;
            foreach (Marking portion in portions)
            {
                Marking copy = portion.Clone();
                if (copy.HasDissem("REL") && copy.RelTo.Count == 0)
                {
                    if (previous != null && previous.RelTo.Count > 0)
                    {
                        copy.RelTo.AddRange(previous.RelTo);
                    }
                    else
                    {
                        missing = true;
                    }
                }
                if (copy.HasDissem("EYES") && copy.EyesOnly.Count == 0 && previous != null)
                {
                    copy.EyesOnly.AddRange(previous.EyesOnly);
                }
                if (copy.Owner == Owner.Joint && copy.JointOwners.Count == 0 && previous != null)
                {
                    copy.JointOwners.AddRange(previous.JointOwners);
                }
                resolved.Add(copy);
            }
            if (missing)
            {
                result.Errors.Add(Text.T("A portion is marked \"REL\" but the document has no REL TO list yet: write the countries in that portion.", "Une portion est marquée « REL », mais le document n'a pas encore de liste REL TO : indiquer les pays dans cette portion.", "Ein Abschnitt ist mit \"REL\" gekennzeichnet, aber das Dokument hat noch keine REL TO-Liste: Tragen Sie die Länder in diesem Abschnitt ein."));
            }
            return resolved;
        }

        //who owns the document and its classification level
        private static void Ownership(List<Marking> portions, Marking banner, RollUpResult result)
        {
            bool anyUs = false, allJoint = true, allNato = true, allNonUs = true;
            string nonUsOwner = null;
            List<string> jointOwners = null;
            string natoActivity = portions[0].NatoActivity;
            bool sameJoint = true, sameNonUs = true, sameActivity = true;
            Level level = Level.Unclassified;
            foreach (Marking p in portions)
            {
                anyUs |= p.Owner == Owner.UnitedStates;
                allJoint &= p.Owner == Owner.Joint;
                allNato &= p.Owner == Owner.Nato;
                allNonUs &= p.Owner == Owner.NonUS && p.OwnerCode != null;
                if (p.Owner == Owner.Joint)
                {
                    if (jointOwners == null) jointOwners = p.JointOwners;
                    else sameJoint &= MarkingFormatter.SameList(jointOwners, p.JointOwners);
                }
                if (p.Owner == Owner.NonUS && p.OwnerCode != null)
                {
                    if (nonUsOwner == null) nonUsOwner = p.OwnerCode;
                    else sameNonUs &= nonUsOwner == p.OwnerCode;
                }
                sameActivity &= p.NatoActivity == natoActivity;
                if (p.Level > level)
                {
                    level = p.Level;
                }
            }
            if (!anyUs && allJoint && sameJoint)
            {
                banner.Owner = Owner.Joint;
                banner.JointOwners.AddRange(jointOwners);
            }
            else if (!anyUs && allNato)
            {
                banner.Owner = Owner.Nato;
                banner.NatoActivity = sameActivity ? natoActivity : null;
            }
            else if (!anyUs && allNonUs && sameNonUs)
            {
                banner.Owner = Owner.NonUS;
                banner.OwnerCode = nonUsOwner;
            }
            else
            {
                banner.Owner = Owner.UnitedStates;
                if (level == Level.Restricted)
                {
                    //the United States has no RESTRICTED level: foreign RESTRICTED information is protected at least as CONFIDENTIAL
                    level = Level.Confidential;
                    result.Notes.Add(Text.T("Foreign RESTRICTED portions roll up to a US banner as CONFIDENTIAL; check the handling agreed with the source country.", "Les portions RESTRICTED étrangères se consolident en CONFIDENTIAL dans un bandeau américain ; vérifier le traitement convenu avec le pays d'origine.", "Ausländische RESTRICTED-Abschnitte werden in einer US-Bannerzeile zu CONFIDENTIAL zusammengeführt; prüfen Sie die mit dem Herkunftsland vereinbarte Behandlung."));
                }
            }
            banner.Level = level;
        }

        //SCI, SAP, AEA and NATO special access markings: every unique value of the portions
        private static void Controls(List<Marking> portions, Marking banner)
        {
            foreach (Marking p in portions)
            {
                MergeSystems(banner.Sci, p.Sci);
                MergeSystems(banner.Sap, p.Sap);
                banner.SapProgramsHidden |= p.SapProgramsHidden;
                banner.RestrictedData |= p.RestrictedData;
                banner.Cnwdi |= p.Cnwdi;
                AddAll(banner.RdSigma, p.RdSigma);
                banner.FormerlyRestrictedData |= p.FormerlyRestrictedData;
                AddAll(banner.FrdSigma, p.FrdSigma);
                banner.DodUcni |= p.DodUcni;
                banner.DoeUcni |= p.DoeUcni;
                banner.Tfni |= p.Tfni;
                AddAll(banner.NatoSap, p.NatoSap);
            }
            if (banner.Sap.Count > 0)
            {
                banner.SapProgramsHidden = false; //the program identifiers already show the category
            }
            //RD supersedes FRD and TFNI, and takes the SIGMA numbers of FRD; FRD supersedes TFNI
            if (banner.RestrictedData)
            {
                AddAll(banner.RdSigma, banner.FrdSigma);
                banner.FrdSigma.Clear();
                banner.FormerlyRestrictedData = false;
                banner.Tfni = false;
            }
            else if (banner.FormerlyRestrictedData)
            {
                banner.Tfni = false;
            }
            if (banner.Level > Level.Unclassified)
            {
                //DOD and DOE UCNI only apply to unclassified information and never show in a classified banner
                banner.DodUcni = false;
                banner.DoeUcni = false;
            }
        }

        private static void MergeSystems(List<ControlSystem> target, List<ControlSystem> source)
        {
            foreach (ControlSystem system in source)
            {
                ControlSystem existing = null;
                foreach (ControlSystem candidate in target)
                {
                    if (candidate.Name == system.Name)
                    {
                        existing = candidate;
                        break;
                    }
                }
                if (existing == null)
                {
                    target.Add(system.Clone());
                    continue;
                }
                if (existing.Abbreviation == null)
                {
                    existing.Abbreviation = system.Abbreviation;
                }
                foreach (Compartment compartment in system.Compartments)
                {
                    Compartment existingCompartment = existing.FindCompartment(compartment.Name);
                    if (existingCompartment == null)
                    {
                        existing.Compartments.Add(compartment.Clone());
                    }
                    else
                    {
                        AddAll(existingCompartment.SubCompartments, compartment.SubCompartments);
                    }
                }
            }
        }

        private static void AddAll<T>(List<T> target, List<T> source)
        {
            foreach (T value in source)
            {
                if (!target.Contains(value))
                {
                    target.Add(value);
                }
            }
        }

        //FGI in a US banner: the codes of every foreign, NATO and JOINT portion, or "FGI" alone when one source is concealed
        private static void Fgi(List<Marking> portions, Marking banner)
        {
            if (banner.Owner != Owner.UnitedStates)
            {
                //a non-US, NATO or JOINT banner keeps the FGI of its own portions
                foreach (Marking p in portions)
                {
                    if (p.Fgi)
                    {
                        banner.Fgi = true;
                        AddAll(banner.FgiSources, p.FgiSources);
                    }
                }
                return;
            }
            bool concealed = false;
            foreach (Marking p in portions)
            {
                switch (p.Owner)
                {
                    case Owner.NonUS:
                        banner.Fgi = true;
                        if (p.OwnerCode == null)
                        {
                            concealed = true;
                        }
                        else
                        {
                            AddAll(banner.FgiSources, new List<string>(p.OwnerCode.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)));
                        }
                        break;
                    case Owner.Nato:
                        banner.Fgi = true;
                        AddAll(banner.FgiSources, new List<string>(new string[] { "NATO" }));
                        if (p.NatoActivity != null && MarkingParser.IsCountryCode(p.NatoActivity))
                        {
                            //NATO/NAC information also names its activity: "(//NIS)" rolls up to "FGI ISAF NATO"
                            AddAll(banner.FgiSources, new List<string>(new string[] { p.NatoActivity }));
                        }
                        break;
                    case Owner.Joint:
                        banner.Fgi = true;
                        foreach (string owner in p.JointOwners)
                        {
                            if (owner != "USA")
                            {
                                AddAll(banner.FgiSources, new List<string>(new string[] { owner }));
                            }
                        }
                        break;
                }
                if (p.Fgi)
                {
                    banner.Fgi = true;
                    if (p.FgiSources.Count == 0)
                    {
                        concealed = true;
                    }
                    AddAll(banner.FgiSources, p.FgiSources);
                }
            }
            if (concealed)
            {
                banner.FgiSources.Clear();
            }
        }

        private static readonly string[] FdrKeys = new string[] { "NF", "REL", "RELIDO", "EYES", "DISPLAYONLY" };
        private static readonly string[] IcCaveats = new string[] { "RS", "FOUO", "OC", "OC-USGOV", "IMC", "PR", "FISA" };
        private static readonly string[] NonIcCaveats = new string[] { "DSEN" }; //DSEN is an IC control handled like the non-IC ones for FD&R

        //the foreign disclosure and release outcome of the whole document
        private class ForeignDisclosure
        {
            public bool Any;               //false: no FD&R marking in the banner
            public bool NoForn;
            public bool Relido;
            public List<string> RelTo;     //null when the banner has no REL TO
            public List<string> Eyes;      //null when the banner has no EYES ONLY
            public List<string> Display;   //null when the banner has no DISPLAY ONLY
            public bool OtherExplicitFdr;  //a portion carries explicit IC FD&R besides SBU-NF and LES-NF
        }

        //what one portion says about release, after Table 2 for portions with no FD&R marking
        private class PortionRelease
        {
            public bool NoForn;
            public bool Relido;
            public List<string> Rel;      //REL TO list, EYES ONLY lists count as REL TO lists when they meet one
            public List<string> Eyes;
            public List<string> Display;
        }

        private static bool HasExplicitFdr(Marking p)
        {
            foreach (string key in FdrKeys)
            {
                if (p.HasDissem(key))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool HasNoFornNonIc(Marking p)
        {
            return p.HasNonIcDissem("SBU-NF") || p.HasNonIcDissem("LES-NF") || p.HasNonIcDissem("XD") || p.HasNonIcDissem("ND");
        }

        private static bool IsForeign(Marking p)
        {
            return p.Owner == Owner.NonUS || p.Owner == Owner.Nato || p.Fgi;
        }

        //caveated: no FD&R marking but at least one AEA, SAP or dissemination control (SCI alone does not count), Table 2
        private static bool IsCaveated(Marking p)
        {
            return p.HasAea || p.Sap.Count > 0 || p.SapProgramsHidden || p.NatoSap.Count > 0 || p.Dissem.Count > 0 || p.NonIcDissem.Count > 0;
        }

        private static bool HasIcCaveat(Marking p)
        {
            foreach (string key in IcCaveats)
            {
                if (p.HasDissem(key))
                {
                    return true;
                }
            }
            return false;
        }

        private static ForeignDisclosure ComputeForeignDisclosure(List<Marking> portions, bool classified, RollUpOptions options, RollUpResult result)
        {
            ForeignDisclosure fdr = new ForeignDisclosure();
            bool anyExplicit = false;
            foreach (Marking p in portions)
            {
                if (HasExplicitFdr(p) || HasNoFornNonIc(p))
                {
                    anyExplicit = true;
                }
                if (HasExplicitFdr(p))
                {
                    fdr.OtherExplicitFdr = true;
                }
            }
            if (!anyExplicit && !options.DisseminatedAnalyticProduct)
            {
                return fdr; //no portion has FD&R: the banner follows the classification and control roll-up only
            }
            List<PortionRelease> releases = new List<PortionRelease>();
            bool assumedRelido = false, assumedNoForn = false;
            foreach (Marking p in portions)
            {
                bool portionClassified = p.Level > Level.Unclassified;
                PortionRelease release = new PortionRelease();
                if (HasExplicitFdr(p) || HasNoFornNonIc(p))
                {
                    release.NoForn = p.HasDissem("NF") || HasNoFornNonIc(p);
                    release.Relido = p.HasDissem("RELIDO");
                    if (p.HasDissem("REL")) release.Rel = new List<string>(p.RelTo);
                    if (p.HasDissem("EYES")) release.Eyes = new List<string>(p.EyesOnly);
                    if (p.HasDissem("DISPLAYONLY")) release.Display = new List<string>(p.DisplayOnly);
                }
                else if (portionClassified)
                {
                    if (options.DisseminatedAnalyticProduct)
                    {
                        result.Errors.Add(Text.T("Portion (" + MarkingFormatter.Portion(p) + ") has no FD&R marking: every classified portion of a disseminated analytic product needs NOFORN, REL TO, RELIDO or DISPLAY ONLY.", "La portion (" + MarkingFormatter.Portion(p) + ") n'a pas de marquage FD&R : toute portion classifiée d'un produit analytique diffusé exige NOFORN, REL TO, RELIDO ou DISPLAY ONLY.", "Abschnitt (" + MarkingFormatter.Portion(p) + ") hat keine FD&R-Kennzeichnung: Jeder eingestufte Abschnitt eines verteilten Analyseprodukts benötigt NOFORN, REL TO, RELIDO oder DISPLAY ONLY."));
                    }
                    //Table 2: FGI and caveated information as NOFORN, uncaveated information as RELIDO since 28 June 2010
                    if (IsForeign(p) || IsCaveated(p) || options.DateOfOrigin < RelidoDefaultDate)
                    {
                        release.NoForn = true;
                        assumedNoForn = true;
                    }
                    else
                    {
                        release.Relido = true;
                        assumedRelido = true;
                    }
                }
                else if (classified && (p.DodUcni || p.DoeUcni || p.NonIcDissem.Count > 0 || p.HasDissem("DSEN")))
                {
                    //unclassified non-IC information in a classified document: NOFORN until the originator approves release
                    release.NoForn = true;
                    assumedNoForn = true;
                }
                else if (classified && HasIcCaveat(p))
                {
                    //unclassified caveated IC information, as (U//FISA) in the UNCLASSIFIED template: considered RELIDO
                    release.Relido = true;
                    assumedRelido = true;
                }
                else
                {
                    continue; //unclassified portions with no control and no FD&R do not count
                }
                releases.Add(release);
            }
            if (assumedNoForn)
            {
                result.Notes.Add(Text.T("Portions with no FD&R marking were handled as NOFORN (FGI, caveated or older than 28 June 2010 information, Table 2).", "Les portions sans marquage FD&R ont été traitées comme NOFORN (informations FGI, avec caveat ou antérieures au 28 juin 2010, Table 2).", "Abschnitte ohne FD&R-Kennzeichnung wurden als NOFORN behandelt (FGI, Informationen mit Warnvermerk oder vor dem 28. Juni 2010, Table 2)."));
            }
            if (assumedRelido)
            {
                result.Notes.Add(Text.T("Portions with no FD&R marking were handled as RELIDO (uncaveated information since 28 June 2010, Table 2).", "Les portions sans marquage FD&R ont été traitées comme RELIDO (informations sans caveat depuis le 28 juin 2010, Table 2).", "Abschnitte ohne FD&R-Kennzeichnung wurden als RELIDO behandelt (Informationen ohne Warnvermerk seit dem 28. Juni 2010, Table 2)."));
            }
            if (releases.Count == 0)
            {
                return fdr;
            }
            fdr.Any = true;
            //rules 1 to 8: one NOFORN portion makes the document NOFORN
            foreach (PortionRelease release in releases)
            {
                if (release.NoForn)
                {
                    fdr.NoForn = true;
                    return fdr;
                }
            }
            //rules 10, 12, 17, 18, 24: RELIDO stays only when every portion carries it
            bool anyRelido = false, allRelido = true;
            foreach (PortionRelease release in releases)
            {
                anyRelido |= release.Relido;
                allRelido &= release.Relido;
            }
            if (anyRelido && !allRelido)
            {
                fdr.NoForn = true;
                return fdr;
            }
            //EYES ONLY on its own: the countries common to all portions
            bool allEyes = true, anyDisplay = false;
            foreach (PortionRelease release in releases)
            {
                allEyes &= release.Eyes != null && release.Rel == null && release.Display == null;
                anyDisplay |= release.Display != null;
            }
            if (allEyes)
            {
                List<string> lists = Common(releases, delegate (PortionRelease r) { return r.Eyes; });
                if (lists == null)
                {
                    fdr.NoForn = true;
                    return fdr;
                }
                fdr.Eyes = lists;
                return fdr;
            }
            //rules 9, 11, 13, 21 to 23, 25 to 27: REL TO keeps the countries common to every portion, DISPLAY ONLY the countries
            //every portion may at least be shown to (release implies disclosure)
            List<string> rel = Common(releases, delegate (PortionRelease r)
            {
                if (r.Rel == null && r.Eyes == null) return null;
                List<string> list = new List<string>();
                if (r.Rel != null) list.AddRange(r.Rel);
                if (r.Eyes != null) list.AddRange(r.Eyes);
                return list;
            });
            List<string> display = null;
            if (anyDisplay)
            {
                List<List<string>> disclosed = new List<List<string>>();
                foreach (PortionRelease r in releases)
                {
                    List<string> list = new List<string>();
                    if (r.Rel != null) list.AddRange(Expand(r.Rel));
                    if (r.Eyes != null) list.AddRange(Expand(r.Eyes));
                    if (r.Display != null) list.AddRange(Expand(r.Display));
                    disclosed.Add(list);
                }
                List<string> common = Intersect(disclosed);
                display = new List<string>();
                List<string> releasable = rel != null ? Expand(rel) : new List<string>();
                foreach (string code in common)
                {
                    if (code != "USA" && !releasable.Contains(code))
                    {
                        display.Add(code);
                    }
                }
                if (display.Count == 0)
                {
                    display = null;
                }
            }
            if (allRelido)
            {
                //rule 24, and "REL TO [USA, LIST]/RELIDO" when every portion is also releasable to common countries
                fdr.Relido = true;
                fdr.RelTo = rel;
                return fdr;
            }
            if (rel == null && display == null)
            {
                fdr.NoForn = true;
                return fdr;
            }
            fdr.RelTo = rel;
            fdr.Display = display;
            return fdr;
        }

        private delegate List<string> ListOf(PortionRelease release);

        //the codes common to every portion's list, besides USA; null when one portion has no list or nothing but USA is common.
        //identical lists are kept as written (FVEY stays FVEY), others are compared country by country
        private static List<string> Common(List<PortionRelease> releases, ListOf listOf)
        {
            List<List<string>> lists = new List<List<string>>();
            foreach (PortionRelease release in releases)
            {
                List<string> list = listOf(release);
                if (list == null)
                {
                    return null;
                }
                lists.Add(list);
            }
            bool identical = true;
            foreach (List<string> list in lists)
            {
                identical &= MarkingFormatter.SameList(list, lists[0]);
            }
            List<string> common;
            if (identical)
            {
                common = new List<string>(lists[0]);
            }
            else
            {
                List<List<string>> expanded = new List<List<string>>();
                foreach (List<string> list in lists)
                {
                    expanded.Add(Expand(list));
                }
                common = Intersect(expanded);
            }
            bool other = false;
            foreach (string code in common)
            {
                other |= code != "USA";
            }
            if (!other)
            {
                return null;
            }
            if (!common.Contains("USA"))
            {
                common.Insert(0, "USA");
            }
            return Sorting.SortCountries(common, true);
        }

        public static List<string> Expand(List<string> codes)
        {
            List<string> expanded = new List<string>();
            foreach (string code in codes)
            {
                string[] members;
                if (Expansions.TryGetValue(code, out members))
                {
                    foreach (string member in members)
                    {
                        if (!expanded.Contains(member)) expanded.Add(member);
                    }
                }
                else if (!expanded.Contains(code))
                {
                    expanded.Add(code);
                }
            }
            return expanded;
        }

        private static List<string> Intersect(List<List<string>> lists)
        {
            List<string> common = new List<string>(lists[0]);
            for (int i = 1; i < lists.Count; i++)
            {
                common.RemoveAll(delegate (string code) { return !lists[i].Contains(code); });
            }
            return common;
        }

        //dissemination and non-IC controls other than FD&R, with the supersessions of their templates
        private static void NonFdrControls(List<Marking> portions, Marking banner, bool classified, ForeignDisclosure fdr)
        {
            foreach (Marking p in portions)
            {
                foreach (string key in p.Dissem)
                {
                    if (Array.IndexOf(FdrKeys, key) < 0 && !banner.Dissem.Contains(key))
                    {
                        banner.Dissem.Add(key);
                    }
                }
                AddAll(banner.NonIcDissem, p.NonIcDissem);
            }
            List<string> dissem = banner.Dissem;
            List<string> nonIc = banner.NonIcDissem;
            //ORCON supersedes ORCON-USGOV, NODIS supersedes EXDIS
            if (dissem.Contains("OC")) dissem.Remove("OC-USGOV");
            if (nonIc.Contains("ND")) nonIc.Remove("XD");
            if (classified)
            {
                //FOUO, LIMDIS, SBU and SBU NOFORN only show in unclassified banners; LES NOFORN shows as NOFORN and LES
                dissem.Remove("FOUO");
                nonIc.Remove("DS");
                nonIc.Remove("SBU");
                nonIc.Remove("SBU-NF");
                if (nonIc.Remove("LES-NF") && !nonIc.Contains("LES"))
                {
                    nonIc.Add("LES");
                }
            }
            else
            {
                if (nonIc.Contains("XD") || nonIc.Contains("ND"))
                {
                    nonIc.Remove("SBU");
                    nonIc.Remove("SBU-NF");
                }
                if (nonIc.Contains("SBU-NF")) nonIc.Remove("SBU");
                if (nonIc.Contains("LES-NF")) nonIc.Remove("LES");
                //SBU NOFORN and LES NOFORN become NOFORN and SBU or LES when other portions carry an explicit FD&R marking
                if (fdr.OtherExplicitFdr)
                {
                    if (nonIc.Remove("SBU-NF")) nonIc.Add("SBU");
                    if (nonIc.Remove("LES-NF")) nonIc.Add("LES");
                }
                else if (nonIc.Contains("SBU-NF") || nonIc.Contains("LES-NF"))
                {
                    fdr.Any = false; //the NOFORN stays inside SBU NOFORN or LES NOFORN
                }
                //FOUO is superseded by any other dissemination or non-IC control, FD&R markings excepted
                bool other = nonIc.Count > 0 || banner.DodUcni || banner.DoeUcni;
                foreach (string key in dissem)
                {
                    other |= key != "FOUO";
                }
                if (other)
                {
                    dissem.Remove("FOUO");
                }
            }
            Register.SortByRegister(dissem, Register.Dissemination);
            Register.SortByRegister(nonIc, Register.NonIcDissemination);
            //EXDIS and NODIS forbid release: NOFORN
            if (nonIc.Contains("XD") || nonIc.Contains("ND"))
            {
                fdr.Any = true;
                fdr.NoForn = true;
            }
        }

        private static void ApplyForeignDisclosure(Marking banner, ForeignDisclosure fdr)
        {
            if (!fdr.Any)
            {
                return;
            }
            List<string> keys = new List<string>();
            if (fdr.NoForn)
            {
                keys.Add("NF");
            }
            else
            {
                if (fdr.RelTo != null)
                {
                    keys.Add("REL");
                    banner.RelTo.AddRange(fdr.RelTo);
                }
                if (fdr.Relido) keys.Add("RELIDO");
                if (fdr.Eyes != null)
                {
                    keys.Add("EYES");
                    banner.EyesOnly.AddRange(fdr.Eyes);
                }
                if (fdr.Display != null)
                {
                    keys.Add("DISPLAYONLY");
                    banner.DisplayOnly.AddRange(Sorting.SortCountries(fdr.Display, false));
                }
            }
            banner.Dissem.AddRange(keys);
            Register.SortByRegister(banner.Dissem, Register.Dissemination);
        }
    }
}
