//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;

namespace SystemBannerOffice.Markings
{
    //classification levels, lowest first; RESTRICTED only exists for non-US and NATO information
    public enum Level
    {
        Unclassified = 0,
        Restricted = 1,
        Confidential = 2,
        Secret = 3,
        TopSecret = 4
    }

    //who owns the classification, see IC Markings System Register and Manual, A.6 Formatting
    public enum Owner
    {
        UnitedStates, //SECRET//...
        NonUS,        //  //GBR SECRET//...
        Nato,         //  //NATO SECRET//...
        Joint         //  //JOINT SECRET CAN GBR USA//...
    }

    //one compartment of an SCI control system or of a SAP program, with its sub-compartments
    public class Compartment
    {
        public string Name;
        public List<string> SubCompartments = new List<string>();
        public Compartment(string name, params string[] subCompartments)
        {
            Name = name;
            SubCompartments.AddRange(subCompartments);
        }
        public Compartment Clone()
        {
            return new Compartment(Name, SubCompartments.ToArray());
        }
    }

    //an SCI control system (SI, TK, HCS, 123...) or a SAP program identifier (ABC...), with its compartments
    public class ControlSystem
    {
        public string Name;
        public string Abbreviation;  //SAP programs only: the authorized abbreviation of the program identifier (BP for BUTTER POPCORN), null when there is none
        public List<Compartment> Compartments = new List<Compartment>();
        public ControlSystem(string name, params Compartment[] compartments)
        {
            Name = name;
            Compartments.AddRange(compartments);
        }
        public ControlSystem Clone()
        {
            ControlSystem clone = new ControlSystem(Name);
            clone.Abbreviation = Abbreviation;
            foreach (Compartment compartment in Compartments)
            {
                clone.Compartments.Add(compartment.Clone());
            }
            return clone;
        }
        public Compartment FindCompartment(string name)
        {
            foreach (Compartment compartment in Compartments)
            {
                if (compartment.Name == name)
                {
                    return compartment;
                }
            }
            return null;
        }
    }

    //the classification and control markings of a portion or of a whole document, in the categories of the Register:
    //classification, SCI, SAP, AEA, FGI, dissemination controls, non-IC dissemination controls
    public class Marking
    {
        public Owner Owner = Owner.UnitedStates;
        public Level Level = Level.Unclassified;
        public string OwnerCode;                                  //trigraph(s) or tetragraph(s) of a non-US owner, space separated ("GBR", "CAN GBR"), null for concealed FGI ("//FGI S")
        public List<string> JointOwners = new List<string>();     //co-owners of JOINT information, USA among them; empty in a "(//JOINT S//REL)" portion: the co-owners of the banner
        public string NatoActivity;                              //NAC activity of NATO/NAC markings, null for plain NATO markings
        public List<string> NatoSap = new List<string>();         //ATOMAL (written in the AEA category), BALK and BOHEMIA (written in the SAP category)

        public List<ControlSystem> Sci = new List<ControlSystem>();
        public List<ControlSystem> Sap = new List<ControlSystem>();
        public bool SapProgramsHidden;                           //"SAR-" category with no program identifier shown, see the optional SAP hierarchy

        //Atomic Energy Act markings, RD and FRD with their SIGMA numbers
        public bool RestrictedData;
        public bool Cnwdi;
        public List<int> RdSigma = new List<int>();
        public bool FormerlyRestrictedData;
        public List<int> FrdSigma = new List<int>();
        public bool DodUcni;
        public bool DoeUcni;
        public bool Tfni;

        //foreign government information inside US information
        public bool Fgi;
        public List<string> FgiSources = new List<string>();     //empty when the source is concealed: banner "FGI" only

        //dissemination controls, by Register key (see Register.Dissemination), and the lists some of them carry
        public List<string> Dissem = new List<string>();
        public List<string> RelTo = new List<string>();          //USA first
        public List<string> EyesOnly = new List<string>();       //USA first
        public List<string> DisplayOnly = new List<string>();

        //non-IC dissemination controls, by Register key (see Register.NonIcDissemination)
        public List<string> NonIcDissem = new List<string>();

        public bool HasDissem(string key)
        {
            return Dissem.Contains(key);
        }

        public bool HasNonIcDissem(string key)
        {
            return NonIcDissem.Contains(key);
        }

        public bool HasAea
        {
            get { return RestrictedData || FormerlyRestrictedData || DodUcni || DoeUcni || Tfni; }
        }

        //true when anything follows the classification
        public bool HasControls
        {
            get
            {
                return Sci.Count > 0 || Sap.Count > 0 || SapProgramsHidden || HasAea || Fgi || Dissem.Count > 0 || NonIcDissem.Count > 0 || NatoSap.Count > 0;
            }
        }

        public bool IsUsDocument
        {
            get { return Owner == Owner.UnitedStates; }
        }

        public Marking Clone()
        {
            Marking clone = (Marking)MemberwiseClone();
            clone.JointOwners = new List<string>(JointOwners);
            clone.NatoSap = new List<string>(NatoSap);
            clone.Sci = new List<ControlSystem>();
            foreach (ControlSystem system in Sci) clone.Sci.Add(system.Clone());
            clone.Sap = new List<ControlSystem>();
            foreach (ControlSystem program in Sap) clone.Sap.Add(program.Clone());
            clone.RdSigma = new List<int>(RdSigma);
            clone.FrdSigma = new List<int>(FrdSigma);
            clone.FgiSources = new List<string>(FgiSources);
            clone.Dissem = new List<string>(Dissem);
            clone.RelTo = new List<string>(RelTo);
            clone.EyesOnly = new List<string>(EyesOnly);
            clone.DisplayOnly = new List<string>(DisplayOnly);
            clone.NonIcDissem = new List<string>(NonIcDissem);
            return clone;
        }

        public override string ToString()
        {
            return MarkingFormatter.Banner(this);
        }
    }
}
