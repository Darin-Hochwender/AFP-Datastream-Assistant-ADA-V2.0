using System.Collections.Generic;

namespace AFPEngineer_Explorer.Services
{
    public class SfInfo
    {
        public string Acronym { get; set; }
        public string FriendlyName { get; set; }
    }

    public static class AfpConstants
    {
        private static readonly KeyValuePair<int, string>[] _sfNamesArray = new[] {
            new KeyValuePair<int,string>(0xD3A8C9, "BAG - Begin Active Environment Group"),
            new KeyValuePair<int,string>(0xD3A8EB, "BBC - Begin Bar Code Object"),
            new KeyValuePair<int,string>(0xD3A88A, "BCF - Begin Coded Font"),
            new KeyValuePair<int,string>(0xD3A887, "BCP - Begin Code Page"),
            new KeyValuePair<int,string>(0xD3EEEB, "BDA - Bar Code Data"),
            new KeyValuePair<int,string>(0xD3A6EB, "BDD - Bar Code Data Descriptor"),
            new KeyValuePair<int,string>(0xD3A8C4, "BDG - Begin Document Environment Group"),
            new KeyValuePair<int,string>(0xD3A8A7, "BDI - Begin Document Index"),
            new KeyValuePair<int,string>(0xD3A8A8, "BDT - Begin Document"),
            new KeyValuePair<int,string>(0xD3A8C5, "BFG - Begin Form Environment Group"),
            new KeyValuePair<int,string>(0xD3A8CD, "BFM - Begin Form Map"),
            new KeyValuePair<int,string>(0xD3A889, "BFN - Begin Font"),
            new KeyValuePair<int,string>(0xD3A8BB, "BGR - Begin Graphics Object"),
            new KeyValuePair<int,string>(0xD3A87B, "BII - Begin IM Image"),
            new KeyValuePair<int,string>(0xD3A8FB, "BIM - Begin Image Object"),
            new KeyValuePair<int,string>(0xD3A8CC, "BMM - Begin Medium Map"),
            new KeyValuePair<int,string>(0xD3A8DF, "BMO - Begin Overlay"),
            new KeyValuePair<int,string>(0xD3A8AD, "BNG - Begin Named Page Group"),
            new KeyValuePair<int,string>(0xD3A892, "BOC - Begin Object Container"),
            new KeyValuePair<int,string>(0xD3A8C7, "BOG - Begin Object Environment Group"),
            new KeyValuePair<int,string>(0xD3A8A5, "BPF - Begin Print File"),
            new KeyValuePair<int,string>(0xD3A8AF, "BPG - Begin Page"),
            new KeyValuePair<int,string>(0xD3A85F, "BPS - Begin Page Segment"),
            new KeyValuePair<int,string>(0xD3A89B, "BPT - Begin Presentation Text Object"),
            new KeyValuePair<int,string>(0xD3A8C6, "BRG - Begin Resource Group"),
            new KeyValuePair<int,string>(0xD3A8CE, "BRS - Begin Resource"),
            new KeyValuePair<int,string>(0xD3A8D9, "BSG - Begin Resource Environment Group"),
            new KeyValuePair<int,string>(0xD3A692, "CDD - Container Data Descriptor"),
            new KeyValuePair<int,string>(0xD3A78A, "CFC - Coded Font Control"),
            new KeyValuePair<int,string>(0xD38C8A, "CFI - Coded Font Index"),
            new KeyValuePair<int,string>(0xD3A787, "CPC - Code Page Control"),
            new KeyValuePair<int,string>(0xD3A687, "CPD - Code Page Descriptor"),
            new KeyValuePair<int,string>(0xD38C87, "CPI - Code Page Index"),
            new KeyValuePair<int,string>(0xD3A79B, "CTC - Composed Text Control"),
            new KeyValuePair<int,string>(0xD3A9C9, "EAG - End Active Environment Group"),
            new KeyValuePair<int,string>(0xD3A9EB, "EBC - End Bar Code Object"),
            new KeyValuePair<int,string>(0xD3A98A, "ECF - End Coded Font"),
            new KeyValuePair<int,string>(0xD3A987, "ECP - End Code Page"),
            new KeyValuePair<int,string>(0xD3A9C4, "EDG - End Document Environment Group"),
            new KeyValuePair<int,string>(0xD3A9A7, "EDI - End Document Index"),
            new KeyValuePair<int,string>(0xD3A9A8, "EDT - End Document"),
            new KeyValuePair<int,string>(0xD3A9C5, "EFG - End Form Environment Group"),
            new KeyValuePair<int,string>(0xD3A9CD, "EFM - End Form Map"),
            new KeyValuePair<int,string>(0xD3A989, "EFN - End Font"),
            new KeyValuePair<int,string>(0xD3A9BB, "EGR - End Graphics Object"),
            new KeyValuePair<int,string>(0xD3A97B, "EII - End IM Image"),
            new KeyValuePair<int,string>(0xD3A9FB, "EIM - End Image Object"),
            new KeyValuePair<int,string>(0xD3A9CC, "EMM - End Medium Map"),
            new KeyValuePair<int,string>(0xD3A9DF, "EMO - End Overlay"),
            new KeyValuePair<int,string>(0xD3A9AD, "ENG - End Named Page Group"),
            new KeyValuePair<int,string>(0xD3A992, "EOC - End Object Container"),
            new KeyValuePair<int,string>(0xD3A9C7, "EOG - End Object Environment Group"),
            new KeyValuePair<int,string>(0xD3A9A5, "EPF - End Print File"),
            new KeyValuePair<int,string>(0xD3A9AF, "EPG - End Page"),
            new KeyValuePair<int,string>(0xD3A95F, "EPS - End Page Segment"),
            new KeyValuePair<int,string>(0xD3A99B, "EPT - End Presentation Text Object"),
            new KeyValuePair<int,string>(0xD3A9C6, "ERG - End Resource Group"),
            new KeyValuePair<int,string>(0xD3A9CE, "ERS - End Resource"),
            new KeyValuePair<int,string>(0xD3A9D9, "ESG - End Resource Environment Group"),
            new KeyValuePair<int,string>(0xD3A6C5, "FGD - Form Environment Group Descriptor"),
            new KeyValuePair<int,string>(0xD3AE89, "FNO - Font 0rientation"),
            new KeyValuePair<int,string>(0xD3A789, "FNC - Font Control"),
            new KeyValuePair<int,string>(0xD3A689, "FND - Font Descriptor"),
            new KeyValuePair<int,string>(0xD3EE89, "FNG - Font Patterns"),
            new KeyValuePair<int,string>(0xD38C89, "FNI - Font Index"),
            new KeyValuePair<int,string>(0xD3A289, "FNM - Font Patterns Map"),
            new KeyValuePair<int,string>(0xD3AB89, "FNN - Font Names (Outline Fonts Only)"),
            new KeyValuePair<int,string>(0xD3AC89, "FNP - Font Position"),
            new KeyValuePair<int,string>(0xD3EEBB, "GAD - Graphics Data"),
            new KeyValuePair<int,string>(0xD3A6BB, "GDD - Graphics Data Descriptor"),
            new KeyValuePair<int,string>(0xD3AC7B, "ICP - IM Image Cell Position "),
            new KeyValuePair<int,string>(0xD3A6FB, "IDD - Image Data Descriptor"),
            new KeyValuePair<int,string>(0xD3B2A7, "IEL - Index Element"),
            new KeyValuePair<int,string>(0xD3A67B, "IID - Image Input Descriptor "),
            new KeyValuePair<int,string>(0xD3ABCC, "IMM - Invoke Medium Map"),
            new KeyValuePair<int,string>(0xD3AFC3, "IOB - Include Object"),
            new KeyValuePair<int,string>(0xD3A77B, "IOC - IM Image Output Control "),
            new KeyValuePair<int,string>(0xD3EEFB, "IPD - Image Picture Data"),
            new KeyValuePair<int,string>(0xD3AFAF, "IPG - Include Page"),
            new KeyValuePair<int,string>(0xD3AFD8, "IPO - Include Page Overlay"),
            new KeyValuePair<int,string>(0xD3AF5F, "IPS - Include Page Segment"),
            new KeyValuePair<int,string>(0xD3EE7B, "IRD - IM Image Raster Data "),
            new KeyValuePair<int,string>(0xD3B490, "LLE - Link Logical Element"),
            new KeyValuePair<int,string>(0xD3ABEB, "MBC - Map Bar Code Object"),
            new KeyValuePair<int,string>(0xD3A288, "MCC - Medium Copy Count"),
            new KeyValuePair<int,string>(0xD3AB92, "MCD - Map Container Data"),
            new KeyValuePair<int,string>(0xD3AB8A, "MCF - Map Coded Font"),
            new KeyValuePair<int,string>(0xD3B18A, "MCF-1 - Map Coded Font Format-1 "),
            new KeyValuePair<int,string>(0xD3A688, "MDD - Medium Descriptor"),
            new KeyValuePair<int,string>(0xD3ABC3, "MDR - Map Data Resource"),
            new KeyValuePair<int,string>(0xD3A088, "MFC - Medium Finishing Control"),
            new KeyValuePair<int,string>(0xD3ABBB, "MGO - Map Graphics Object"),
            new KeyValuePair<int,string>(0xD3ABFB, "MIO - Map Image Object"),
            new KeyValuePair<int,string>(0xD3A788, "MMC - Medium Modification Control"),
            new KeyValuePair<int,string>(0xD3ABCD, "MMD - Map Media Destination"),
            new KeyValuePair<int,string>(0xD3B1DF, "MMO - Map Medium Overlay"),
            new KeyValuePair<int,string>(0xD3AB88, "MMT - Map Media Type"),
            new KeyValuePair<int,string>(0xD3ABAF, "MPG - Map Page"),
            new KeyValuePair<int,string>(0xD3ABD8, "MPO - Map Page Overlay"),
            new KeyValuePair<int,string>(0xD3B15F, "MPS - Map Page Segment"),
            new KeyValuePair<int,string>(0xD3AB9B, "MPT - Map Presentation Text"),
            new KeyValuePair<int,string>(0xD3ABEA, "MSU - Map Suppression"),
            new KeyValuePair<int,string>(0xD3EEEE, "NOP - No Operation"),
            new KeyValuePair<int,string>(0xD3A66B, "OBD - Object Area Descriptor"),
            new KeyValuePair<int,string>(0xD3AC6B, "OBP - Object Area Position"),
            new KeyValuePair<int,string>(0xD3EE92, "OCD - Object Container Data"),
            new KeyValuePair<int,string>(0xD3A7A8, "PEC - Presentation Environment Control"),
            new KeyValuePair<int,string>(0xD3B288, "PFC - Presentation Fidelity Control"),
            new KeyValuePair<int,string>(0xD3A6AF, "PGD - Page Descriptor"),
            new KeyValuePair<int,string>(0xD3B1AF, "PGP - Page Position"),
            new KeyValuePair<int,string>(0xD3ACAF, "PGP-1 - Page Position Format-1 "),
            new KeyValuePair<int,string>(0xD3A7AF, "PMC - Page Modification Control"),
            new KeyValuePair<int,string>(0xD3ADC3, "PPO - Preprocess Presentation Object"),
            new KeyValuePair<int,string>(0xD3B19B, "PTD - Presentation Text Data Descriptor"),
            new KeyValuePair<int,string>(0xD3A69B, "PTD-1 - Presentation Text Descriptor Format-1 "),
            new KeyValuePair<int,string>(0xD3EE9B, "PTX - Presentation Text Data"),
            new KeyValuePair<int,string>(0xD3A090, "TLE - Tag Logical Element"),
        };

        public static readonly Dictionary<string, SfInfo> HexToSf = new Dictionary<string, SfInfo>();

        static AfpConstants()
        {
            foreach (var kvp in _sfNamesArray)
            {
                string hex = kvp.Key.ToString("X6");
                string acronym = "UNK";
                string friendlyName = kvp.Value;
                
                int dashIndex = kvp.Value.IndexOf(" - ");
                if (dashIndex > 0)
                {
                    acronym = kvp.Value.Substring(0, dashIndex).Trim();
                    friendlyName = kvp.Value.Substring(dashIndex + 3).Trim();
                }
                
                HexToSf[hex] = new SfInfo { Acronym = acronym, FriendlyName = friendlyName };
            }
        }
        
        public static SfInfo GetSfInfo(string hexId)
        {
            if (HexToSf.TryGetValue(hexId.ToUpper(), out SfInfo sf))
                return sf;
            
            return new SfInfo { Acronym = "UNKNOWN", FriendlyName = "Unknown Structured Field" };
        }
    }
}