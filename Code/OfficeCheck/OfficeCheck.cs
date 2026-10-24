//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using SystemBannerOffice;
using SystemBannerOffice.Markings;

namespace SystemBannerOffice.Check
{
    //drives the installed Word, Excel and PowerPoint through automation with the document code of the add-in,
    //without installing the add-in: OfficeCheck.exe [word] [excel] [powerpoint]; OfficeCheck.exe diag reports the installation
    public static class OfficeCheck
    {
        private static int failures;

        private static void Check(bool condition, string what)
        {
            Console.WriteLine((condition ? "ok   " : "FAIL ") + what);
            if (!condition) failures++;
        }

        [STAThread] //a COM message filter only works on a single threaded apartment
        public static int Main(string[] args)
        {
            MessageFilter.Register();
            if (args.Length > 0 && args[0] == "diag")
            {
                return Diag.Run();
            }
            if (args.Length > 0 && args[0] == "load")
            {
                //OfficeCheck.exe load [word|excel|powerpoint]: Office itself loads the installed add-in, as the COM add-ins dialog does
                Run((args.Length > 1 ? args[1] : "word") == "excel" ? "Excel.Application" : (args.Length > 1 && args[1] == "powerpoint" ? "PowerPoint.Application" : "Word.Application"), LoadAddIn);
                Console.WriteLine(failures == 0 ? "add-in loaded" : "add-in did not load");
                return failures == 0 ? 0 : 1;
            }
            if (args.Length > 0 && args[0] == "toolbar")
            {
                //OfficeCheck.exe toolbar: starts the add-in in each Office application as Office does, with the toolbar
                Connect.ToolbarAlways = true;
                Run("Word.Application", Toolbar);
                Run("Excel.Application", Toolbar);
                Run("PowerPoint.Application", Toolbar);
                Console.WriteLine(failures == 0 ? "all checks passed" : failures + " check(s) failed");
                return failures == 0 ? 0 : 1;
            }
            List<string> apps = new List<string>(args.Length > 0 ? args : new string[] { "word", "excel", "powerpoint" });
            if (apps.Contains("word")) Run("Word.Application", Word);
            if (apps.Contains("excel")) Run("Excel.Application", Excel);
            if (apps.Contains("powerpoint")) Run("PowerPoint.Application", PowerPoint);
            Console.WriteLine(failures == 0 ? "all checks passed" : failures + " check(s) failed");
            return failures == 0 ? 0 : 1;
        }

        private delegate void Scenario(object application);

        private static object FindToolbar(object application)
        {
            try
            {
                return Late.Item(Late.Get(application, "CommandBars"), Commands.Title);
            }
            catch (Exception)
            {
                return null;
            }
        }

        //the toolbar of Office 2000 to 2003: created at startup, kept visible by the later checks, removed at the end
        private static void Toolbar(object application)
        {
            if (Convert.ToString(Late.Get(application, "Name")).IndexOf("PowerPoint") >= 0)
            {
                Late.TrySet(application, "Visible", -1); //PowerPoint shows no command bar while hidden
            }
            //PowerPoint runs one instance: the add-in installed on this computer may already be loaded in it, with its own toolbar
            bool installedLoaded = FindToolbar(application) != null;
            if (installedLoaded)
            {
                Console.WriteLine("note the installed add-in is loaded in this instance: its toolbar stays after the test");
            }
            Connect addIn = new Connect();
            Array custom = new object[0];
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_Startup, null, ref custom);
            addIn.OnStartupComplete(ref custom);
            object bar = FindToolbar(application);
            Check(bar != null, "toolbar created");
            if (bar == null) return;
            Check(Late.Count(Late.Get(bar, "Controls")) == 5, "toolbar has its 5 buttons");
            Check(Convert.ToBoolean(Late.Get(bar, "Visible")), "toolbar visible");
            //Office hides it, as Word and Excel 2000 can do when they restore their own layout: the later check shows it again
            Late.Set(bar, "Visible", false);
            DateTime end = DateTime.Now.AddSeconds(3);
            while (DateTime.Now < end)
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(50);
            }
            Check(Convert.ToBoolean(Late.Get(bar, "Visible")), "hidden toolbar shown again by the later check");
            //Office deletes it: the later check creates it again
            Late.Call(bar, "Delete");
            end = DateTime.Now.AddSeconds(3);
            while (DateTime.Now < end)
            {
                System.Windows.Forms.Application.DoEvents();
                System.Threading.Thread.Sleep(50);
            }
            bar = FindToolbar(application);
            Check(bar != null && Late.Count(Late.Get(bar, "Controls")) == 5, "deleted toolbar created again by the later check");
            addIn.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, ref custom);
            if (!installedLoaded)
            {
                Check(FindToolbar(application) == null, "toolbar removed when the add-in unloads");
            }
        }

        private static void Run(string progId, Scenario scenario)
        {
            Type type = Type.GetTypeFromProgID(progId);
            if (type == null)
            {
                Console.WriteLine("skip " + progId + ": not installed");
                return;
            }
            object application = Activator.CreateInstance(type);
            try
            {
                Console.WriteLine(progId + " " + Late.Get(application, "Version"));
                Late.TrySet(application, "DisplayAlerts", progId.StartsWith("Word") ? (object)0 : (object)false);
                scenario(application);
            }
            catch (Exception ex)
            {
                Exception inner = ex.InnerException ?? ex;
                Console.WriteLine("FAIL " + progId + ": " + inner.Message + Environment.NewLine + ex);
                failures++;
            }
            finally
            {
                try { Late.Call(application, "Quit"); } catch (Exception) { }
                Marshal.ReleaseComObject(application);
                //the late bound calls leave runtime callable wrappers behind: collect them so the automation instance ends
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        //records the events Word raises, to check the save check is wired to the right DISPIDs
        [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
        public class WordRecorder : WordApplicationEvents
        {
            public int Saves, Prints;
            public void DocumentBeforePrint(object document, ref bool cancel) { Prints++; }
            public void DocumentBeforeSave(object document, ref bool saveAsUi, ref bool cancel) { Saves++; cancel = true; }
        }

        [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
        public class ExcelRecorder : ExcelApplicationEvents
        {
            public int Saves;
            public void WorkbookBeforeSave(object workbook, bool saveAsUi, ref bool cancel) { Saves++; cancel = true; }
            public void WorkbookBeforePrint(object workbook, ref bool cancel) { }
        }

        private static void LoadAddIn(object application)
        {
            object addIns = Late.Get(application, "COMAddIns");
            object addIn;
            try
            {
                addIn = Late.Item(addIns, "SystemBannerOffice.Connect");
            }
            catch (Exception)
            {
                Check(false, "the add-in is registered in Office (COMAddIns)");
                return;
            }
            Console.WriteLine("     LoadBehavior seen by Office, connected: " + Late.Get(addIn, "Connect"));
            if (!Convert.ToBoolean(Late.Get(addIn, "Connect")))
            {
                Late.Set(addIn, "Connect", true);
            }
            Check(Convert.ToBoolean(Late.Get(addIn, "Connect")), "the add-in is connected");
            Console.WriteLine("     log: " + Log.FilePath);
            if (System.IO.File.Exists(Log.FilePath))
            {
                string[] lines = System.IO.File.ReadAllLines(Log.FilePath);
                for (int i = Math.Max(0, lines.Length - 12); i < lines.Length; i++)
                {
                    Console.WriteLine("     " + lines[i]);
                }
            }
            Late.Set(addIn, "Connect", false);
        }

        private static void Word(object word)
        {
            object document = Late.Call(Late.Get(word, "Documents"), "Add");
            object content = Late.Get(document, "Content");
            Late.Set(content, "Text", "(S//REL TO USA, GBR) First paragraph.\r(U) Second paragraph.\r(TS//SI//NF) Third paragraph.\rA paragraph with no mark.");
            //an existing header line the banner must not erase
            object primary = Late.Item(Late.Get(Late.Item(Late.Get(document, "Sections"), 1), "Headers"), 1);
            Late.Set(Late.Get(primary, "Range"), "Text", "Company header");
            DocumentHost host = DocumentHost.Active(word);
            Check(host is WordHost, "Word host");
            PortionScan scan = host.ScanPortions();
            Check(scan.Marks.Count == 3 && scan.Unmarked == 1, "Word scan: 3 marked, 1 unmarked (" + scan.Marks.Count + ", " + scan.Unmarked + ")");
            RollUpResult result = RollUp.Compute(scan.Marks, null);
            string banner = MarkingFormatter.Banner(result.Banner);
            Check(banner == "TOP SECRET//SI//NOFORN", "Word roll-up: " + banner);
            host.Mark(result.Banner);
            string header = Convert.ToString(Late.Get(Late.Get(primary, "Range"), "Text"));
            Check(header.StartsWith("TOP SECRET//SI//NOFORN\r") && header.Contains("Company header"), "Word header keeps its content: " + header.Replace("\r", " | "));
            object footer = Late.Item(Late.Get(Late.Item(Late.Get(document, "Sections"), 1), "Footers"), 1);
            Check(Convert.ToString(Late.Get(Late.Get(footer, "Range"), "Text")).Trim() == "TOP SECRET//SI//NOFORN", "Word footer");
            //marking again replaces the banner, it does not add a second one
            Marking secret = MarkingParser.Parse("SECRET//NOFORN").Marking;
            host.Mark(secret);
            header = Convert.ToString(Late.Get(Late.Get(primary, "Range"), "Text"));
            Check(header.StartsWith("SECRET//NOFORN\r") && header.IndexOf("TOP SECRET") < 0 && header.Contains("Company header"), "Word banner replaced: " + header.Replace("\r", " | "));
            Check(host.GetProperty(DocumentHost.BannerProperty) == "SECRET//NOFORN" && host.StoredClassification() == "SECRET", "Word document properties");
            //portion mark at the cursor: replaces "(U) " on the second paragraph
            object second = Late.Get(Late.Item(Late.Get(document, "Paragraphs"), 2), "Range");
            Late.Call(Late.Get(word, "Selection"), "SetRange", Convert.ToInt32(Late.Get(second, "Start")) + 6, Convert.ToInt32(Late.Get(second, "Start")) + 6);
            host.InsertPortionMark("(C//NF)");
            string secondText = Convert.ToString(Late.Get(Late.Get(Late.Item(Late.Get(document, "Paragraphs"), 2), "Range"), "Text"));
            Check(secondText.StartsWith("(C//NF) Second paragraph."), "Word portion mark replaced: " + secondText.Trim());
            //authority block, written once and replaced
            AuthorityBlock block = new AuthorityBlock();
            block.ClassifiedBy = "J. Doe";
            block.DerivedFrom = "Guide X";
            block.DeclassifyOn = "20450101";
            Late.Call(Late.Get(word, "Selection"), "EndKey", 6); //wdStory
            host.WriteBlock(DocumentHost.AuthorityBookmark, block.Lines());
            block.DeclassifyOn = "20460101";
            host.WriteBlock(DocumentHost.AuthorityBookmark, block.Lines());
            string all = Convert.ToString(Late.Get(Late.Get(document, "Content"), "Text"));
            Check(all.Contains("Declassify On: 20460101") && all.IndexOf("20450101") < 0, "Word authority block replaced");
            host.StoreAuthorityBlock(block);
            Check(host.StoredAuthorityBlock().DeclassifyOn == "20460101", "Word authority block stored");
            //save event: the recorder cancels the save
            WordRecorder recorder = new WordRecorder();
            using (new ComEventLink(word, typeof(WordApplicationEvents).GUID, recorder))
            {
                string path = Path.Combine(Path.GetTempPath(), "SystemBannerOfficeCheck.docx");
                File.Delete(path);
                try { Late.Call(document, "SaveAs", path); } catch (Exception) { }
                Check(recorder.Saves == 1 && !File.Exists(path), "Word DocumentBeforeSave event (DISPID 8) cancels the save");
            }
            Late.Call(document, "Close", 0); //wdDoNotSaveChanges
            WordFrench(word);
            WordGerman(word);
            WordBritish(word);
        }

        //VS-Anweisung: part marks, GEHEIM in red at the top and bottom with "amtlich geheim gehalten", VS-NfD at the top only, Seite X von Y
        private static void WordGerman(object word)
        {
            object document = Late.Call(Late.Get(word, "Documents"), "Add");
            Late.Set(Late.Get(document, "Content"), "Text", "[offen] Betreff.\r[Geh.] Geheimer Teil.\r[VS-NfD] Teil VS-NfD.");
            DocumentHost host = DocumentHost.Active(word);
            PortionScan scan = host.ScanPortions();
            Check(scan.GermanMarks.Count == 3 && scan.FrenchMarks.Count == 0 && scan.Systems == 1, "Word German scan: 3 part marks (" + scan.GermanMarks.Count + ")");
            GermanMarking marking = GermanRollUp.Compute(scan.GermanMarks);
            Check(marking.Level == GermanLevel.Geheim, "Word German roll-up: " + marking);
            marking.Vermerke.Add("Nur Deutschen zur Kenntnis");
            host.Mark(marking);
            object section = Late.Item(Late.Get(document, "Sections"), 1);
            object header = Late.Get(Late.Item(Late.Get(section, "Headers"), 1), "Range");
            object footer = Late.Get(Late.Item(Late.Get(section, "Footers"), 1), "Range");
            Check(Text(header).StartsWith("GEHEIM\ramtlich geheim gehalten\rNur Deutschen zur Kenntnis"), "Word German header: " + Show(Text(header)));
            Check(Text(footer).StartsWith("GEHEIM\ramtlich geheim gehalten\r") && Text(footer).Contains("Seite 1 von 1") && Text(footer).IndexOf("Nur Deutschen") < 0, "Word German footer: " + Show(Text(footer)));
            object first = Late.Get(Late.Get(header, "Characters"), "First");
            Check(Convert.ToInt32(Late.Get(Late.Get(first, "Font"), "Color")) == StampSet.Red, "Word GEHEIM in red");
            Check(host.StoredClassification() == "GEHEIM" && host.StoredSystem() == DocumentHost.GermanSystem && host.StoredGermanMarking().Vermerke.Count == 1, "Word German document properties");
            host.Mark(GermanParser.ParsePortion("[VS-NfD]"));
            Check(Text(header).StartsWith("VS-NUR FÜR DEN DIENSTGEBRAUCH\r") && Text(header).IndexOf("amtlich") < 0, "Word VS-NfD header, the level alone: " + Show(Text(header)));
            Check(Text(footer).IndexOf("VS-NUR") < 0 && Text(footer).Contains("Seite 1 von 1"), "Word VS-NfD: no bottom stamp, page numbers: " + Show(Text(footer)));
            GermanIdentification id = new GermanIdentification();
            id.Herausgeber = "Bundesministerium X";
            Late.Call(Late.Get(word, "Selection"), "HomeKey", 6);
            host.WriteBlock(DocumentHost.IdentificationBookmark, id.Lines(GermanParser.ParsePortion("[VS-NfD]")));
            Check(Text(Late.Get(document, "Content")).Contains("Herausgeber: Bundesministerium X") && Text(Late.Get(document, "Content")).Contains("1 Seiten"), "Word German first page with the page count field");
            Late.Call(document, "Close", 0);
        }

        //GSCP: the classification and its additional markings at the top and bottom, black, no frame, no page numbers
        private static void WordBritish(object word)
        {
            object document = Late.Call(Late.Get(word, "Documents"), "Add");
            DocumentHost host = DocumentHost.Active(word);
            UkMarking marking = new UkMarking();
            marking.Sensitive = true;
            marking.HandlingInstruction = UkRegister.Embargoed;
            marking.Embargo = "until 10:00 1 November 2026, then OFFICIAL";
            marking.Descriptors.Add("PERSONAL DATA");
            host.Mark(marking);
            object section = Late.Item(Late.Get(document, "Sections"), 1);
            object header = Late.Get(Late.Item(Late.Get(section, "Headers"), 1), "Range");
            object footer = Late.Get(Late.Item(Late.Get(section, "Footers"), 1), "Range");
            Check(Text(header).StartsWith("OFFICIAL-SENSITIVE EMBARGOED PERSONAL DATA\rEMBARGOED: until 10:00 1 November 2026, then OFFICIAL"), "Word British header: " + Show(Text(header)));
            Check(Text(footer).StartsWith("OFFICIAL-SENSITIVE EMBARGOED PERSONAL DATA") && Text(footer).IndexOf("page") < 0, "Word British footer: " + Show(Text(footer)));
            object first = Late.Get(Late.Get(header, "Characters"), "First");
            Check(Convert.ToInt32(Late.Get(Late.Get(first, "Font"), "Color")) == 0, "Word British marking in black, no colour coding");
            Check(host.StoredClassification() == "OFFICIAL-SENSITIVE" && host.StoredUkMarking().Descriptors.Count == 1, "Word British document properties");
            host.Mark(FrenchParser.ParsePortion("[S]"), new FrenchIdentification(), false);
            Check(Text(header).StartsWith("SECRET\r") && Text(header).IndexOf("OFFICIAL") < 0 && Text(header).IndexOf("EMBARGOED") < 0, "Word French stamps replace the British marking: " + Show(Text(header)));
            Late.Call(document, "Close", 0);
        }

        private static void Excel(object excel)
        {
            object workbook = Late.Call(Late.Get(excel, "Workbooks"), "Add");
            Late.Call(Late.Get(workbook, "Worksheets"), "Add");
            DocumentHost host = DocumentHost.Active(excel);
            Check(host is ExcelHost, "Excel host");
            host.Mark(MarkingParser.Parse("SECRET//REL TO USA, FVEY").Marking);
            object sheets = Late.Get(workbook, "Worksheets");
            for (int i = 1; i <= Late.Count(sheets); i++)
            {
                object setup = Late.Get(Late.Item(sheets, i), "PageSetup");
                string header = Convert.ToString(Late.Get(setup, "CenterHeader"));
                Check(header.EndsWith("SECRET//REL TO USA, FVEY") && header.Contains("&KC8102E"), "Excel sheet " + i + " header: " + header);
                Check(Convert.ToString(Late.Get(setup, "CenterFooter")) == header, "Excel sheet " + i + " footer");
            }
            Check(host.StoredClassification() == "SECRET", "Excel document properties");
            ExcelRecorder recorder = new ExcelRecorder();
            using (new ComEventLink(excel, typeof(ExcelApplicationEvents).GUID, recorder))
            {
                string path = Path.Combine(Path.GetTempPath(), "SystemBannerOfficeCheck.xlsx");
                File.Delete(path);
                try { Late.Call(workbook, "SaveAs", path); } catch (Exception) { }
                Console.WriteLine("     Excel saves seen: " + recorder.Saves + ", file written: " + File.Exists(path));
                Check(recorder.Saves == 1 && !File.Exists(path), "Excel WorkbookBeforeSave event (DISPID 0x623) cancels the save");
            }
            Late.Call(workbook, "Close", false);
            ExcelFrench(excel);
        }

        private static void PowerPoint(object powerPoint)
        {
            object presentation = Late.Call(Late.Get(powerPoint, "Presentations"), "Add", 0); //no window
            object slide = Late.Call(Late.Get(presentation, "Slides"), "Add", 1, 2); //ppLayoutText
            object shapes = Late.Get(slide, "Shapes");
            Late.Set(Late.Get(Late.Get(Late.Item(shapes, 1), "TextFrame"), "TextRange"), "Text", "(S) Operational Plan");
            Late.Set(Late.Get(Late.Get(Late.Item(shapes, 2), "TextFrame"), "TextRange"), "Text", "(U) Only UNCLASSIFIED information.\r(S//NF) SECRET information.");
            DocumentHost host = DocumentHost.For(powerPoint, presentation);
            Check(host is PowerPointHost, "PowerPoint host");
            PortionScan scan = host.ScanPortions();
            Check(scan.Marks.Count == 3, "PowerPoint scan: 3 portion marks (" + scan.Marks.Count + ")");
            RollUpResult result = RollUp.Compute(scan.Marks, null);
            Check(MarkingFormatter.Banner(result.Banner) == "SECRET//NOFORN", "PowerPoint roll-up: " + MarkingFormatter.Banner(result.Banner));
            host.Mark(result.Banner);
            host.Mark(result.Banner); //twice: still one banner of each
            object masterShapes = Late.Get(Late.Get(presentation, "SlideMaster"), "Shapes");
            int top = 0, bottom = 0;
            for (int i = 1; i <= Late.Count(masterShapes); i++)
            {
                object shape = Late.Item(masterShapes, i);
                string name = Convert.ToString(Late.Get(shape, "Name"));
                if (name == "SystemBannerTop") top++;
                if (name == "SystemBannerBottom")
                {
                    bottom++;
                    Check(Convert.ToString(Late.Get(Late.Get(Late.Get(shape, "TextFrame"), "TextRange"), "Text")) == "SECRET//NOFORN", "PowerPoint bottom banner text");
                }
            }
            Check(top == 1 && bottom == 1, "PowerPoint master has one top and one bottom banner (" + top + ", " + bottom + ")");
            Late.Call(presentation, "Close");
            PowerPointFrench(powerPoint);
        }

        private static FrenchIdentification Identification()
        {
            FrenchIdentification id = new FrenchIdentification();
            id.AutoriteEmettrice = "Ministère X";
            id.Auteur = "Bureau Y";
            id.NumeroEnregistrement = "2026-042";
            id.Echeance = DateTime.Today.AddYears(10);
            id.Exemplaire = 1;
            id.NombreExemplaires = 3;
            id.Reference = "TS-2026-042";
            return id;
        }

        private static string Text(object range)
        {
            return Convert.ToString(Late.Get(range, "Text"));
        }

        private static string Show(string text)
        {
            return text.Replace("\r", " | ");
        }

        //IGI 1300: paragraph marks, roll-up, red framed stamps, Spécial France at the top only, page numbers, identification, cover
        private static void WordFrench(object word)
        {
            object document = Late.Call(Late.Get(word, "Documents"), "Add");
            Late.Set(Late.Get(document, "Content"), "Text", "[NP] Objet du document.\r[S SF] Paragraphe classifié Secret, Spécial France.\r[DR] Paragraphe Diffusion Restreinte.\rParagraphe sans marque.");
            DocumentHost host = DocumentHost.Active(word);
            PortionScan scan = host.ScanPortions();
            Check(scan.FrenchMarks.Count == 3 && scan.Unmarked == 1 && scan.Marks.Count == 0, "Word French scan: 3 marks, 1 unmarked (" + scan.FrenchMarks.Count + ", " + scan.Unmarked + ")");
            FrenchMarking marking = FrenchRollUp.Compute(scan.FrenchMarks).Document;
            Check(marking.ToString() == "SECRET SPÉCIAL FRANCE", "Word French roll-up: " + marking);
            FrenchIdentification id = Identification();
            host.StoreFrenchIdentification(id);
            host.Mark(marking, id, false);
            object section = Late.Item(Late.Get(document, "Sections"), 1);
            object header = Late.Get(Late.Item(Late.Get(section, "Headers"), 1), "Range");
            object footer = Late.Get(Late.Item(Late.Get(section, "Footers"), 1), "Range");
            Check(Text(header).StartsWith("SECRET\rSPÉCIAL FRANCE"), "Word French header: " + Show(Text(header)));
            object first = Late.Get(Late.Get(header, "Characters"), "First");
            Check(Convert.ToInt32(Late.Get(Late.Get(first, "Font"), "Color")) == StampSet.Red, "Word stamp in red");
            Check(Convert.ToInt32(Late.Get(Late.Get(Late.Get(first, "Font"), "Borders"), "OutsideLineStyle")) == 1, "Word stamp framed");
            object bodyStart = Late.Get(Late.Get(Late.Get(document, "Content"), "Characters"), "First");
            Check(Convert.ToInt32(Late.Get(Late.Get(bodyStart, "Font"), "Color")) != StampSet.Red, "Word body text keeps its formatting");
            string footerText = Text(footer);
            Check(footerText.StartsWith("SECRET\r") && footerText.Contains("page 1 / 1") && footerText.IndexOf("SPÉCIAL") < 0, "Word French footer, Spécial France at the top only: " + Show(footerText));
            Check(host.StoredClassification() == "SECRET SPÉCIAL FRANCE" && host.StoredSystem() == DocumentHost.FrenchSystem, "Word French document properties");
            Late.Call(Late.Get(word, "Selection"), "HomeKey", 6); //wdStory
            host.WriteBlock(DocumentHost.IdentificationBookmark, id.Lines(marking));
            string body = Text(Late.Get(document, "Content"));
            Check(body.Contains("N° 2026-042") && body.Contains("Déclassifié le ") && body.Contains("Ce document comporte 1 pages"), "Word identification block with the page count field: " + Show(body.Substring(0, Math.Min(200, body.Length))));
            Check(host.ScanPortions().FrenchMarks.Count == 3 && body.Contains("pages\r[NP] Objet"), "Word identification block is its own paragraph, the [NP] mark stays at the start of its paragraph");
            host.WriteBlock(DocumentHost.IdentificationBookmark, id.Lines(marking));
            Check(Text(Late.Get(document, "Content")).Split(new string[] { "N° 2026-042" }, StringSplitOptions.None).Length == 2, "Word identification block written again in place, not twice");
            //Diffusion Restreinte: top only, no page numbers
            host.Mark(FrenchParser.ParsePortion("[DR SF]"), id, false);
            Check(Text(header).StartsWith("DIFFUSION RESTREINTE\rSPÉCIAL FRANCE"), "Word Diffusion Restreinte header: " + Show(Text(header)));
            Check(Text(footer).Trim('\r', ' ').Length == 0, "Word Diffusion Restreinte footer is empty: " + Show(Text(footer)));
            //TRÈS SECRET: the reference on every page
            host.Mark(FrenchParser.ParsePortion("[TS]"), id, false);
            Check(Text(footer).Contains("TS-2026-042 - page 1 / 1") && Text(header).IndexOf("SPÉCIAL") < 0, "Word Très Secret footer with the reference: " + Show(Text(footer)));
            //bound document: the cover stamp at the bottom of the first page
            host.Mark(marking, id, true);
            object firstFooter = Late.Get(Late.Item(Late.Get(section, "Footers"), 2), "Range");
            Check(Text(firstFooter).Contains(FrenchFormatter.CoverWarning) && Text(footer).StartsWith("SECRET\r"), "Word cover stamp on the first page: " + Show(Text(firstFooter)));
            //back to an IC banner: the frame goes, the background comes back
            host.Mark(MarkingParser.Parse("SECRET//NOFORN").Marking);
            first = Late.Get(Late.Get(header, "Characters"), "First");
            Check(Text(header).StartsWith("SECRET//NOFORN\r") && Text(header).IndexOf("SPÉCIAL") < 0 && Text(footer).IndexOf("page") < 0, "Word IC banner replaces the French stamps: " + Show(Text(header)) + " / " + Show(Text(footer)));
            Check(Convert.ToInt32(Late.Get(Late.Get(Late.Get(first, "Font"), "Borders"), "OutsideLineStyle")) == 0, "Word IC banner has no frame");
            Late.Call(document, "Close", 0);
        }

        private static void ExcelFrench(object excel)
        {
            object workbook = Late.Call(Late.Get(excel, "Workbooks"), "Add");
            DocumentHost host = DocumentHost.Active(excel);
            host.Mark(FrenchParser.ParsePortion("[TS SF]"), Identification(), false);
            object setup = Late.Get(Late.Item(Late.Get(workbook, "Worksheets"), 1), "PageSetup");
            string header = Convert.ToString(Late.Get(setup, "CenterHeader"));
            string footer = Convert.ToString(Late.Get(setup, "CenterFooter"));
            Check(header.Contains("&KC8102ETRÈS SECRET") && header.Contains("&K0033A0SPÉCIAL FRANCE"), "Excel French header: " + header);
            Check(footer.Contains("TRÈS SECRET") && footer.Contains("TS-2026-042 - page &P / &N") && footer.IndexOf("SPÉCIAL") < 0, "Excel French footer: " + footer);
            host.Mark(FrenchParser.ParsePortion("[DR]"), Identification(), false);
            Check(Convert.ToString(Late.Get(setup, "CenterFooter")).Length == 0, "Excel Diffusion Restreinte: no footer");
            host.Mark(GermanParser.ParsePortion("[Str. Geh.]"));
            Check(Convert.ToString(Late.Get(setup, "CenterHeader")).Contains("STRENG GEHEIM") && Convert.ToString(Late.Get(setup, "CenterFooter")).Contains("Seite &P von &N"), "Excel German: " + Convert.ToString(Late.Get(setup, "CenterFooter")));
            UkMarking uk = new UkMarking();
            uk.Level = UkLevel.Secret;
            uk.Caveats.Add("UK EYES ONLY");
            host.Mark(uk);
            Check(Convert.ToString(Late.Get(setup, "CenterHeader")).Contains("&K000000SECRET UK EYES ONLY") && Convert.ToString(Late.Get(setup, "CenterFooter")).IndexOf("Seite") < 0, "Excel British: " + Convert.ToString(Late.Get(setup, "CenterHeader")));
            Late.Call(workbook, "Close", false);
        }

        private static void PowerPointFrench(object powerPoint)
        {
            object presentation = Late.Call(Late.Get(powerPoint, "Presentations"), "Add", 0);
            DocumentHost host = DocumentHost.For(powerPoint, presentation);
            host.Mark(FrenchParser.ParsePortion("[S SF]"), Identification(), false);
            host.Mark(FrenchParser.ParsePortion("[S SF]"), Identification(), false); //twice: still one stamp of each
            object shapes = Late.Get(Late.Get(presentation, "SlideMaster"), "Shapes");
            List<string> names = new List<string>();
            float weight = 0;
            for (int i = 1; i <= Late.Count(shapes); i++)
            {
                object shape = Late.Item(shapes, i);
                string name = Convert.ToString(Late.Get(shape, "Name"));
                if (!name.StartsWith("SystemBanner")) continue;
                names.Add(name + "=" + Text(Late.Get(Late.Get(shape, "TextFrame"), "TextRange")));
                if (name == "SystemBannerTop") weight = Convert.ToSingle(Late.Get(Late.Get(shape, "Line"), "Weight"));
            }
            names.Sort(string.CompareOrdinal);
            string all = string.Join(", ", names.ToArray());
            Check(all == "SystemBannerBottom=SECRET, SystemBannerTop2=SPÉCIAL FRANCE, SystemBannerTop=SECRET", "PowerPoint French stamps: " + all);
            Check(weight == 2.5f, "PowerPoint stamp frame 2.5 pt (" + weight + ")");
            host.Mark(FrenchParser.ParsePortion("[DR]"), Identification(), false);
            int bottom = 0;
            for (int i = 1; i <= Late.Count(shapes); i++)
            {
                if (Convert.ToString(Late.Get(Late.Item(shapes, i), "Name")).StartsWith("SystemBannerBottom")) bottom++;
            }
            Check(bottom == 0, "PowerPoint Diffusion Restreinte: no bottom stamp");
            host.Mark(GermanParser.ParsePortion("[VS-Vertr.]"));
            names.Clear();
            for (int i = 1; i <= Late.Count(shapes); i++)
            {
                object shape = Late.Item(shapes, i);
                string name = Convert.ToString(Late.Get(shape, "Name"));
                if (name.StartsWith("SystemBanner")) names.Add(name + "=" + Text(Late.Get(Late.Get(shape, "TextFrame"), "TextRange")));
            }
            names.Sort(string.CompareOrdinal);
            all = string.Join(", ", names.ToArray());
            Check(all == "SystemBannerTop2=amtlich geheim gehalten, SystemBannerTop=VS-VERTRAULICH", "PowerPoint VS-VERTRAULICH at the top only: " + all);
            Late.Call(presentation, "Close");
        }
    }
}
