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

        public static int Main(string[] args)
        {
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
            List<string> apps = new List<string>(args.Length > 0 ? args : new string[] { "word", "excel", "powerpoint" });
            if (apps.Contains("word")) Run("Word.Application", Word);
            if (apps.Contains("excel")) Run("Excel.Application", Excel);
            if (apps.Contains("powerpoint")) Run("PowerPoint.Application", PowerPoint);
            Console.WriteLine(failures == 0 ? "all checks passed" : failures + " check(s) failed");
            return failures == 0 ? 0 : 1;
        }

        private delegate void Scenario(object application);

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
            host.InsertPortionMark("C//NF");
            string secondText = Convert.ToString(Late.Get(Late.Get(Late.Item(Late.Get(document, "Paragraphs"), 2), "Range"), "Text"));
            Check(secondText.StartsWith("(C//NF) Second paragraph."), "Word portion mark replaced: " + secondText.Trim());
            //authority block, written once and replaced
            AuthorityBlock block = new AuthorityBlock();
            block.ClassifiedBy = "J. Doe";
            block.DerivedFrom = "Guide X";
            block.DeclassifyOn = "20450101";
            Late.Call(Late.Get(word, "Selection"), "EndKey", 6); //wdStory
            host.WriteAuthorityBlock(block.Lines());
            block.DeclassifyOn = "20460101";
            host.WriteAuthorityBlock(block.Lines());
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
        }
    }
}
