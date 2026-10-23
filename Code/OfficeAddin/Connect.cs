//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("SystemBanner Office add-in")]
[assembly: AssemblyProduct("SystemBanner")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: ComVisible(false)]

namespace SystemBannerOffice
{
    //Word application events: DocumentBeforePrint and DocumentBeforeSave of ApplicationEvents2 (Word 2000 and later)
    //ComImport: RegAsm must never register these Office interfaces again
    [ComImport, Guid("000209FE-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface WordApplicationEvents
    {
        [DispId(7)] void DocumentBeforePrint([MarshalAs(UnmanagedType.IDispatch)] object document, ref bool cancel);
        [DispId(8)] void DocumentBeforeSave([MarshalAs(UnmanagedType.IDispatch)] object document, ref bool saveAsUi, ref bool cancel);
    }

    //Excel application events: WorkbookBeforeSave and WorkbookBeforePrint of AppEvents (Excel 2000 and later)
    [ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ExcelApplicationEvents
    {
        [DispId(0x623)] void WorkbookBeforeSave([MarshalAs(UnmanagedType.IDispatch)] object workbook, bool saveAsUi, ref bool cancel);
        [DispId(0x624)] void WorkbookBeforePrint([MarshalAs(UnmanagedType.IDispatch)] object workbook, ref bool cancel);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class WordEventSink : WordApplicationEvents
    {
        private readonly Commands commands;
        public WordEventSink(Commands commands) { this.commands = commands; }
        public void DocumentBeforePrint(object document, ref bool cancel) { cancel |= commands.BlockUnclassified(document, "printed"); }
        public void DocumentBeforeSave(object document, ref bool saveAsUi, ref bool cancel) { cancel |= commands.BlockUnclassified(document, "saved"); }
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class ExcelEventSink : ExcelApplicationEvents
    {
        private readonly Commands commands;
        public ExcelEventSink(Commands commands) { this.commands = commands; }
        public void WorkbookBeforeSave(object workbook, bool saveAsUi, ref bool cancel) { cancel |= commands.BlockUnclassified(workbook, "saved"); }
        public void WorkbookBeforePrint(object workbook, ref bool cancel) { cancel |= commands.BlockUnclassified(workbook, "printed"); }
    }

    //a toolbar button of Office 2000 to 2003
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class ButtonSink : CommandBarButtonEvents
    {
        private readonly Commands commands;
        private readonly string command;
        public ButtonSink(Commands commands, string command) { this.commands = commands; this.command = command; }
        public void Click(object ctrl, ref bool cancelDefault) { commands.Run(command); }
    }

    //the add-in Office loads, registered under Software\Microsoft\Office\[Word|Excel|PowerPoint]\Addins\SystemBannerOffice.Connect
    [ComVisible(true), Guid("6F1C1F5E-5B0A-4E3B-9C55-3D2B7A0B6E41"), ProgId("SystemBannerOffice.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string[][] Buttons = new string[][]
        {
            new string[] { "Classify", "Classify", "Set the classification of the document" },
            new string[] { "PortionMark", "Portion mark", "Mark the paragraph at the cursor" },
            new string[] { "RollUp", "Roll up", "Build the banner line from the portion marks" },
            new string[] { "AuthorityBlock", "Authority block", "Classified By, Derived From, Declassify On" }
        };

        private object application;
        private Commands commands;
        private object toolbar;
        private readonly List<ComEventLink> links = new List<ComEventLink>();

        public Connect()
        {
            Log.Write("add-in created, .NET " + Environment.Version + ", " + IntPtr.Size * 8 + " bit, " + Environment.OSVersion);
        }

        //every method Office calls catches its errors: an exception thrown back into Office can stop it from starting
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                this.application = application;
                commands = new Commands(application);
                string name = Convert.ToString(Late.Get(application, "Name"));
                Log.Write("OnConnection: " + name + " " + Late.Get(application, "Version") + ", connect mode " + connectMode);
                //the save and print check is only wired when the administrator asks for it, one less hook in Office otherwise
                if (!Markings.SystemSettings.Read().RequireClassification)
                {
                    return;
                }
                if (name.IndexOf("Word") >= 0)
                {
                    links.Add(new ComEventLink(application, typeof(WordApplicationEvents).GUID, new WordEventSink(commands)));
                }
                else if (name.IndexOf("Excel") >= 0)
                {
                    links.Add(new ComEventLink(application, typeof(ExcelApplicationEvents).GUID, new ExcelEventSink(commands)));
                }
                Log.Write("save and print check wired");
            }
            catch (Exception ex)
            {
                Log.Error("OnConnection", ex); //no save and print check rather than no add-in
            }
            if (connectMode == ext_ConnectMode.ext_cm_AfterStartup)
            {
                //ext_cm_AfterStartup: loaded from the COM add-ins dialog, Office will not call OnStartupComplete
                Array none = new object[0];
                OnStartupComplete(ref none);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                int version = MajorVersion();
                Log.Write("OnStartupComplete: Office major version " + version);
                if (version < 12 || Markings.SystemSettings.Read().Toolbar)
                {
                    CreateToolbar(); //Office 2007 and later ask for the ribbon instead, see GetCustomUI
                }
            }
            catch (Exception ex)
            {
                Log.Error("OnStartupComplete", ex);
            }
        }

        private int MajorVersion()
        {
            string version = Convert.ToString(Late.Get(application, "Version"));
            int dot = version.IndexOf('.');
            int major;
            return int.TryParse(dot > 0 ? version.Substring(0, dot) : version, out major) ? major : 0;
        }

        //a temporary toolbar, rebuilt at each start so nothing stays behind in Normal.dot when the add-in is removed
        private void CreateToolbar()
        {
            string step = "CommandBars";
            try
            {
                object bars = Late.Get(application, "CommandBars");
                try
                {
                    Late.Call(Late.Item(bars, Commands.Title), "Delete");
                    Log.Write("toolbar left from a previous session deleted");
                }
                catch (Exception)
                {
                    //no toolbar left from a previous session
                }
                step = "CommandBars.Add";
                toolbar = Late.Call(bars, "Add", Commands.Title, 1, Missing.Value, true); //msoBarTop, temporary
                object controls = Late.Get(toolbar, "Controls");
                foreach (string[] button in Buttons)
                {
                    step = "button " + button[0];
                    object control = Late.Call(controls, "Add", 1, Missing.Value, Missing.Value, Missing.Value, true); //msoControlButton
                    Late.Set(control, "Caption", button[1]);
                    Late.Set(control, "Style", 2); //msoButtonCaption
                    Late.Set(control, "TooltipText", button[2]);
                    Late.Set(control, "Tag", "SystemBanner." + button[0]);
                    step = "Click event of " + button[0];
                    links.Add(new ComEventLink(control, typeof(CommandBarButtonEvents).GUID, new ButtonSink(commands, button[0])));
                }
                step = "Visible";
                Late.Set(toolbar, "Visible", true);
                Log.Write("toolbar created with " + Buttons.Length + " buttons");
            }
            catch (Exception ex)
            {
                Log.Error("CreateToolbar, " + step, ex); //the commands stay unavailable rather than breaking Office's startup
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            Log.Write("GetCustomUI: " + ribbonId);
            System.Text.StringBuilder xml = new System.Text.StringBuilder();
            xml.Append("<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"><ribbon><tabs>");
            xml.Append("<tab id=\"SystemBannerTab\" label=\"Classification\"><group id=\"SystemBannerGroup\" label=\"SystemBanner\">");
            foreach (string[] button in Buttons)
            {
                xml.Append("<button id=\"").Append(button[0]).Append("\" label=\"").Append(button[1]).Append("\" screentip=\"").Append(button[2])
                    .Append("\" size=\"large\" imageMso=\"").Append(Icon(button[0])).Append("\" onAction=\"OnAction\"/>");
            }
            xml.Append("</group></tab></tabs></ribbon></customUI>");
            return xml.ToString();
        }

        private static string Icon(string command)
        {
            switch (command)
            {
                case "Classify": return "Lock";
                case "PortionMark": return "ParagraphMarks";
                case "RollUp": return "RefreshAll";
                default: return "SignatureLineInsert";
            }
        }

        //ribbon callback: Office passes the IRibbonControl, read late bound to need no interop assembly
        public void OnAction(object control)
        {
            try
            {
                commands.Run(Convert.ToString(Late.Get(control, "Id")));
            }
            catch (Exception ex)
            {
                Log.Error("OnAction", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Log.Write("OnDisconnection, mode " + removeMode);
            foreach (ComEventLink link in links)
            {
                link.Dispose();
            }
            links.Clear();
            if (toolbar != null)
            {
                try
                {
                    Late.Call(toolbar, "Delete");
                }
                catch (Exception)
                {
                }
                toolbar = null;
            }
            application = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }
    }
}
