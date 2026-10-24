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
        //id, label and screen tip of each button, in the language of the add-in
        private static string[][] Buttons()
        {
            return new string[][]
            {
                new string[] { "Classify", Markings.Text.T("Classify", "Classifier"), Markings.Text.T("Set the classification of the document", "Fixer la classification du document") },
                new string[] { "PortionMark", Markings.Text.T("Portion mark", "Marque de paragraphe"), Markings.Text.T("Mark the paragraph at the cursor", "Marquer le paragraphe du curseur") },
                new string[] { "RollUp", Markings.Text.T("Roll up", "Calculer"), Markings.Text.T("Build the classification from the portion marks", "Déduire la classification des marques de paragraphe") },
                new string[] { "AuthorityBlock", Markings.Text.T("Authority block", "Bloc d'autorité"), Markings.Text.T("IC classification authority block; French and German first page", "Bloc d'autorité IC ; première page française et allemande") },
                new string[] { "Downgrade", Markings.Text.T("Downgrade", "Déclasser"), Markings.Text.T("Lower, raise or lift the classification of French and German documents", "Déclasser, reclasser ou déclassifier un document français ou allemand") }
            };
        }

        //OfficeCheck.exe toolbar: the Office 2000-2003 toolbar on any Office version, without changing the settings
        public static bool ToolbarAlways;

        private object application;
        private Commands commands;
        private object toolbar;
        private string name = "";
        private readonly List<ComEventLink> links = new List<ComEventLink>();
        private readonly List<ComEventLink> buttonLinks = new List<ComEventLink>();

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
                name = Convert.ToString(Late.Get(application, "Name"));
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
                if (version < 12 || Markings.SystemSettings.Read().Toolbar || ToolbarAlways)
                {
                    CreateToolbar(); //Office 2007 and later ask for the ribbon instead, see GetCustomUI
                    CheckToolbarLater();
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
            //Word ignores Temporary and keeps toolbars in Normal.dot: it must not ask to save Normal.dot because of the add-in
            object normal = name.IndexOf("Word") >= 0 ? Late.TryGet(application, "NormalTemplate") : null;
            bool normalSaved = normal != null && Convert.ToBoolean(Late.TryGet(normal, "Saved") ?? false);
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
                if (normal != null)
                {
                    step = "CustomizationContext";
                    Late.TrySet(application, "CustomizationContext", normal);
                }
                step = "CommandBars.Add";
                toolbar = Late.Call(bars, "Add", Commands.Title, 1, Missing.Value, true); //msoBarTop, temporary
                //shown before the buttons: a button that fails must not leave the whole toolbar hidden
                step = "Visible";
                Late.Set(toolbar, "Visible", true);
                object controls = Late.Get(toolbar, "Controls");
                commands.ChooseLanguage();
                int added = 0;
                foreach (string[] button in Buttons())
                {
                    step = "button " + button[0];
                    try
                    {
                        object control = Late.Call(controls, "Add", 1, Missing.Value, Missing.Value, Missing.Value, true); //msoControlButton
                        Late.Set(control, "Caption", button[1]);
                        Late.Set(control, "Style", 2); //msoButtonCaption
                        Late.Set(control, "TooltipText", button[2]);
                        Late.Set(control, "Tag", "SystemBanner." + button[0]);
                        step = "Click event of " + button[0];
                        buttonLinks.Add(new ComEventLink(control, typeof(CommandBarButtonEvents).GUID, new ButtonSink(commands, button[0])));
                        added++;
                    }
                    catch (Exception ex)
                    {
                        Log.Error("CreateToolbar, " + step, ex); //one button less rather than no toolbar
                    }
                }
                Log.Write("toolbar created with " + added + " buttons, visible " + Late.TryGet(toolbar, "Visible") + ", position " + Late.TryGet(toolbar, "Position"));
            }
            catch (Exception ex)
            {
                Log.Error("CreateToolbar, " + step, ex); //the commands stay unavailable rather than breaking Office's startup
            }
            if (normalSaved)
            {
                Late.TrySet(normal, "Saved", true);
            }
        }

        //Word and Excel 2000 can restore their own toolbar layout after OnStartupComplete (Normal.dot, Excel.xlb):
        //the toolbar is checked again once Office has finished starting, and shown or rebuilt if needed
        private System.Windows.Forms.Timer toolbarCheck;
        private int toolbarChecks;

        private void CheckToolbarLater()
        {
            toolbarCheck = new System.Windows.Forms.Timer();
            toolbarCheck.Interval = 2000;
            toolbarCheck.Tick += delegate
            {
                toolbarChecks++;
                if (toolbarChecks >= 3)
                {
                    toolbarCheck.Stop();
                }
                CheckToolbar();
            };
            toolbarCheck.Start();
        }

        private void CheckToolbar()
        {
            try
            {
                if (application == null) return;
                object bar = null;
                try
                {
                    bar = Late.Item(Late.Get(application, "CommandBars"), Commands.Title);
                }
                catch (Exception)
                {
                    //gone
                }
                if (bar == null)
                {
                    Log.Write("toolbar check " + toolbarChecks + ": toolbar gone, created again");
                    DisposeAll(buttonLinks);
                    CreateToolbar();
                    return;
                }
                object visible = Late.TryGet(bar, "Visible");
                Log.Write("toolbar check " + toolbarChecks + ": visible " + visible + ", position " + Late.TryGet(bar, "Position") + ", " + Late.TryGet(Late.TryGet(bar, "Controls"), "Count") + " buttons");
                if (!Convert.ToBoolean(visible ?? false))
                {
                    Late.Set(bar, "Visible", true);
                    Log.Write("toolbar shown again");
                }
            }
            catch (Exception ex)
            {
                Log.Error("CheckToolbar", ex);
            }
        }

        public string GetCustomUI(string ribbonId)
        {
            Log.Write("GetCustomUI: " + ribbonId);
            if (commands != null)
            {
                commands.ChooseLanguage();
            }
            System.Text.StringBuilder xml = new System.Text.StringBuilder();
            xml.Append("<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"><ribbon><tabs>");
            xml.Append("<tab id=\"SystemBannerTab\" label=\"Classification\"><group id=\"SystemBannerGroup\" label=\"SystemBanner\">");
            foreach (string[] button in Buttons())
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
                case "Downgrade": return "SetPertCalculationsAutomatic";
                default: return "SignatureLineInsert";
            }
        }

        //ribbon callback: Office passes the IRibbonControl, read late bound to need no interop assembly
        public void OnAction(object control)
        {
            try
            {
                string id = Convert.ToString(Late.Get(control, "Id"));
                Log.Write("OnAction " + id); //no such line after a click: Office did not call the add-in
                commands.Run(id);
            }
            catch (Exception ex)
            {
                Log.Error("OnAction", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            Log.Write("OnDisconnection, mode " + removeMode);
            if (toolbarCheck != null)
            {
                toolbarCheck.Stop();
                toolbarCheck = null;
            }
            DisposeAll(links);
            DisposeAll(buttonLinks);
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

        private static void DisposeAll(List<ComEventLink> list)
        {
            foreach (ComEventLink link in list)
            {
                link.Dispose();
            }
            list.Clear();
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }
    }
}
