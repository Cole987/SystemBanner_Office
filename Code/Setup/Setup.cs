//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("SystemBanner Office add-in setup")]
[assembly: AssemblyProduct("SystemBanner")]
[assembly: AssemblyVersion("1.0.0.0")]

namespace SystemBannerOffice.Setup
{
    //SystemBannerOfficeSetup.exe: installs or removes the add-in for every user of the computer, the same steps as
    //InstallSystemBannerOffice.bat, with the add-in, its registry files and the Group Policy templates inside the program.
    //  SystemBannerOfficeSetup.exe            a window with Install and Remove
    //  SystemBannerOfficeSetup.exe /install   /uninstall, with /quiet for deployment scripts: no window, exit code 0 when it worked
    public static class Program
    {
        private static readonly StringBuilder log = new StringBuilder();
        private static TextBox output;

        [STAThread]
        public static int Main(string[] args)
        {
            bool install = false, uninstall = false, quiet = false;
            foreach (string arg in args)
            {
                string a = arg.ToLowerInvariant().TrimStart('/', '-');
                if (a == "install" || a == "i") install = true;
                else if (a == "uninstall" || a == "remove" || a == "u" || a == "x") uninstall = true;
                else if (a == "quiet" || a == "q" || a == "s" || a == "silent") quiet = true;
            }
            if (install || uninstall)
            {
                bool done = Run(install);
                if (!quiet)
                {
                    MessageBox.Show(log.ToString(), Title, MessageBoxButtons.OK, done ? MessageBoxIcon.Information : MessageBoxIcon.Error);
                }
                return done ? 0 : 1;
            }
            Application.EnableVisualStyles();
            Application.Run(new SetupForm());
            return 0;
        }

        private const string Title = "SystemBanner Office";

        //the language of Windows: English, French or German
        public static string T(string english, string french, string german)
        {
            string language = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
            return language == "fr" ? french : (language == "de" ? german : english);
        }

        private static void Say(string line)
        {
            log.AppendLine(line);
            try
            {
                File.AppendAllText(Path.Combine(Path.GetTempPath(), "SystemBannerOfficeSetup.log"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + line + Environment.NewLine);
            }
            catch (Exception)
            {
            }
            if (output != null)
            {
                output.AppendText(line + Environment.NewLine);
                Application.DoEvents();
            }
        }

        //where the add-in goes: the Program Files folder of the bat, the 64 bit one on 64 bit Windows
        private static string Target
        {
            get
            {
                string programFiles = Environment.GetEnvironmentVariable("ProgramW6432");
                if (string.IsNullOrEmpty(programFiles)) programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                return Path.Combine(programFiles, @"SystemBanner\Office");
            }
        }

        private static string Windows
        {
            get { return Environment.GetEnvironmentVariable("WINDIR") ?? @"C:\WINDOWS"; }
        }

        private static bool IsAdministrator()
        {
            try
            {
                return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static List<string> RunningOffice()
        {
            List<string> running = new List<string>();
            foreach (string name in new string[] { "WINWORD", "EXCEL", "POWERPNT" })
            {
                if (Process.GetProcessesByName(name).Length > 0) running.Add(name);
            }
            return running;
        }

        public static bool Run(bool install)
        {
            log.Length = 0;
            try
            {
                if (!IsAdministrator())
                {
                    Say(T("Run the setup as an administrator: it writes in Program Files and in HKEY_LOCAL_MACHINE.",
                        "Lancer l'installation en administrateur : elle écrit dans Program Files et dans HKEY_LOCAL_MACHINE.",
                        "Starten Sie die Installation als Administrator: Sie schreibt in Program Files und in HKEY_LOCAL_MACHINE."));
                    return false;
                }
                List<string> running = RunningOffice();
                if (running.Count > 0)
                {
                    Say(T("Close Word, Excel and PowerPoint first (running: ", "Fermer d'abord Word, Excel et PowerPoint (ouverts : ", "Schließen Sie zuerst Word, Excel und PowerPoint (geöffnet: ") + string.Join(", ", running.ToArray()) + ").");
                    return false;
                }
                if (install) Install(); else Uninstall();
                Say("");
                Say(install ? T("SystemBanner Office add-in installed: start Word, Excel or PowerPoint.", "Complément Office SystemBanner installé : lancer Word, Excel ou PowerPoint.", "SystemBanner Office-Add-in installiert: Starten Sie Word, Excel oder PowerPoint.")
                    : T("SystemBanner Office add-in removed.", "Complément Office SystemBanner retiré.", "SystemBanner Office-Add-in entfernt."));
                return true;
            }
            catch (Exception ex)
            {
                Say("");
                Say(T("Failed: ", "Échec : ", "Fehlgeschlagen: ") + ex.Message);
                return false;
            }
        }

        private static void Install()
        {
            string target = Target;
            Directory.CreateDirectory(target);
            foreach (string file in new string[] { "SystemBannerOffice.dll", "SystemBannerOffice.Markings.dll" })
            {
                Extract(file, Path.Combine(target, file));
            }
            Say(T("Files copied to ", "Fichiers copiés dans ", "Dateien kopiert nach ") + target);
            string self = Assembly.GetExecutingAssembly().Location;
            string setupCopy = Path.Combine(target, "SystemBannerOfficeSetup.exe");
            if (!string.Equals(Path.GetFullPath(self), Path.GetFullPath(setupCopy), StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(self, setupCopy, true); //kept for Add or Remove Programs
            }
            string addIn = Path.Combine(target, "SystemBannerOffice.dll");
            //32 bit Office (every Office from 2000 to 2007, and 32 bit Office 2010 and later)
            Exec(Path.Combine(Windows, @"Microsoft.NET\Framework\v2.0.50727\RegAsm.exe"), "/nologo /codebase \"" + addIn + "\"", true);
            string reg = Path.Combine(target, "SystemBannerOffice.reg");
            Extract("SystemBannerOffice.reg", reg);
            Regedit(reg);
            Say(T("Registered in COM and in Word, Excel and PowerPoint.", "Enregistré dans COM et dans Word, Excel et PowerPoint.", "In COM und in Word, Excel und PowerPoint registriert."));
            OfficeConfigs(true);
            PolicyTemplates(true);
            using (RegistryKey key = Registry.LocalMachine.CreateSubKey(UninstallKey))
            {
                key.SetValue("DisplayName", "SystemBanner Office add-in");
                key.SetValue("DisplayVersion", "1.0.0");
                key.SetValue("Publisher", "SystemBanner");
                key.SetValue("URLInfoAbout", "https://github.com/Cole987/SystemBanner_Office");
                key.SetValue("InstallLocation", target);
                key.SetValue("UninstallString", "\"" + setupCopy + "\" /uninstall");
                key.SetValue("QuietUninstallString", "\"" + setupCopy + "\" /uninstall /quiet");
                key.SetValue("NoModify", 1);
                key.SetValue("NoRepair", 1);
            }
            Say(T("Added to Add or Remove Programs.", "Ajouté à Ajout/Suppression de programmes.", "Unter Software hinzugefügt."));
        }

        private const string UninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SystemBannerOffice";

        private static void Uninstall()
        {
            string target = Target;
            string addIn = Path.Combine(target, "SystemBannerOffice.dll");
            string reg = Path.Combine(Path.GetTempPath(), "SystemBannerOffice_remove.reg");
            Extract("SystemBannerOffice_remove.reg", reg);
            Regedit(reg);
            File.Delete(reg);
            if (File.Exists(addIn))
            {
                Exec(Path.Combine(Windows, @"Microsoft.NET\Framework\v2.0.50727\RegAsm.exe"), "/nologo /unregister \"" + addIn + "\"", false);
                string regAsm64 = Path.Combine(Windows, @"Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe");
                if (File.Exists(regAsm64)) Exec(regAsm64, "/nologo /unregister \"" + addIn + "\"", false);
            }
            Say(T("Removed from COM and from Word, Excel and PowerPoint.", "Retiré de COM et de Word, Excel et PowerPoint.", "Aus COM und aus Word, Excel und PowerPoint entfernt."));
            OfficeConfigs(false);
            PolicyTemplates(false);
            Registry.LocalMachine.DeleteSubKey(UninstallKey, false);
            foreach (string file in new string[] { "SystemBannerOffice.dll", "SystemBannerOffice.Markings.dll", "SystemBannerOffice.reg", "SystemBannerOfficeSetup.exe" })
            {
                string path = Path.Combine(target, file);
                if (!File.Exists(path)) continue;
                try
                {
                    File.Delete(path);
                }
                catch (Exception)
                {
                    MoveFileEx(path, null, 4); //in use (this program, or a DLL Office still holds): deleted at the next restart
                }
            }
            if (Directory.Exists(target) && Directory.GetFileSystemEntries(target).Length == 0)
            {
                Directory.Delete(target);
                string parent = Path.GetDirectoryName(target);
                if (Directory.Exists(parent) && Directory.GetFileSystemEntries(parent).Length == 0) Directory.Delete(parent);
            }
            else if (Directory.Exists(target))
            {
                MoveFileEx(target, null, 4);
            }
            Say(T("Files removed from ", "Fichiers retirés de ", "Dateien entfernt aus ") + target);
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern bool MoveFileEx(string existing, string replacement, int flags);

        //the 32 bit registry view, and on 64 bit Windows the 64 bit one, for 64 bit Office: regedit, as the bat does,
        //since the .NET Framework 2.0 cannot choose the registry view
        private static void Regedit(string file)
        {
            string wow = Path.Combine(Windows, @"SysWOW64\regedit.exe");
            string native = Path.Combine(Windows, "regedit.exe");
            Exec(File.Exists(wow) ? wow : native, "/s \"" + file + "\"", true);
            string regAsm64 = Path.Combine(Windows, @"Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe");
            if (File.Exists(wow) && File.Exists(regAsm64))
            {
                if (file.IndexOf("_remove", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    Exec(regAsm64, "/nologo /codebase \"" + Path.Combine(Target, "SystemBannerOffice.dll") + "\"", true);
                }
                Exec(native, "/s \"" + file + "\"", true);
            }
        }

        //Office 2000 to 2003 start the .NET runtime of their first managed add-in, which can be .NET 1.1:
        //a config file next to each program asks for .NET 2.0; an existing file is never replaced, nor removed
        private static void OfficeConfigs(bool install)
        {
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            foreach (string folder in new string[] { @"Microsoft Office\Office", @"Microsoft Office\Office10", @"Microsoft Office\Office11" })
            {
                foreach (string exe in new string[] { "WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE" })
                {
                    string program = Path.Combine(Path.Combine(programFiles, folder), exe);
                    string config = program + ".config";
                    bool ours = File.Exists(config) && File.ReadAllText(config).IndexOf("SystemBanner Office add-in installer") >= 0;
                    if (install && File.Exists(program) && !File.Exists(config))
                    {
                        Extract("SystemBannerOffice.exe.config", config);
                        Say(T("Office 2000 to 2003 asked to start .NET 2.0: ", "Office 2000 à 2003 réglé pour démarrer .NET 2.0 : ", "Office 2000 bis 2003 auf .NET 2.0 eingestellt: ") + config);
                    }
                    else if (install && File.Exists(config) && !ours)
                    {
                        Say(T("Warning: ", "Attention : ", "Warnung: ") + config + T(" was not written by SystemBanner: add supportedRuntime v2.0.50727 to it.", " n'a pas été écrit par SystemBanner : y ajouter supportedRuntime v2.0.50727.", " wurde nicht von SystemBanner geschrieben: Fügen Sie supportedRuntime v2.0.50727 hinzu."));
                    }
                    else if (!install && ours)
                    {
                        File.Delete(config);
                    }
                }
            }
        }

        //the Group Policy templates, where the Group Policy editor of the computer reads them, like the SystemBanner installer:
        //ADMX and ADML in PolicyDefinitions from Windows Vista, the ADM files in inf on Windows 2000 and XP
        private static void PolicyTemplates(bool install)
        {
            string definitions = Path.Combine(Windows, "PolicyDefinitions");
            if (Directory.Exists(definitions))
            {
                Template("SystemBannerOffice.admx", Path.Combine(definitions, "SystemBannerOffice.admx"), install);
                foreach (string language in new string[] { "en-US", "fr-FR", "de-DE" })
                {
                    string folder = Path.Combine(definitions, language);
                    if (install && !Directory.Exists(folder))
                    {
                        if (language != "en-US") continue; //only the languages of this Windows
                        Directory.CreateDirectory(folder);
                    }
                    Template(language + ".SystemBannerOffice.adml", Path.Combine(folder, "SystemBannerOffice.adml"), install);
                }
                Say(install ? T("Group Policy templates copied to ", "Modèles de stratégie copiés dans ", "Gruppenrichtlinienvorlagen kopiert nach ") + definitions
                    : T("Group Policy templates removed.", "Modèles de stratégie retirés.", "Gruppenrichtlinienvorlagen entfernt."));
            }
            else
            {
                string inf = Path.Combine(Windows, "inf");
                foreach (string adm in new string[] { "SystemBannerOffice.adm", "SystemBannerOffice_FR.adm", "SystemBannerOffice_DE.adm" })
                {
                    Template(adm, Path.Combine(inf, adm), install);
                }
                Say(install ? T("Group Policy templates (ADM) copied to ", "Modèles de stratégie (ADM) copiés dans ", "Gruppenrichtlinienvorlagen (ADM) kopiert nach ") + inf
                    : T("Group Policy templates removed.", "Modèles de stratégie retirés.", "Gruppenrichtlinienvorlagen entfernt."));
            }
        }

        private static void Template(string resource, string path, bool install)
        {
            if (install)
            {
                Extract(resource, path);
            }
            else if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private static void Extract(string resource, string path)
        {
            using (Stream input = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
            {
                if (input == null) throw new InvalidOperationException(resource + T(" is missing from the setup program.", " manque dans le programme d'installation.", " fehlt im Installationsprogramm."));
                using (FileStream file = File.Create(path))
                {
                    byte[] buffer = new byte[65536];
                    int read;
                    while ((read = input.Read(buffer, 0, buffer.Length)) > 0) file.Write(buffer, 0, read);
                }
            }
        }

        private static void Exec(string program, string arguments, bool required)
        {
            if (!File.Exists(program))
            {
                if (required) throw new InvalidOperationException(program + T(" was not found: install the .NET Framework 2.0 or later.", " est introuvable : installer le .NET Framework 2.0 ou plus récent.", " wurde nicht gefunden: Installieren Sie das .NET Framework 2.0 oder höher."));
                return;
            }
            ProcessStartInfo start = new ProcessStartInfo(program, arguments);
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.RedirectStandardOutput = true;
            start.RedirectStandardError = true;
            using (Process process = Process.Start(start))
            {
                string text = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    Say(Path.GetFileName(program) + ": " + text.Trim());
                    if (required) throw new InvalidOperationException(Path.GetFileName(program) + T(" failed, exit code ", " a échoué, code ", " ist fehlgeschlagen, Code ") + process.ExitCode + ".");
                }
            }
        }

        private class SetupForm : Form
        {
            public SetupForm()
            {
                Text = Title;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                StartPosition = FormStartPosition.CenterScreen;
                Font = new Font("Tahoma", 8.25f);
                ClientSize = new Size(480, 330);
                Label title = new Label();
                title.Text = "SystemBanner Office";
                title.Font = new Font("Tahoma", 14f, FontStyle.Bold);
                title.SetBounds(12, 10, 456, 28);
                Controls.Add(title);
                Label about = new Label();
                about.Text = T("Classification markings for Word, Excel and PowerPoint, from Office 2000 to Microsoft 365. Installs for every user of this computer; Word, Excel and PowerPoint must be closed.",
                    "Marquages de classification pour Word, Excel et PowerPoint, d'Office 2000 à Microsoft 365. S'installe pour tous les utilisateurs du poste ; Word, Excel et PowerPoint doivent être fermés.",
                    "Einstufungskennzeichnungen für Word, Excel und PowerPoint, von Office 2000 bis Microsoft 365. Wird für alle Benutzer dieses Rechners installiert; Word, Excel und PowerPoint müssen geschlossen sein.");
                about.SetBounds(12, 44, 456, 44);
                Controls.Add(about);
                output = new TextBox();
                output.Multiline = true;
                output.ReadOnly = true;
                output.ScrollBars = ScrollBars.Vertical;
                output.SetBounds(12, 92, 456, 194);
                Controls.Add(output);
                Button install = Add(T("Install", "Installer", "Installieren"), 222);
                Button remove = Add(T("Remove", "Retirer", "Entfernen"), 306);
                Button close = Add(T("Close", "Fermer", "Schließen"), 390);
                install.Click += delegate { Act(true); };
                remove.Click += delegate { Act(false); };
                close.Click += delegate { Close(); };
                AcceptButton = install;
                CancelButton = close;
                remove.Enabled = Registry.LocalMachine.OpenSubKey(UninstallKey) != null || File.Exists(Path.Combine(Target, "SystemBannerOffice.dll"));
            }

            private Button Add(string text, int x)
            {
                Button button = new Button();
                button.Text = text;
                button.SetBounds(x, 296, 78, 24);
                Controls.Add(button);
                return button;
            }

            private void Act(bool install)
            {
                output.Clear();
                Cursor = Cursors.WaitCursor;
                try
                {
                    Run(install);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }
            }
        }
    }
}
