//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.IO;
using System.Text;
using Microsoft.Win32;

namespace SystemBannerOffice.Check
{
    //OfficeCheck.exe diag: what Windows and Office see of the add-in, written to SystemBannerOffice_diag.txt on the Desktop,
    //like SystemBanner.exe /diag
    public static class Diag
    {
        private const string Clsid = "{6F1C1F5E-5B0A-4E3B-9C55-3D2B7A0B6E41}";
        private const string ProgId = "SystemBannerOffice.Connect";
        private static readonly string[] Apps = new string[] { "Word", "Excel", "PowerPoint" };
        //Office 2000, XP, 2003, 2007, 2010
        private static readonly string[] Versions = new string[] { "9.0", "10.0", "11.0", "12.0", "14.0", "15.0", "16.0" };

        public static int Run()
        {
            StringBuilder report = new StringBuilder();
            report.AppendLine("SystemBanner Office diagnostics, " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            report.AppendLine("Windows: " + Environment.OSVersion + ", 64 bit Windows: " + Is64BitWindows() + ", this process: " + IntPtr.Size * 8 + " bit, .NET running: " + Environment.Version);
            report.AppendLine(".NET installed: " + string.Join(", ", NetVersions()));
            report.AppendLine();

            report.AppendLine("== Office installed (InstallRoot) ==");
            foreach (string version in Versions)
            {
                foreach (string view in Views())
                {
                    string path = Value(Registry.LocalMachine, view + @"Microsoft\Office\" + version + @"\Common\InstallRoot", "Path");
                    if (path != null) report.AppendLine("  " + version + " (" + view + "): " + path);
                }
            }
            report.AppendLine();

            report.AppendLine("== Add-in registration in Office (LoadBehavior 3 = load at startup) ==");
            foreach (string app in Apps)
            {
                foreach (string view in Views())
                {
                    Key(report, Registry.LocalMachine, "HKLM", view + @"Microsoft\Office\" + app + @"\Addins\" + ProgId);
                }
                Key(report, Registry.CurrentUser, "HKCU", @"SOFTWARE\Microsoft\Office\" + app + @"\Addins\" + ProgId);
            }
            report.AppendLine();

            report.AppendLine("== COM registration ==");
            Key(report, Registry.ClassesRoot, "HKCR", ProgId + @"\CLSID");
            string[] classRoots = Is64BitWindows() ? new string[] { @"CLSID\", @"Wow6432Node\CLSID\" } : new string[] { @"CLSID\" };
            foreach (string root in classRoots)
            {
                string inproc = root + Clsid + @"\InprocServer32";
                Key(report, Registry.ClassesRoot, "HKCR", inproc);
                string codeBase = Value(Registry.ClassesRoot, inproc, "CodeBase");
                if (codeBase != null)
                {
                    try
                    {
                        string file = new Uri(codeBase).LocalPath;
                        report.AppendLine("    file " + file + (File.Exists(file) ? " exists" : " IS MISSING"));
                        string markings = Path.Combine(Path.GetDirectoryName(file), "SystemBannerOffice.Markings.dll");
                        report.AppendLine("    file " + markings + (File.Exists(markings) ? " exists" : " IS MISSING"));
                    }
                    catch (Exception ex)
                    {
                        report.AppendLine("    CodeBase not readable: " + ex.Message);
                    }
                }
            }
            report.AppendLine();

            report.AppendLine("== Add-ins Office disabled after a problem (Resiliency) ==");
            foreach (string version in Versions)
            {
                foreach (string app in Apps)
                {
                    string resiliency = @"Software\Microsoft\Office\" + version + @"\" + app + @"\Resiliency";
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(resiliency))
                    {
                        if (key == null) continue;
                        foreach (string sub in key.GetSubKeyNames())
                        {
                            using (RegistryKey child = key.OpenSubKey(sub))
                            {
                                report.AppendLine("  HKCU\\" + resiliency + "\\" + sub + ": " + child.ValueCount + " value(s)");
                            }
                        }
                    }
                }
            }
            report.AppendLine();

            report.AppendLine("== Creating the add-in outside Office ==");
            try
            {
                Type type = Type.GetTypeFromProgID(ProgId);
                if (type == null)
                {
                    report.AppendLine("  " + ProgId + " is not registered");
                }
                else
                {
                    object addIn = Activator.CreateInstance(type);
                    report.AppendLine("  created: " + addIn.GetType().FullName);
                }
            }
            catch (Exception ex)
            {
                report.AppendLine("  FAILED: " + (ex.InnerException ?? ex));
            }
            report.AppendLine();

            report.AppendLine("== End of " + Log.FilePath + " ==");
            if (File.Exists(Log.FilePath))
            {
                string[] lines = File.ReadAllLines(Log.FilePath);
                for (int i = Math.Max(0, lines.Length - 60); i < lines.Length; i++)
                {
                    report.AppendLine("  " + lines[i]);
                }
            }
            else
            {
                report.AppendLine("  no log: the add-in has never been started by Office for this user");
            }

            string output = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "SystemBannerOffice_diag.txt");
            File.WriteAllText(output, report.ToString());
            Console.WriteLine(report.ToString());
            Console.WriteLine("Written to " + output);
            return 0;
        }

        private static bool Is64BitWindows()
        {
            return IntPtr.Size == 8 || !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"));
        }

        //.NET 2.0 has no registry views: the 32 bit view of a 64 bit Windows is read through Wow6432Node
        private static string[] Views()
        {
            return Is64BitWindows() ? new string[] { @"SOFTWARE\", @"SOFTWARE\Wow6432Node\" } : new string[] { @"SOFTWARE\" };
        }

        private static string[] NetVersions()
        {
            string framework = Path.Combine(Environment.GetEnvironmentVariable("WINDIR"), @"Microsoft.NET\Framework");
            if (!Directory.Exists(framework)) return new string[0];
            System.Collections.Generic.List<string> versions = new System.Collections.Generic.List<string>();
            foreach (string dir in Directory.GetDirectories(framework, "v*"))
            {
                if (File.Exists(Path.Combine(dir, "mscorlib.dll")) || File.Exists(Path.Combine(dir, "csc.exe")))
                {
                    versions.Add(Path.GetFileName(dir));
                }
            }
            return versions.ToArray();
        }

        private static string Value(RegistryKey root, string path, string name)
        {
            try
            {
                using (RegistryKey key = root.OpenSubKey(path))
                {
                    return key == null ? null : Convert.ToString(key.GetValue(name));
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void Key(StringBuilder report, RegistryKey root, string rootName, string path)
        {
            try
            {
                using (RegistryKey key = root.OpenSubKey(path))
                {
                    if (key == null)
                    {
                        report.AppendLine("  " + rootName + "\\" + path + ": none");
                        return;
                    }
                    report.AppendLine("  " + rootName + "\\" + path + ":");
                    foreach (string name in key.GetValueNames())
                    {
                        report.AppendLine("    " + (name.Length == 0 ? "(default)" : name) + " = " + key.GetValue(name));
                    }
                }
            }
            catch (Exception ex)
            {
                report.AppendLine("  " + rootName + "\\" + path + ": " + ex.Message);
            }
        }
    }
}
