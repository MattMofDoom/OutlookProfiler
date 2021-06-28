using System;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.Win32;
using ngbt.regis3;

// ReSharper disable SwitchStatementHandlesSomeKnownEnumValuesWithDefault
// ReSharper disable LocalizableElement
// ReSharper disable ArrangeTypeModifiers

namespace OutlookProfiler
{
    public static class StringExtensions
    {
        private const StringComparison DefaultComparison = StringComparison.OrdinalIgnoreCase;

        public static string ReplaceCaseless(this string str, string oldValue, string newValue)
        {
            var sb = new StringBuilder();

            var previousIndex = 0;
            var index = str.IndexOf(oldValue, DefaultComparison);
            while (index != -1)
            {
                sb.Append(str.Substring(previousIndex, index - previousIndex));
                sb.Append(newValue);
                index += oldValue.Length;

                previousIndex = index;
                index = str.IndexOf(oldValue, index, DefaultComparison);
            }

            sb.Append(str.Substring(previousIndex));

            return sb.ToString();
        }
    }

    static class Program
    {
        private static bool _isLog;
        private static string _logPath = string.Empty;

        private static void CreateLog(string path)
        {
            try
            {
                var logFile = File.Create(path, 1024, FileOptions.WriteThrough);
                logFile.Close();
                _logPath = path;
                _isLog = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log file cannot be created - logging will not be performed!");
                Console.WriteLine("Exception details: " + ex);
                _isLog = false;
            }
        }

        private static void Output(string text = null)
        {
            if (!string.IsNullOrEmpty(text))
            {
                Console.WriteLine(text);
                if (_isLog)
                    File.AppendAllText(_logPath, text + Environment.NewLine);
            }
            else
            {
                Console.WriteLine();
                if (_isLog)
                    File.AppendAllText(_logPath, Environment.NewLine);
            }
        }

        private static string GetDefault(string profile)
        {
            var profilePath = Registry.CurrentUser.OpenSubKey(profile);
            if (profilePath == null)
                return string.Empty;

            return (string) profilePath.GetValue("DefaultProfile");
        }

        private static bool SetDefault(string profile, string profileName, OfficeVersion version)
        {
            var profilePath = Registry.CurrentUser.OpenSubKey(profile, true);
            if (profilePath == null)
                return false;

            try
            {
                profilePath.SetValue("DefaultProfile", profileName);
            }
            catch (Exception ex)
            {
                Output("Failed to set default profile: " + ex.Message);
                Output("Exception details: " + ex);
                return false;
            }

            try
            {
                //Set the first run values so Outlook won't prompt to create a profile
                var setupPath = Registry.CurrentUser.OpenSubKey(profile + "\\Setup");
                setupPath = setupPath == null
                    ? Registry.CurrentUser.CreateSubKey(profile + "\\Setup", true)
                    : Registry.CurrentUser.OpenSubKey(profile + "\\Setup", true);

                switch (version)
                {
                    case OfficeVersion.Office2016:
                        if (File.Exists("C:\\Program Files (x86)\\Microsoft Office\\Custom16.prf"))
                            setupPath?.SetValue("ImportPrf", "C:\\Program Files (x86)\\Microsoft Office\\Custom16.prf");
                        else if (File.Exists("C:\\Program Files (x86)\\Microsoft Office\\Custom15.prf"))
                            setupPath?.SetValue("ImportPrf", "C:\\Program Files (x86)\\Microsoft Office\\Custom15.prf");
                        break;
                    default:
                        if (File.Exists("C:\\Program Files (x86)\\Microsoft Office\\Custom15.prf"))
                            setupPath?.SetValue("ImportPrf", "C:\\Program Files (x86)\\Microsoft Office\\Custom15.prf");
                        break;
                }

                return true;
            }
            catch (Exception ex)
            {
                Output("Failed to set first run values:" + ex.Message);
                Output("Exception details: " + ex);
                return false;
            }
        }

        private static bool ProfileExists(string profile)
        {
            var profilePath = Registry.CurrentUser.OpenSubKey(profile);
            return profilePath != null;
        }

        private static void Main(string[] args)
        {
            var operation = OperationType.None;
            var version = OfficeVersion.Office2013;
            var sourceProfile = "Outlook";
            var targetProfile = "Email";
            const string rootKey = "HKEY_CURRENT_USER\\";
            const string exchangeOptionsPath = "Software\\Microsoft\\Exchange\\Client\\Options";
            const string office2010ProfilePath =
                "Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows Messaging Subsystem\\Profiles";
            const string outlookProfilePath = "Software\\Microsoft\\Office\\{0}\\Outlook";
            var default2013Profile = string.Format(outlookProfilePath, "15.0");
            var default2016Profile = string.Format(outlookProfilePath, "16.0");
            var office2013Profile = string.Format(outlookProfilePath, "15.0") + "\\Profiles";
            var office2016Profile = string.Format(outlookProfilePath, "16.0") + "\\Profiles";
            var isDefault = false;
            var ignoreDefault = false;
            var useOptions = false;

            var path = string.Empty;
            var optionsPath = string.Empty;

            if (!args.Length.Equals(0))
                foreach (var arg in args)
                {
                    if (arg.StartsWith("export", StringComparison.CurrentCultureIgnoreCase) ||
                        arg.StartsWith("import", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var argPath = arg.Split('=');
                        if (!string.IsNullOrEmpty(argPath[1]) && Directory.Exists(Path.GetDirectoryName(argPath[1])) &&
                            !string.IsNullOrEmpty(Path.GetFileName(argPath[1])))
                        {
                            if (arg.StartsWith("export2013", StringComparison.CurrentCultureIgnoreCase))
                                operation = OperationType.Export2013;
                            else if (arg.StartsWith("export2016", StringComparison.CurrentCultureIgnoreCase))
                                operation = OperationType.Export2016;
                            else if (arg.StartsWith("export2010", StringComparison.CurrentCultureIgnoreCase))
                                operation = OperationType.Export2010;
                            else
                                operation = OperationType.Import;
                            path = argPath[1];
                        }
                    }

                    if (arg.StartsWith("options", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var argPath = arg.Split('=');
                        if (!string.IsNullOrEmpty(argPath[1]) && Directory.Exists(Path.GetDirectoryName(argPath[1])) &&
                            !string.IsNullOrEmpty(Path.GetFileName(argPath[1])))
                        {
                            optionsPath = argPath[1];
                            useOptions = true;
                        }
                    }

                    if (arg.StartsWith("targetProfile", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var argPath = arg.Split('=');
                        if (!string.IsNullOrEmpty(argPath[1])) targetProfile = argPath[1];
                    }

                    if (arg.StartsWith("sourceProfile", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var argPath = arg.Split('=');
                        if (!string.IsNullOrEmpty(argPath[1])) sourceProfile = argPath[1];
                    }

                    if (arg.StartsWith("targetVersion", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var argPath = arg.Split('=');
                        if (!string.IsNullOrEmpty(argPath[1]))
                            switch (argPath[1])
                            {
                                case "2016":
                                case "Office2016":
                                    version = OfficeVersion.Office2016;
                                    break;
                                default:
                                    version = OfficeVersion.Office2013;
                                    break;
                            }
                    }


                    if (arg.StartsWith("log", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var argPath = arg.Split('=');
                        if (!string.IsNullOrEmpty(argPath[1]) && Directory.Exists(Path.GetDirectoryName(argPath[1])) &&
                            !string.IsNullOrEmpty(Path.GetFileName(argPath[1])))
                            CreateLog(argPath[1]);
                    }

                    if (arg.StartsWith("IgnoreDefault", StringComparison.CurrentCultureIgnoreCase))
                        ignoreDefault = true;
                }

            Output("*** Outlook Profiler v" + Assembly.GetEntryAssembly()?.GetName().Version + " ***");
            Output();
            Output("Operation Type: " + operation);
            if (_isLog)
                Output("Log File:" + _logPath);

            switch (operation)
            {
                case OperationType.Export2010:
                case OperationType.Export2013:
                case OperationType.Export2016:
                    try
                    {
                        var exExists = true;
                        string allProfileKey;
                        string allProfilePath;
                        string outlookPath;

                        //This mode allows us to export an Outlook 2013 profile instead of the default Outlook 2010
                        string defaultProfile;
                        switch (operation)
                        {
                            case OperationType.Export2013:
                                defaultProfile = GetDefault(default2013Profile);
                                allProfileKey = rootKey + office2013Profile;
                                allProfilePath = office2013Profile;

                                if (!string.IsNullOrEmpty(defaultProfile) && !ignoreDefault)
                                    isDefault = ProfileExists(office2013Profile + "\\" + defaultProfile);

                                if (!isDefault)
                                    outlookPath = office2013Profile + "\\" + sourceProfile;
                                else
                                    outlookPath = office2013Profile + "\\" + defaultProfile;
                                break;
                            case OperationType.Export2016:
                                defaultProfile = GetDefault(default2016Profile);
                                allProfileKey = rootKey + office2016Profile;
                                allProfilePath = office2016Profile;

                                if (!string.IsNullOrEmpty(defaultProfile) && !ignoreDefault)
                                    isDefault = ProfileExists(office2016Profile + "\\" + defaultProfile);

                                if (!isDefault)
                                    outlookPath = office2016Profile + "\\" + sourceProfile;
                                else
                                    outlookPath = office2016Profile + "\\" + defaultProfile;
                                break;
                            default:
                                defaultProfile = GetDefault(office2010ProfilePath);
                                allProfileKey = rootKey + office2010ProfilePath;
                                allProfilePath = office2010ProfilePath;

                                if (!string.IsNullOrEmpty(defaultProfile) && !ignoreDefault)
                                    isDefault = ProfileExists(office2010ProfilePath + "\\" + defaultProfile);

                                if (!isDefault)
                                    outlookPath = office2010ProfilePath + "\\" + sourceProfile;
                                else
                                    outlookPath = office2010ProfilePath + "\\" + defaultProfile;
                                break;
                        }

                        if (isDefault)
                        {
                            Output("Default Outlook profile is '" + defaultProfile + "' and will be used.");
                            sourceProfile = defaultProfile;
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(defaultProfile))
                                Output("No default Outlook profile exists");
                            else
                                Output("Default Outlook profile is '" + defaultProfile +
                                       "', source profile override: '" + sourceProfile + "'");
                        }

                        if (!isDefault && !ProfileExists(outlookPath))
                        {
                            Output("Outlook profile '" + sourceProfile +
                                   "' does not exist, no Outlook profile export will be performed.");
                            exExists = false;
                        }

                        if (exExists)
                        {
                            //Export the Outlook profile to a reg file
                            Output("Exporting Outlook profile '" + sourceProfile + "' to " + path);
                            var registryImporter = new RegistryImporter(allProfileKey, RegistryView.Default);
                            var outlookProfile = registryImporter.Import();
                            var regFileFormat5Exporter = new RegFileFormat5Exporter();
                            regFileFormat5Exporter.Export(outlookProfile, path, RegFileExportOptions.None);
                            Output("Exported Outlook profile '" + sourceProfile + "' to " + path);

                            if (useOptions)
                            {
                                if (ProfileExists(exchangeOptionsPath))
                                {
                                    Output("Exporting Client options to " + optionsPath);
                                    var optionsImporter = new RegistryImporter(rootKey + exchangeOptionsPath,
                                        RegistryView.Default);
                                    var options = optionsImporter.Import();
                                    var optionsExporter = new RegFileFormat5Exporter();
                                    optionsExporter.Export(options, optionsPath, RegFileExportOptions.None);
                                    Output("Exported Client options to " + optionsPath);
                                }
                                else
                                {
                                    Output("No client options exist");
                                }
                            }

                            string convertPath;
                            string allConvertPath;

                            switch (version)
                            {
                                case OfficeVersion.Office2016:
                                    convertPath = office2016Profile + "\\" + targetProfile;
                                    allConvertPath = office2016Profile;
                                    break;
                                default:
                                    convertPath = office2013Profile + "\\" + targetProfile;
                                    allConvertPath = office2013Profile;
                                    break;
                            }

                            if (!sourceProfile.Equals(targetProfile))
                            {
                                //Now perform a replace for a new version of Office
                                Output("Converting Outlook profile '" + sourceProfile + " to " + version +
                                       " with target profile name '" + targetProfile + "' in " + path);
                                var regFile = File.ReadAllText(path);
                                regFile = regFile.ReplaceCaseless(outlookPath + "]", convertPath + "]");
                                regFile = regFile.ReplaceCaseless(outlookPath + "\\", convertPath + "\\");
                                regFile = regFile.ReplaceCaseless(allProfilePath, allConvertPath);
                                File.WriteAllText(path, regFile);
                                Output("Updated Outlook profile '" + sourceProfile + "' to " + version +
                                       " with target profile name '" + targetProfile + "' in " + path);
                            }
                            else
                            {
                                Output("Outlook profile conversion not required for '" + targetProfile +
                                       "' - source and target profiles are the same.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Output("Error exporting Outlook profile: " + ex.Message);
                        Output("Exception details:" + ex);
                    }

                    break;
                case OperationType.Import:
                    string importPath;
                    string defaultPath;

                    switch (version)
                    {
                        case OfficeVersion.Office2016:
                            defaultPath = default2016Profile;
                            importPath = office2016Profile + "\\" + targetProfile;
                            break;
                        default:
                            defaultPath = default2013Profile;
                            importPath = office2013Profile + "\\" + targetProfile;
                            break;
                    }

                    //Check the Outlook 2013 profile exists
                    if (!File.Exists(path))
                    {
                        Output("Import file does not exist, no Outlook profile import will be performed.");
                    }
                    else
                    {
                        Output("Checking for existing Outlook profile ...");
                        var imExists = false;
                        if (!ProfileExists(importPath))
                        {
                            Output("Outlook profile '" + targetProfile +
                                   "' does not exist, Outlook profile will be imported.");
                        }
                        else
                        {
                            imExists = true;
                            Output("Outlook profile '" + targetProfile + "' already exists - will not be overwritten.");
                        }

                        //If profile didn't exist, attempt the import
                        if (!imExists)
                            try
                            {
                                Output("Importing Outlook profile '" + targetProfile + "' from " + path);
                                var importFile = RegFile.CreateImporterFromFile(path, RegFileImportOptions.None);
                                var newProfile = importFile.Import();
                                var regEnvReplace = new RegEnvReplace();

                                newProfile.WriteToTheRegistry(RegistryWriteOptions.Recursive, regEnvReplace,
                                    RegistryView.Default);
                                Output("Imported Outlook profile '" + targetProfile + "' from " + path);
                                if (SetDefault(defaultPath, targetProfile, version))
                                    Output("Set Outlook profile '" + targetProfile + "' as the default profile.");
                                else
                                    Output("Failed to set Outlook profile '" + targetProfile +
                                           "' as the default profile.");

                                if (useOptions)
                                {
                                    Output("Importing Client options from " + optionsPath);
                                    var optionsFile =
                                        RegFile.CreateImporterFromFile(optionsPath, RegFileImportOptions.None);
                                    var newOptions = optionsFile.Import();
                                    var newOptionsReplace = new RegEnvReplace();

                                    newOptions.WriteToTheRegistry(RegistryWriteOptions.Recursive, newOptionsReplace,
                                        RegistryView.Default);
                                    Output("Imported Client options from " + optionsPath);
                                }
                            }
                            catch (Exception ex)
                            {
                                Output("Error importing Outlook profile: " + ex.Message);
                                Output("Exception details:" + ex);
                            }
                    }

                    break;
                default:
                    Console.WriteLine();
                    Console.WriteLine(
                        "USAGE: OutlookProfiler Export2010={FilePath} [Options={OptionsFilePath}] [TargetProfile ={ProfileName}] [SourceProfile={ProfileName}] [TargetVersion=2013|2016] [Log={LogPath}] [IgnoreDefault]");
                    Console.WriteLine(
                        "       OutlookProfiler Export2013 ={FilePath} [Options={OptionsFilePath}] [TargetProfile ={ProfileName}] [SourceProfile={ProfileName}] [TargetVersion=2013|2016] [Log={LogPath}] [IgnoreDefault]");
                    Console.WriteLine(
                        "       OutlookProfiler Export2016={FilePath} [Options={OptionsFilePath}] [TargetProfile ={ProfileName}] [SourceProfile={ProfileName}] [TargetVersion=2013|2016] [Log={LogPath}] [IgnoreDefault]");
                    Console.WriteLine(
                        "       OutlookProfiler Import={FilePath} [Options={OptionsFilePath}] [TargetProfile ={ProfileName}] [TargetVersion=2013|2016] [Log={LogPath}]");
                    Console.WriteLine();
                    Console.WriteLine(
                        "NOTE: For export operations, you can use IgnoreDefault as an optional parameter to force the SourceProfile to be used instead of the Default Profile.");
                    break;
            }
        }

        private enum OperationType
        {
            Export2010,
            Export2013,
            Export2016,
            Import,
            None = -1
        }

        private enum OfficeVersion
        {
            Office2013,
            Office2016
        }
    }
}