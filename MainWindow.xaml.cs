using System;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.Win32;
using System.Management;

namespace RemoveMilestoneManagementFailover
{
    public partial class MainWindow : Window
    {
        private readonly string _logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt");

        public MainWindow()
        {
            InitializeComponent();
            WriteLog("MainWindow Constructor", "Initialized successfully.");
        }

        private void WriteLog(string functionName, string message)
        {
            try
            {
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {functionName} | {message}\n";
                File.AppendAllText(_logFilePath, logEntry);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }

        public string UninstallXProtectManagementServerFailover()
        {
            try
            {
                string uninstallString = null;
                string registryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";

                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryPath))
                {
                    if (key != null)
                    {
                        foreach (string subkeyName in key.GetSubKeyNames())
                        {
                            using (RegistryKey subkey = key.OpenSubKey(subkeyName))
                            {
                                if (subkey != null)
                                {
                                    string displayName = subkey.GetValue("DisplayName")?.ToString();
                                    if (displayName == "XProtect Management Server Failover")
                                    {
                                        uninstallString = subkey.GetValue("UninstallString")?.ToString();
                                        WriteLog("UninstallXProtectManagementServerFailover", $"Found uninstall string for {displayName}: {uninstallString}");
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        WriteLog("UninstallXProtectManagementServerFailover", "Registry key not found.");
                    }
                }

                if (string.IsNullOrEmpty(uninstallString))
                {
                    WriteLog("UninstallXProtectManagementServerFailover", "Failure: XProtect Management Server Failover not found in registry.");
                    return "Error: XProtect Management Server Failover not found in registry.";
                }

                string modifiedUninstallString = uninstallString.Replace("/I", "/X");
                WriteLog("UninstallXProtectManagementServerFailover", $"Modified uninstall string: {modifiedUninstallString}");

                ProcessStartInfo processInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/C {modifiedUninstallString}",
                    UseShellExecute = true,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(processInfo))
                {
                    process.WaitForExit();
                    if (process.ExitCode == 0)
                    {
                        WriteLog("UninstallXProtectManagementServerFailover", "Success: Uninstallation completed.");
                        return "Uninstallation completed successfully.";
                    }
                    else
                    {
                        WriteLog("UninstallXProtectManagementServerFailover", $"Failure: Uninstallation failed with exit code {process.ExitCode}.");
                        return $"Uninstallation failed with exit code: {process.ExitCode}";
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("UninstallXProtectManagementServerFailover", $"Failure: Error during uninstallation: {ex.Message}");
                return $"Error during uninstallation: {ex.Message}";
            }
        }

        public string DeleteXProtectFailoverFolder()
        {
            try
            {
                string installPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                string milestonePath = Path.Combine(installPath, "Milestone");
                string targetFolder = Path.Combine(milestonePath, "XProtect Management Server Failover");
                WriteLog("DeleteXProtectFailoverFolder", $"Checking for folder: {targetFolder}");

                if (Directory.Exists(targetFolder))
                {
                    Directory.Delete(targetFolder, true);
                    WriteLog("DeleteXProtectFailoverFolder", $"Success: Deleted folder {targetFolder}.");
                    return "XProtect Management Server Failover folder deleted successfully.";
                }
                else
                {
                    WriteLog("DeleteXProtectFailoverFolder", $"Failure: Folder {targetFolder} not found.");
                    return "Error: XProtect Management Server Failover folder not found.";
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                WriteLog("DeleteXProtectFailoverFolder", $"Failure: Insufficient permissions to delete folder: {ex.Message}");
                return "Error: Insufficient permissions to delete the folder. Try running as administrator.";
            }
            catch (IOException ex)
            {
                WriteLog("DeleteXProtectFailoverFolder", $"Failure: IO error deleting folder: {ex.Message}");
                return $"Error: Failed to delete folder due to IO issue: {ex.Message}";
            }
            catch (Exception ex)
            {
                WriteLog("DeleteXProtectFailoverFolder", $"Failure: Unexpected error: {ex.Message}");
                return $"Error: An unexpected error occurred: {ex.Message}";
            }
        }

        public string DeleteFailoverWizardFile()
        {
            string targetPath = @"C:\ProgramData\Milestone\XProtect Management Server";
            string targetFile = Path.Combine(targetPath, "failoverwizard.json");

            try
            {
                WriteLog("DeleteFailoverWizardFile", $"Checking for file: {targetFile}");

                if (File.Exists(targetFile))
                {
                    File.Delete(targetFile);
                    WriteLog("DeleteFailoverWizardFile", $"Success: Deleted file {targetFile}.");
                    return "failoverwizard.json deleted successfully.";
                }
                else
                {
                    WriteLog("DeleteFailoverWizardFile", $"Failure: File {targetFile} not found.");
                    return "Error: failoverwizard.json not found.";
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                WriteLog("DeleteFailoverWizardFile", $"Failure: Insufficient permissions to delete file {targetFile}: {ex.Message}");
                return "Error: Insufficient permissions to delete the file. Try running as administrator.";
            }
            catch (IOException ex)
            {
                WriteLog("DeleteFailoverWizardFile", $"Failure: IO error deleting file {targetFile}: {ex.Message}");
                return $"Error: Failed to delete file due to IO issue: {ex.Message}";
            }
            catch (Exception ex)
            {
                WriteLog("DeleteFailoverWizardFile", $"Failure: Unexpected error: {ex.Message}");
                return $"Error: An unexpected error occurred: {ex.Message}";
            }
        }

        public string RemoveUninstallRegistryKey()
        {
            try
            {
                string registryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{2FA15F94-4EFE-4883-A8B5-D2B0B4804EE0}";
                WriteLog("RemoveUninstallRegistryKey", $"Attempting to remove registry key: {registryPath}");

                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryPath, writable: true))
                {
                    if (key != null)
                    {
                        key.Close(); // Ensure it's closed before deleting
                    }
                }

                Registry.LocalMachine.DeleteSubKey(registryPath, throwOnMissingSubKey: false);
                WriteLog("RemoveUninstallRegistryKey", "Success: Registry key removed.");
                return "Uninstall registry key removed successfully.";
            }
            catch (UnauthorizedAccessException ex)
            {
                WriteLog("RemoveUninstallRegistryKey", $"Failure: Insufficient permissions to delete registry key: {ex.Message}");
                return "Error: Insufficient permissions to delete the registry key. Try running as administrator.";
            }
            catch (Exception ex)
            {
                WriteLog("RemoveUninstallRegistryKey", $"Failure: Unexpected error: {ex.Message}");
                return $"Error: An unexpected error occurred: {ex.Message}";
            }
        }

        private string ManageWebPublishingService()
        {
            try
            {
                string serviceName = "W3SVC";
                WriteLog("ManageWebPublishingService", $"Checking service: {serviceName}");

                using (ServiceController sc = new ServiceController(serviceName))
                {
                    ServiceController[] services = ServiceController.GetServices();
                    var service = Array.Find(services, s => s.ServiceName == serviceName);

                    if (service != null && service.StartType == ServiceStartMode.Manual)
                    {
                        WriteLog("ManageWebPublishingService", $"{serviceName} is set to Manual. Attempting to change to Automatic.");
                        MessageBox.Show($"{serviceName} is set to Manual. Changing to Automatic...");
                        WriteLog("ManageWebPublishingService", $"{serviceName} start mode changed to Automatic (Note: Admin rights required).");
                    }

                    if (sc.Status == ServiceControllerStatus.Stopped)
                    {
                        WriteLog("ManageWebPublishingService", $"{serviceName} is stopped. Starting service...");
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                        WriteLog("ManageWebPublishingService", $"Success: {serviceName} started.");
                        return $"{serviceName} has been started.";
                    }
                    else
                    {
                        WriteLog("ManageWebPublishingService", $"Success: {serviceName} is already running.");
                        return $"{serviceName} is already running.";
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("ManageWebPublishingService", $"Failure: Error managing service: {ex.Message}");
                return $"An error occurred: {ex.Message}";
            }
        }

        public static (string Hostname, string InstanceName) GetSqlServerDetails(MainWindow logger)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\VideoOS\Server\ConnectionString"))
                {
                    if (key == null)
                    {
                        logger.WriteLog("GetSqlServerDetails", "Failure: Registry key not found.");
                        throw new Exception("Registry key not found.");
                    }

                    string connectionString = key.GetValue("ManagementServer") as string;
                    logger.WriteLog("GetSqlServerDetails", $"Retrieved connection string: {connectionString}");

                    if (string.IsNullOrEmpty(connectionString))
                    {
                        logger.WriteLog("GetSqlServerDetails", "Failure: ManagementServer value not found or is empty.");
                        throw new Exception("ManagementServer value not found or is empty.");
                    }

                    string[] parts = connectionString.Split(';');
                    string dataSource = null;
                    foreach (string part in parts)
                    {
                        if (part.Trim().StartsWith("Data Source=", StringComparison.OrdinalIgnoreCase))
                        {
                            dataSource = part.Replace("Data Source=", "").Trim();
                            logger.WriteLog("GetSqlServerDetails", $"Found Data Source: {dataSource}");
                            break;
                        }
                    }

                    if (string.IsNullOrEmpty(dataSource))
                    {
                        logger.WriteLog("GetSqlServerDetails", "Failure: Data Source not found in connection string.");
                        throw new Exception("Data Source not found in connection string.");
                    }

                    string hostname = dataSource;
                    string instanceName = null;

                    int instanceIndex = dataSource.IndexOf('\\');
                    if (instanceIndex >= 0)
                    {
                        hostname = dataSource.Substring(0, instanceIndex);
                        instanceName = dataSource.Substring(instanceIndex + 1);
                    }

                    logger.WriteLog("GetSqlServerDetails", $"Success: Hostname={hostname}, InstanceName={instanceName}");
                    return (hostname, instanceName);
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog("GetSqlServerDetails", $"Failure: Error: {ex.Message}");
                Console.WriteLine($"Error: {ex.Message}");
                return (null, null);
            }
        }

        public static string ManageSqlServerService(MainWindow logger)
        {
            try
            {
                var (hostname, instanceName) = GetSqlServerDetails(logger);
                if (hostname == null)
                {
                    logger.WriteLog("ManageSqlServerService", "Failure: Could not retrieve SQL Server details.");
                    return "Error: Could not retrieve SQL Server details.";
                }

                string serviceName = string.IsNullOrEmpty(instanceName) ? "MSSQLSERVER" : $"MSSQL${instanceName}";
                logger.WriteLog("ManageSqlServerService", $"Managing service: {serviceName}");

                using (ServiceController sc = new ServiceController(serviceName))
                {
                    ServiceController[] services = ServiceController.GetServices();
                    var service = Array.Find(services, s => s.ServiceName == serviceName);

                    if (service != null)
                    {
                        // Check if service is Disabled and change to Automatic
                        if (service.StartType == ServiceStartMode.Disabled)
                        {
                            logger.WriteLog("ManageSqlServerService", $"{serviceName} is set to Disabled. Attempting to change to Automatic.");
                            using (var managementObject = new ManagementObject($"Win32_Service.Name='{serviceName}'"))
                            {
                                managementObject.InvokeMethod("ChangeStartMode", new object[] { "Automatic" });
                            }
                            logger.WriteLog("ManageSqlServerService", $"{serviceName} start mode changed to Automatic.");
                        }

                        // Check if service is Manual and change to Automatic
                        if (service.StartType == ServiceStartMode.Manual)
                        {
                            logger.WriteLog("ManageSqlServerService", $"{serviceName} is set to Manual. Attempting to change to Automatic.");
                            using (var managementObject = new ManagementObject($"Win32_Service.Name='{serviceName}'"))
                            {
                                managementObject.InvokeMethod("ChangeStartMode", new object[] { "Automatic" });
                            }
                            logger.WriteLog("ManageSqlServerService", $"{serviceName} start mode changed to Automatic.");
                        }

                        // Check if service is stopped and start it
                        if (sc.Status == ServiceControllerStatus.Stopped)
                        {
                            logger.WriteLog("ManageSqlServerService", $"{serviceName} is stopped. Starting service...");
                            sc.Start();
                            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                            logger.WriteLog("ManageSqlServerService", $"Success: {serviceName} started.");
                            return $"{serviceName} was stopped and has been started.";
                        }
                        else
                        {
                            logger.WriteLog("ManageSqlServerService", $"Success: {serviceName} is already running.");
                            return $"{serviceName} is already running.";
                        }
                    }
                    else
                    {
                        logger.WriteLog("ManageSqlServerService", $"Failure: Service {serviceName} not found.");
                        return $"Error: Service {serviceName} not found.";
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog("ManageSqlServerService", $"Failure: Error: {ex.Message}");
                return $"Error: {ex.Message}";
            }
        }

        public static string ManageMilestoneServices(MainWindow logger)
        {
            try
            {
                ServiceController[] services = ServiceController.GetServices();
                string result = "No matching services found.";
                logger.WriteLog("ManageMilestoneServices", "Starting to process Milestone services.");

                foreach (ServiceController service in services)
                {
                    if (service.ServiceName.StartsWith("Milestone", StringComparison.OrdinalIgnoreCase))
                    {
                        logger.WriteLog("ManageMilestoneServices", $"Processing service: {service.ServiceName}");
                        result = $"Processing service: {service.ServiceName}\n";

                        if (service.StartType == ServiceStartMode.Manual)
                        {
                            logger.WriteLog("ManageMilestoneServices", $"{service.ServiceName} is set to Manual. Changing to Automatic.");
                            result += $"{service.ServiceName} is set to Manual. Changing to Automatic...\n";
                            using (var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_Service WHERE Name = '{service.ServiceName}'"))
                            {
                                foreach (ManagementObject obj in searcher.Get())
                                {
                                    obj.InvokeMethod("ChangeStartMode", new object[] { "Automatic" });
                                }
                            }
                            logger.WriteLog("ManageMilestoneServices", $"Success: {service.ServiceName} start mode changed to Automatic.");
                            result += $"{service.ServiceName} start mode changed to Automatic.\n";
                        }

                        if (service.Status == ServiceControllerStatus.Stopped)
                        {
                            logger.WriteLog("ManageMilestoneServices", $"{service.ServiceName} is stopped. Starting service...");
                            result += $"{service.ServiceName} is stopped. Starting the service...\n";
                            service.Start();
                            service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                            logger.WriteLog("ManageMilestoneServices", $"Success: {service.ServiceName} started.");
                            result += $"{service.ServiceName} has been started.\n";
                        }
                        else
                        {
                            logger.WriteLog("ManageMilestoneServices", $"Success: {service.ServiceName} is already running.");
                            result += $"{service.ServiceName} is already running.\n";
                        }
                    }
                }

                logger.WriteLog("ManageMilestoneServices", $"Success: Completed processing services. Result: {result.Trim()}");
                return result.Trim();
            }
            catch (Exception ex)
            {
                logger.WriteLog("ManageMilestoneServices", $"Failure: Error: {ex.Message}");
                return $"Error: {ex.Message}";
            }
        }

        public static string CheckAndSetW3SVCStartMode(MainWindow logger)
        {
            try
            {
                string serviceName = "W3SVC";
                logger.WriteLog("CheckAndSetW3SVCStartMode", $"Checking service: {serviceName}");

                using (ServiceController sc = new ServiceController(serviceName))
                {
                    ServiceController[] services = ServiceController.GetServices();
                    var service = Array.Find(services, s => s.ServiceName == serviceName);

                    if (service != null)
                    {
                        // Check if service is set to Manual
                        if (service.StartType == ServiceStartMode.Manual)
                        {
                            logger.WriteLog("CheckAndSetW3SVCStartMode", $"{serviceName} is set to Manual. Attempting to change to Automatic.");
                            using (var managementObject = new ManagementObject($"Win32_Service.Name='{serviceName}'"))
                            {
                                managementObject.InvokeMethod("ChangeStartMode", new object[] { "Automatic" });
                            }
                            logger.WriteLog("CheckAndSetW3SVCStartMode", $"{serviceName} start mode changed to Automatic.");
                            return $"{serviceName} start mode changed to Automatic.";
                        }
                        else
                        {
                            logger.WriteLog("CheckAndSetW3SVCStartMode", $"{serviceName} is already set to {service.StartType}.");
                            return $"{serviceName} is already set to {service.StartType}.";
                        }
                    }
                    else
                    {
                        logger.WriteLog("CheckAndSetW3SVCStartMode", $"Failure: Service {serviceName} not found.");
                        return $"Error: Service {serviceName} not found.";
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog("CheckAndSetW3SVCStartMode", $"Failure: Error: {ex.Message}");
                return $"Error: {ex.Message}";
            }
        }

        public void RegisterServer(string newHostName = null)
        {
            try
            {
                string fullPath = @"C:\Program Files\Milestone\Server Configurator\ServerConfigurator.exe";

                if (!File.Exists(fullPath))
                {
                    throw new FileNotFoundException("The specified executable could not be found.", fullPath);
                }

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = fullPath,
                    WorkingDirectory = System.IO.Path.GetDirectoryName(fullPath),
                    Verb = "runas",
                    Arguments = "/register"
                };

                if (newHostName != null)
                {
                    psi.Arguments += " /managementserveraddress=" + newHostName;
                }

                var process = Process.Start(psi);
                while (!process.HasExited)
                {
                    Thread.Sleep(2000);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    //AppendMessageToRichTextBox(ex.Message);
                });
            }

        }

        public static string CheckAndSetMilestoneXProtectManagementServerStartMode(MainWindow logger)
        {
            try
            {
                string serviceName = "MilestoneXProtectManagementServer"; // Assumed service name
                logger.WriteLog("CheckAndSetMilestoneXProtectManagementServerStartMode", $"Checking service: {serviceName}");

                using (ServiceController sc = new ServiceController(serviceName))
                {
                    ServiceController[] services = ServiceController.GetServices();
                    var service = Array.Find(services, s => s.ServiceName == serviceName);

                    if (service != null)
                    {
                        // Check if service is set to Manual
                        if (service.StartType == ServiceStartMode.Manual)
                        {
                            logger.WriteLog("CheckAndSetMilestoneXProtectManagementServerStartMode", $"{serviceName} is set to Manual. Attempting to change to Automatic.");
                            using (var managementObject = new ManagementObject($"Win32_Service.Name='{serviceName}'"))
                            {
                                managementObject.InvokeMethod("ChangeStartMode", new object[] { "Automatic" });
                            }
                            logger.WriteLog("CheckAndSetMilestoneXProtectManagementServerStartMode", $"{serviceName} start mode changed to Automatic.");
                            return $"{serviceName} start mode changed to Automatic.";
                        }
                        else
                        {
                            logger.WriteLog("CheckAndSetMilestoneXProtectManagementServerStartMode", $"{serviceName} is already set to {service.StartType}.");
                            return $"{serviceName} is already set to {service.StartType}.";
                        }
                    }
                    else
                    {
                        logger.WriteLog("CheckAndSetMilestoneXProtectManagementServerStartMode", $"Failure: Service {serviceName} not found.");
                        return $"Error: Service {serviceName} not found.";
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog("CheckAndSetMilestoneXProtectManagementServerStartMode", $"Failure: Error: {ex.Message}");
                return $"Error: {ex.Message}";
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            WriteLog("Button_Click", "Button clicked, starting operations.");

            string uninstallResult = UninstallXProtectManagementServerFailover();
            WriteLog("Button_Click", $"UninstallXProtectManagementServerFailover result: {uninstallResult}");

            string folderDeleteResult = DeleteXProtectFailoverFolder();
            WriteLog("Button_Click", $"DeleteXProtectFailoverFolder result: {folderDeleteResult}");

            string fileDeleteResult = DeleteFailoverWizardFile();
            WriteLog("Button_Click", $"DeleteFailoverWizardFile result: {fileDeleteResult}");

            string registryDeleteResult = RemoveUninstallRegistryKey();
            WriteLog("Button_Click", $"RemoveUninstallRegistryKey result: {registryDeleteResult}");

            Paragraph paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run($"{uninstallResult}\n{folderDeleteResult}\n{fileDeleteResult}\n{registryDeleteResult}"));
            ReportBox.Document.Blocks.Add(paragraph);

            string startWebService = ManageWebPublishingService();
            WriteLog("Button_Click", $"ManageWebPublishingService result: {startWebService}");

            string startSqlService = ManageSqlServerService(this);
            WriteLog("Button_Click", $"ManageSqlServerService result: {startSqlService}");

            string startMilestoneServices = ManageMilestoneServices(this);
            WriteLog("Button_Click", $"ManageMilestoneServices result: {ManageMilestoneServices}");

            Paragraph serviceParagraph = new Paragraph();
            serviceParagraph.Inlines.Add(new Run($"{startSqlService}\n{startMilestoneServices}"));
            ReportBox.Document.Blocks.Add(serviceParagraph);

            string checkW3SVC = CheckAndSetW3SVCStartMode(this);
            serviceParagraph.Inlines.Add(new Run($"\n{checkW3SVC}"));
            ReportBox.Document.Blocks.Add(serviceParagraph);
            WriteLog("starting www", $"CheckAndSetW3SVCStartMode result: {checkW3SVC}");

            string checkMilestoneService = CheckAndSetMilestoneXProtectManagementServerStartMode(this);
            serviceParagraph.Inlines.Add(new Run($"\n{checkMilestoneService}"));
            ReportBox.Document.Blocks.Add(serviceParagraph);
            WriteLog("starting MilestoneXProtectManagementServer", $"CheckAndSetMilestoneXProtectManagementServerStartMode result: {checkMilestoneService}");

            Paragraph registerParagraph = new Paragraph();
            registerParagraph.Inlines.Add(new Run($"Re-registering the server..."));
            ReportBox.Document.Blocks.Add(registerParagraph);
            RegisterServer();
        }
    }
}