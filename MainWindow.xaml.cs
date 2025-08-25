using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Diagnostics;
using System;
using System.ServiceProcess;
using System.IO;
using Path = System.IO.Path;
using System.Management;


namespace RemoveMilestoneManagementFailover
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
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
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                if (string.IsNullOrEmpty(uninstallString))
                {
                    return "Error: XProtect Management Server Failover not found in registry.";
                }

                // Replace /I with /X in the uninstall string
                string modifiedUninstallString = uninstallString.Replace("/I", "/X");

                // Execute the uninstall command
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
                        return "Uninstallation completed successfully.";
                    }
                    else
                    {
                        return $"Uninstallation failed with exit code: {process.ExitCode}";
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Error during uninstallation: {ex.Message}";
            }
        }


        public string DeleteXProtectFailoverFolder()
        {
            try
            {
                // Get the install path (e.g., Program Files) from environment variable
                string installPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                string milestonePath = Path.Combine(installPath, "Milestone");
                string targetFolder = Path.Combine(milestonePath, "XProtect Management Server Failover");

                // Check if the folder exists
                if (Directory.Exists(targetFolder))
                {
                    // Delete the folder and all its contents
                    Directory.Delete(targetFolder, true);
                    return "XProtect Management Server Failover folder deleted successfully.";
                }
                else
                {
                    return "Error: XProtect Management Server Failover folder not found.";
                }
            }
            catch (UnauthorizedAccessException)
            {
                return "Error: Insufficient permissions to delete the folder. Try running as administrator.";
            }
            catch (IOException ex)
            {
                return $"Error: Failed to delete folder due to IO issue: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"Error: An unexpected error occurred: {ex.Message}";
            }

        }

        public string DeleteFailoverWizardFile()
        {
            try
            {
                string targetPath = @"C:\ProgramData\Milestone\XProtect Management Server";
                string targetFile = Path.Combine(targetPath, "failoverwizard.json");

                // Check if the file exists
                if (File.Exists(targetFile))
                {
                    // Delete the file
                    File.Delete(targetFile);
                    return "failoverwizard.json deleted successfully.";
                }
                else
                {
                    return "Error: failoverwizard.json not found.";
                }
            }
            catch (UnauthorizedAccessException)
            {
                return "Error: Insufficient permissions to delete the file. Try running as administrator.";
            }
            catch (IOException ex)
            {
                return $"Error: Failed to delete file due to IO issue: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"Error: An unexpected error occurred: {ex.Message}";
            }
        }



        private string ManageWebPublishingService()
        {
            try
            {
                // Define the service name
                string serviceName = "W3SVC";

                // Get the service controller
                using (ServiceController sc = new ServiceController(serviceName))
                {
                    // Check if the service is set to manual
                    ServiceController[] services = ServiceController.GetServices();
                    var service = Array.Find(services, s => s.ServiceName == serviceName);

                    if (service != null && service.StartType == ServiceStartMode.Manual)
                    {
                        MessageBox.Show($"{serviceName} is set to Manual. Changing to Automatic...");
                        // Change start mode to Automatic (requires admin privileges)
                        // Note: Changing start mode directly requires WMI or registry access; this is a simplified check
                        // For full control, use WMI as shown in comments below
                        MessageBox.Show($"{serviceName} start mode changed to Automatic (Note: Admin rights required for actual change).");
                    }

                    // Check if the service is stopped
                    if (sc.Status == ServiceControllerStatus.Stopped)
                    {
                       return $"{serviceName} is stopped. Starting the service...";
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                        MessageBox.Show($"{serviceName} has been started.");
                    }
                    else
                    {
                        return $"{serviceName} is already running.";
                    }
                }
            }
            catch (Exception ex)
            {
                return $"An error occurred: {ex.Message}";
            }
        }

        public static (string Hostname, string InstanceName) GetSqlServerDetails()
        {
            try
            {
                // Open the registry key
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\VideoOS\Server\ConnectionString"))
                {
                    if (key == null)
                    {
                        throw new Exception("Registry key not found.");
                    }

                    // Read the ManagementServer value
                    string connectionString = key.GetValue("ManagementServer") as string;

                    if (string.IsNullOrEmpty(connectionString))
                    {
                        throw new Exception("ManagementServer value not found or is empty.");
                    }

                    // Parse the connection string to get the Data Source
                    string[] parts = connectionString.Split(';');
                    string dataSource = null;
                    foreach (string part in parts)
                    {
                        if (part.Trim().StartsWith("Data Source=", StringComparison.OrdinalIgnoreCase))
                        {
                            dataSource = part.Replace("Data Source=", "").Trim();
                            break;
                        }
                    }

                    if (string.IsNullOrEmpty(dataSource))
                    {
                        throw new Exception("Data Source not found in connection string.");
                    }

                    // Split Data Source into hostname and instance name
                    string hostname = dataSource;
                    string instanceName = null;

                    int instanceIndex = dataSource.IndexOf('\\');
                    if (instanceIndex >= 0)
                    {
                        hostname = dataSource.Substring(0, instanceIndex);
                        instanceName = dataSource.Substring(instanceIndex + 1);
                    }

                    return (hostname, instanceName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return (null, null);
            }
        }

        public static string ManageSqlServerService()
        {
            try
            {
                var (hostname, instanceName) = GetSqlServerDetails();
                if (hostname == null) return $"";

                // Determine the service name based on instance
                string serviceName = string.IsNullOrEmpty(instanceName) ? "MSSQLSERVER" : $"MSSQL${instanceName}";

                using (ServiceController sc = new ServiceController(serviceName))
                {
                    ServiceController[] services = ServiceController.GetServices();
                    var service = Array.Find(services, s => s.ServiceName == serviceName);

                    if (service != null && service.StartType == ServiceStartMode.Manual)
                    {
                        return $"{serviceName} is set to Manual. Changing to Automatic...";
                        // Note: Changing StartType requires WMI or admin rights; this is a placeholder
                        // For actual change, use WMI as shown in comments below
                        return $"{serviceName} start mode changed to Automatic (Admin rights required).";
                    }

                    if (sc.Status == ServiceControllerStatus.Stopped)
                    {
                        return $"{serviceName} is stopped. Starting the service...";
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                        return $"{serviceName} has been started.";
                    }
                    else
                    {
                        return $"{serviceName} is already running.";
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }



        public static string ManageMilestoneServices()
        {
            try
            {
                // Get all services
                ServiceController[] services = ServiceController.GetServices();
                string result = "No matching services found.";

                foreach (ServiceController service in services)
                {
                    // Check if the service name begins with "Milestone"
                    if (service.ServiceName.StartsWith("Milestone", StringComparison.OrdinalIgnoreCase))
                    {
                        result = $"Processing service: {service.ServiceName}\n";

                        // Check if the service is set to Manual
                        if (service.StartType == ServiceStartMode.Manual)
                        {
                            result += $"{service.ServiceName} is set to Manual. Changing to Automatic...\n";
                            // Use WMI to change the start mode (requires admin privileges)
                            using (var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_Service WHERE Name = '{service.ServiceName}'"))
                            {
                                foreach (ManagementObject obj in searcher.Get())
                                {
                                    obj.InvokeMethod("ChangeStartMode", new object[] { "Automatic" });
                                }
                            }
                            result += $"{service.ServiceName} start mode changed to Automatic.\n";
                        }

                        // Check if the service is stopped and start it if necessary
                        if (service.Status == ServiceControllerStatus.Stopped)
                        {
                            result += $"{service.ServiceName} is stopped. Starting the service...\n";
                            service.Start();
                            service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                            result += $"{service.ServiceName} has been started.\n";
                        }
                        else
                        {
                            result += $"{service.ServiceName} is already running.\n";
                        }
                    }
                }

                return result.Trim(); // Return the accumulated result, trimming trailing newline
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Run the uninstall process
            string uninstallResult = UninstallXProtectManagementServerFailover();

            // Delete the folder
            string folderDeleteResult = DeleteXProtectFailoverFolder();

            // Delete the failoverwizard.json file
            string fileDeleteResult = DeleteFailoverWizardFile();
            Paragraph paragraph = new Paragraph();
            paragraph.Inlines.Add(new Run($"{uninstallResult}\n{folderDeleteResult}\n{fileDeleteResult}"));
            ReportBox.Document.Blocks.Add(paragraph);
            string startWebServece = ManageWebPublishingService();
            string startSqlServic =  ManageSqlServerService();

        }
    }
}