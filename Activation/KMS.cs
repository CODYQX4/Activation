using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms.VisualStyles;
using Activation.Properties;
using Common;
using License;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using ProductDetection;

namespace Activation
{
    /// <summary>
    /// Methods for setting KMS Connection parameters
    /// </summary>
    public static class KMSConnection
    {
        /// <summary>
        /// Remove Microsoft Office KMS Hostname
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        public static void RemoveKMSHostOffice(LicenseList licenses)
        {
            // Call Windows if Microsoft Office uses Windows Licensing Services
            if (OfficeVersion.IsOfficeSPP())
            {
                RemoveKMSHostWMI(licenses, "SoftwareLicensingService.Version=" + "'" + OSVersion.GetSPPSVCVersion() + "'", "SoftwareLicensingProduct.ID=");
            }
            else
            {
                RemoveKMSHostWMI(licenses, "OfficeSoftwareProtectionService.Version=" + "'" + OfficeVersion.GetOSPPSVCVersion() + "'", "OfficeSoftwareProtectionProduct.ID=");
            }
        }

        /// <summary>
        /// Remove Microsoft Windows KMS Hostname
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        public static void RemoveKMSHostWindows(LicenseList licenses)
        {
            // Windows Vista
            if (Math.Abs(OSVersion.GetWindowsNumber() - 6.0) < Double.Epsilon)
            {
                RemoveKMSHostRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL");
                return;
            }
            // Windows 7 and Later
            RemoveKMSHostRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform");
        }

        /// <summary>
        /// Remove KMS Hostname using the Registry
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="registryPath">Registry Key where KeyManagementServiceName exists</param>
        private static void RemoveKMSHostRegistry(LicenseList licenses, string registryPath)
        {
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath, true))
            {
                if (registry != null)
                {
                    // Delete Global KMS Hostname
                    registry.DeleteValue("KeyManagementServiceName");

                    // Only Search KMS Licenses
                    licenses.List = licenses.GetListKMS();
                    foreach (LicenseInstance license in licenses.GetListLicensed())
                    {
                        // Application Specific
                        using (RegistryKey registrySubKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID, true))
                        {
                            if (registrySubKey != null && registrySubKey.GetValue("KeyManagementServiceName") != null)
                            {
                                // Delete Application Specific KMS Hostname
                                registrySubKey.DeleteValue("KeyManagementServiceName");

                                // SKU Specific
                                using (RegistryKey registrySubKey2 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID + Path.DirectorySeparatorChar + license.SKUID, true))
                                {
                                    if (registrySubKey2 != null && registrySubKey2.GetValue("KeyManagementServiceName") != null)
                                    {
                                        // Delete SKU Specific KMS Hostname
                                        registrySubKey2.DeleteValue("KeyManagementServiceName");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    throw new ApplicationException(@"HKLM\" + registryPath + " does not exist!");
                }
            }
        }

        /// <summary>
        /// Remove KMS Hostname using WMI
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="wmiInfo1">WMI Provider and associated data to remove Global KMS Hostname</param>
        /// <param name="wmiInfo2">WMI Provider and associated data to remove Specific KMS Hostnames</param>
        private static void RemoveKMSHostWMI(LicenseList licenses, string wmiInfo1, string wmiInfo2)
        {
            // Get Global KMS Hostname WMI Object
            ManagementObject classInstance = new ManagementObject("root\\CIMV2", wmiInfo1, null);

            // Execute Method for Global KMS Hostname
            classInstance.InvokeMethod("ClearKeyManagementServiceMachine", null, null);

            // Remove Specific KMS Hostnames
            foreach (LicenseInstance license in licenses.GetListLicensed())
            {
                // Get Specific KMS Hostname WMI Object
                classInstance = new ManagementObject("root\\CIMV2", wmiInfo2 + "'" + license.SKUID + "'", null);

                // Execute Method for Specific KMS Hostname
                classInstance.InvokeMethod("ClearKeyManagementServiceMachine", null, null);
            }
        }

        /// <summary>
        /// Set Microsoft Office KMS Hostname
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="kmsServer">KMS Hostname</param>
        public static void SetKMSHostOffice(LicenseList licenses, string kmsServer)
        {
            // Call Windows if Microsoft Office uses Windows Licensing Services
            if (OfficeVersion.IsOfficeSPP())
            {
                SetKMSHostWMI(licenses, kmsServer, "SoftwareLicensingService.Version=" + "'" + OSVersion.GetSPPSVCVersion() + "'", "SoftwareLicensingProduct.ID=");
            }
            else
            {
                SetKMSHostWMI(licenses, kmsServer, "OfficeSoftwareProtectionService.Version=" + "'" + OfficeVersion.GetOSPPSVCVersion() + "'", "OfficeSoftwareProtectionProduct.ID=");
            }
        }

        /// <summary>
        /// Set Microsoft Windows KMS Hostname
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="kmsServer">KMS Hostname</param>
        public static void SetKMSHostWindows(LicenseList licenses, string kmsServer)
        {
            // Windows Vista
            if (Math.Abs(OSVersion.GetWindowsNumber() - 6.0) < Double.Epsilon)
            {
                SetKMSHostRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL", kmsServer);
                return;
            }
            // Windows 7 and Later
            SetKMSHostRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform", kmsServer);
        }

        /// <summary>
        /// Set KMS Hostname using the Registry
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="registryPath">Registry Key where KeyManagementServiceName belongs</param>
        /// <param name="kmsServer">KMS Hostname</param>
        private static void SetKMSHostRegistry(LicenseList licenses, string registryPath, string kmsServer)
        {
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath, true))
            {
                if (registry != null)
                {
                    // Set Global KMS Hostname
                    registry.SetValue("KeyManagementServiceName", kmsServer);

                    // Only Search KMS Licenses
                    licenses.List = licenses.GetListKMS();
                    foreach (LicenseInstance license in licenses.GetListLicensed())
                    {
                        // Application Specific
                        using (RegistryKey registrySubKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID, true))
                        {
                            if (registrySubKey != null && registrySubKey.GetValue("KeyManagementServiceName") != null)
                            {
                                // Set Application Specific KMS Hostname
                                registrySubKey.SetValue("KeyManagementServiceName", kmsServer);

                                // SKU Specific
                                using (RegistryKey registrySubKey2 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID + Path.DirectorySeparatorChar + license.SKUID, true))
                                {
                                    if (registrySubKey2 != null && registrySubKey2.GetValue("KeyManagementServiceName") != null)
                                    {
                                        // Set SKU Specific KMS Hostname
                                        registrySubKey2.SetValue("KeyManagementServiceName", kmsServer);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    throw new ApplicationException(@"HKLM\" + registryPath + " does not exist!");
                }
            }
        }

        /// <summary>
        /// Set KMS Hostname using WMI
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="kmsServer">KMS Hostname to set</param>
        /// <param name="wmiInfo1">WMI Provider and associated data to set Global KMS Hostname</param>
        /// <param name="wmiInfo2">WMI Provider and associated data to set Specific KMS Hostnames</param>
        private static void SetKMSHostWMI(LicenseList licenses, string kmsServer, string wmiInfo1, string wmiInfo2)
        {
            // Get Global KMS Hostname WMI Object
            ManagementObject classInstance = new ManagementObject("root\\CIMV2", wmiInfo1, null);

            // Obtain and Add Input Parameters
            ManagementBaseObject inParams = classInstance.GetMethodParameters("SetKeyManagementServiceMachine");
            inParams["MachineName"] = kmsServer;

            // Execute Method for Global KMS Hostname
            classInstance.InvokeMethod("SetKeyManagementServiceMachine", inParams, null);

            // Set Specific KMS Hostnames
            foreach (LicenseInstance license in licenses.GetListLicensed())
            {
                // Get Global KMS Hostname WMI Object
                classInstance = new ManagementObject("root\\CIMV2", wmiInfo2 + "'" + license.SKUID + "'", null);

                // Execute Method for Specific KMS Hostname
                classInstance.InvokeMethod("SetKeyManagementServiceMachine", inParams, null);
            }
        }

        /// <summary>
        /// Remove Microsoft Office KMS Port
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        public static void RemoveKMSPortOffice(LicenseList licenses)
        {
            // Call Windows if Microsoft Office uses Windows Licensing Services
            if (OfficeVersion.IsOfficeSPP())
            {
                RemoveKMSPortWMI(licenses, "SoftwareLicensingService.Version=" + "'" + OSVersion.GetSPPSVCVersion() + "'", "SoftwareLicensingProduct.ID=");
            }
            else
            {
                RemoveKMSPortWMI(licenses, "OfficeSoftwareProtectionService.Version=" + "'" + OfficeVersion.GetOSPPSVCVersion() + "'", "OfficeSoftwareProtectionProduct.ID=");
            }
        }

        /// <summary>
        /// Remove Microsoft Windows KMS Port
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        public static void RemoveKMSPortWindows(LicenseList licenses)
        {
            // Windows Vista
            if (Math.Abs(OSVersion.GetWindowsNumber() - 6.0) < Double.Epsilon)
            {
                RemoveKMSPortRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL");
                return;
            }
            // Windows 7 and Later
            RemoveKMSPortRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform");
        }

        /// <summary>
        /// Remove KMS Port using the Registry
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="registryPath">Registry Key where KeyManagementServicePort exists</param>
        private static void RemoveKMSPortRegistry(LicenseList licenses, string registryPath)
        {
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath, true))
            {
                if (registry != null)
                {
                    // Delete Global KMS Port
                    registry.DeleteValue("KeyManagementServicePort");

                    // Only Search KMS Licenses
                    licenses.List = licenses.GetListKMS();
                    foreach (LicenseInstance license in licenses.GetListLicensed())
                    {
                        // Application Specific
                        using (RegistryKey registrySubKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID, true))
                        {
                            if (registrySubKey != null && registrySubKey.GetValue("KeyManagementServicePort") != null)
                            {
                                // Delete Application Specific KMS Port
                                registrySubKey.DeleteValue("KeyManagementServicePort");

                                // SKU Specific
                                using (RegistryKey registrySubKey2 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID + Path.DirectorySeparatorChar + license.SKUID, true))
                                {
                                    if (registrySubKey2 != null && registrySubKey2.GetValue("KeyManagementServicePort") != null)
                                    {
                                        // Delete SKU Specific KMS Port
                                        registrySubKey2.DeleteValue("KeyManagementServicePort");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    throw new ApplicationException(@"HKLM\" + registryPath + " does not exist!");
                }
            }
        }

        /// <summary>
        /// Remove KMS Port using WMI
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="wmiInfo1">WMI Provider and associated data to remove Global KMS Port</param>
        /// <param name="wmiInfo2">WMI Provider and associated data to remove Specific KMS Ports</param>
        private static void RemoveKMSPortWMI(LicenseList licenses, string wmiInfo1, string wmiInfo2)
        {
            // Get Global KMS Port WMI Object
            ManagementObject classInstance = new ManagementObject("root\\CIMV2", wmiInfo1, null);

            // Execute Method for Global KMS Port
            classInstance.InvokeMethod("ClearKeyManagementServicePort", null, null);

            // Remove Specific KMS Ports
            foreach (LicenseInstance license in licenses.GetListLicensed())
            {
                // Get Specific KMS Port WMI Object
                classInstance = new ManagementObject("root\\CIMV2", wmiInfo2 + "'" + license.SKUID + "'", null);

                // Execute Method for Specific KMS Port
                classInstance.InvokeMethod("ClearKeyManagementServicePort", null, null);
            }
        }

        /// <summary>
        /// Set Microsoft Office KMS Port
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="kmsPort">Port to use when connecting to KMS Host</param>
        public static void SetKMSPortOffice(LicenseList licenses, string kmsPort)
        {
            // Call Windows if Microsoft Office uses Windows Licensing Services
            if (OfficeVersion.IsOfficeSPP())
            {
                SetKMSPortWMI(licenses, kmsPort, "SoftwareLicensingService.Version=" + "'" + OSVersion.GetSPPSVCVersion() + "'", "SoftwareLicensingProduct.ID=");
            }
            else
            {
                SetKMSPortWMI(licenses, kmsPort, "OfficeSoftwareProtectionService.Version=" + "'" + OfficeVersion.GetOSPPSVCVersion() + "'", "OfficeSoftwareProtectionProduct.ID=");
            }
        }

        /// <summary>
        /// Set Microsoft Windows KMS Port
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="kmsPort">Port to use when connecting to KMS Host</param>
        public static void SetKMSPortWindows(LicenseList licenses, string kmsPort)
        {
            // Windows Vista
            if (Math.Abs(OSVersion.GetWindowsNumber() - 6.0) < Double.Epsilon)
            {
                SetKMSPortRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL", kmsPort);
                return;
            }
            // Windows 7 and Later
            SetKMSPortRegistry(licenses, @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform", kmsPort);
        }

        /// <summary>
        /// Set KMS Port using the Registry
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="kmsPort">Port to use when connecting to KMS Host</param>
        /// <param name="registryPath">Registry Key where KeyManagementServicePort belongs</param>
        private static void SetKMSPortRegistry(LicenseList licenses, string registryPath, string kmsPort)
        {
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath, true))
            {
                if (registry != null)
                {
                    // Set Global KMS Port
                    registry.SetValue("KeyManagementServicePort", kmsPort);

                    // Only Search KMS Licenses
                    licenses.List = licenses.GetListKMS();
                    foreach (LicenseInstance license in licenses.GetListLicensed())
                    {
                        // Application Specific
                        using (RegistryKey registrySubKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID, true))
                        {
                            if (registrySubKey != null && registrySubKey.GetValue("KeyManagementServicePort") != null)
                            {
                                // Set Application Specific KMS Port
                                registrySubKey.SetValue("KeyManagementServicePort", kmsPort);

                                // SKU Specific
                                using (RegistryKey registrySubKey2 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(registryPath + Path.DirectorySeparatorChar + license.AppID + Path.DirectorySeparatorChar + license.SKUID, true))
                                {
                                    if (registrySubKey2 != null && registrySubKey2.GetValue("KeyManagementServicePort") != null)
                                    {
                                        // Set SKU Specific KMS Port
                                        registrySubKey2.SetValue("KeyManagementServicePort", kmsPort);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    throw new ApplicationException(@"HKLM\" + registryPath + " does not exist");
                }
            }
        }

        /// <summary>
        /// Set KMS Port using WMI
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="kmsPort">KMS Port to set</param>
        /// <param name="wmiInfo1">WMI Provider and associated data to set Global KMS Port</param>
        /// <param name="wmiInfo2">WMI Provider and associated data to set Specific KMS Ports</param>
        private static void SetKMSPortWMI(LicenseList licenses, string kmsPort, string wmiInfo1, string wmiInfo2)
        {
            // Get Global KMS Port WMI Object
            ManagementObject classInstance = new ManagementObject("root\\CIMV2", wmiInfo1, null);

            // Obtain and Add Input Parameters
            ManagementBaseObject inParams = classInstance.GetMethodParameters("SetKeyManagementServicePort");
            inParams["PortNumber"] = Convert.ToUInt32(kmsPort);

            // Execute Method for Global KMS Port
            classInstance.InvokeMethod("SetKeyManagementServicePort", inParams, null);

            // Set Specific KMS Ports
            foreach (LicenseInstance license in licenses.GetListLicensed())
            {
                // Get Specific KMS Port WMI Object
                classInstance = new ManagementObject("root\\CIMV2", wmiInfo2 + "'" + license.SKUID + "'", null);

                // Execute Method for Specific KMS Port
                classInstance.InvokeMethod("SetKeyManagementServicePort", inParams, null);
            }
        }

        /// <summary>
        /// Disable KMS Online Validation
        /// </summary>
        public static void DisableKMSGenuineChecks()
        {
            string[] registryKeys = 
            {
                @"SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform",
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Activation"
                //@"SOFTWARE\Classes\AppID\slui.exe"
            };
            foreach (string regKey in registryKeys)
            {
                using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(regKey, true))
                {
                    if (registry == null)
                    {
                        RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).CreateSubKey(regKey).SetValue("NoGenTicket", 1, RegistryValueKind.DWord);
                    }
                    else
                    {
                        // Set Value
                        registry.SetValue("NoGenTicket", 1, RegistryValueKind.DWord);
                    }
                }
            }
        }
    }

    /// <summary>
    /// Methods for using DLL Injection
    /// </summary>
    public static class KMSDLLInjection
    {
        /// <summary>
        /// Handle to Target Process
        /// </summary>
        private static Process _targetProcess;

        /// <summary>
        /// Architecture Specific and IP Patched DLL
        /// </summary>
        private static byte[] _targetDLL;

        /// <summary>
        /// Path to DLL
        /// </summary>
        private static string _pathDLL;

        /// <summary>
        /// Get Target Process ID for Later DLL Management
        /// </summary>
        public static void Initialize(string ip)
        {
            /*
            // Get Process ID of SVCHOST DcomLaunch
            int procId = 0;
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = 'svchost.exe'"))
            {
                foreach (ManagementBaseObject queryObj in searcher.Get())
                {
                    // Target DcomLaunch
                    if (queryObj["CommandLine"].ToString().Contains("DcomLaunch"))
                    {
                        // Save Process ID
                        procId = Convert.ToInt32(queryObj["ProcessId"]);
                        break;
                    }
                }
            }

            // Check Process ID
            if (procId == 0)
            {
                throw new ApplicationException("DLL Injection target process not found!");
            }

            // Set Process
            _targetProcess = Process.GetProcessById(procId);
            */

            // Set DLL Architecture
            if (Architecture.GetOSArch() == Architecture.X64)
            {
                _targetDLL = Resources.SppExtComObjHook64;
            }
            else
            {
                _targetDLL = Resources.SppExtComObjHook;
            }

            // Set DLL Path
            if (Architecture.GetOSArch() == Architecture.X64 && !Environment.Is64BitProcess)
            {
                _pathDLL = Environment.GetEnvironmentVariable("windir") + @"\Sysnative";
            }
            else
            {
                _pathDLL = Environment.GetEnvironmentVariable("windir") + @"\System32";
            }
            //_targetDLL = Replace(_targetDLL, Encoding.Unicode.GetBytes("123.123.123.123"), Encoding.Unicode.GetBytes(ip));
        }

        /// <summary>
        /// Load DLL into SVCHOST (DcomLaunch) to fool KMS Connection Broker into connecting on LocalHost
        /// </summary>
        /// <returns>True if DLL Injection was successful, False if it failed.</returns>
        public static bool LoadDLL()
        {
            // Create Windows Defender Exclusion
            if (OSVersion.GetWindowsNumber() >= 10.0)
            {
                CommonUtilities.ExecuteCommand("powershell.exe Add-MpPreference -ExclusionPath " + _pathDLL + "\\SppExtComObjHook.dll -Force", true);
            }

            // Create DLL File
            CommonUtilities.FileCreate("SppExtComObjHook.dll", _targetDLL, _pathDLL);

            // Kill KMS Connection Broker
            CommonUtilities.KillProcess("SppExtComObj");

            // Set Permissions on DLL File
            if (File.Exists(_pathDLL + "\\SppExtComObjHook.dll"))
            {
                // Get a FileSecurity object that represents the current security settings.
                FileSecurity fSecurity = File.GetAccessControl(_pathDLL + "\\SppExtComObjHook.dll");

                // Add the FileSystemAccessRule to the security settings.
                fSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null), FileSystemRights.FullControl, AccessControlType.Allow));

                // Set the new access settings.
                File.SetAccessControl(_pathDLL + "\\SppExtComObjHook.dll", fSecurity);
            }

            // IFEO
            using (RegistryKey registry = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\SppExtComObj.exe"))
            {
                if (registry != null)
                {
                    registry.SetValue("Debugger", "RunDll32.exe SppExtComObjHook.dll,PatcherMain", RegistryValueKind.String);
                }
            }
            return true;

            /*
            // Check if DLL is Loaded
            foreach (ProcessModule module in _targetProcess.Modules)
            {
                // Return from Function if DLL is Loaded
                if (module.FileName.Contains("SppExtComObjHook.dll"))
                {
                    return true;
                }
            }

            // Load DLL
            return BInject((uint)_targetProcess.Id, dllPath + "\\SppExtComObjHook.dll");
            */
        }

        /// <summary>
        /// Unload DLL from SVCHOST (DcomLaunch)
        /// </summary>
        /// <returns>True if DLL Ejection was successful, False if it failed.</returns>
        public static bool UnloadDLL()
        {
            // Unload DLL
            bool success = BEject((uint)_targetProcess.Id, _pathDLL + "\\SppExtComObjHook.dll");

            // Delete DLL
            if (success)
            {
                // Unload DLL from KMS Connection Broker
                foreach (Process process in Process.GetProcessesByName("SppExtComObj"))
                {
                    BEject((uint)process.Id, _pathDLL + "\\SppExtComObjHook.dll");
                }
                
                // Delete DLL File
                CommonUtilities.FileDelete(_pathDLL + "\\SppExtComObjHook.dll");
            }

            // Remove Windows Defender Exclusion
            if (OSVersion.GetWindowsNumber() >= 10.0)
            {
                CommonUtilities.ExecuteCommand("powershell.exe Remove-MpPreference -ExclusionPath " + _pathDLL + "\\SppExtComObjHook.dll -Force", true);
            }

            // Return
            return success;
        }

        /// <summary>
        /// Load DLL via Image File Execution Options to fool KMS Connection Broker into connecting on LocalHost
        /// </summary>
        public static void LoadDLLIFEO()
        {
            // Create Windows Defender Exclusion
            if (OSVersion.GetWindowsNumber() >= 10.0)
            {
                CommonUtilities.ExecuteCommand("powershell.exe Add-MpPreference -ExclusionPath " + _pathDLL + "\\SppExtComObjHook.dll -Force", true);
            }

            // Create DLL File
            CommonUtilities.FileCreate("SppExtComObjHook.dll", _targetDLL, _pathDLL);

            // Kill KMS Connection Broker
            CommonUtilities.KillProcess("SppExtComObj");

            // Set Permissions on DLL File
            if (File.Exists(_pathDLL + "\\SppExtComObjHook.dll"))
            {
                // Get a FileSecurity object that represents the current security settings.
                FileSecurity fSecurity = File.GetAccessControl(_pathDLL + "\\SppExtComObjHook.dll");

                // Add the FileSystemAccessRule to the security settings.
                fSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null), FileSystemRights.FullControl, AccessControlType.Allow));

                // Set the new access settings.
                File.SetAccessControl(_pathDLL + "\\SppExtComObjHook.dll", fSecurity);
            }

            // Create Image File Execution Options Settings
            using (RegistryKey registry = Registry.LocalMachine.CreateSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\SppExtComObj.exe"))
            {
                if (registry != null)
                {
                    registry.SetValue("Debugger", "RunDll32.exe " + _pathDLL + "\\SppExtComObjHook.dll,PatcherMain", RegistryValueKind.String);
                }
            }
        }

        /// <summary>
        /// Unload DLL via Image File Execution Options
        /// </summary>
        public static void UnloadDLLIFEO()
        {
            // Kill KMS Connection Broker
            CommonUtilities.KillProcess("SppExtComObj");
            Thread.Sleep(5000);

            // Delete DLL File
            CommonUtilities.FileDelete(_pathDLL + "\\SppExtComObjHook.dll");

            // Remove Windows Defender Exclusion
            if (OSVersion.GetWindowsNumber() >= 10.0)
            {
                CommonUtilities.ExecuteCommand("powershell.exe Remove-MpPreference -ExclusionPath " + _pathDLL + "\\SppExtComObjHook.dll -Force", true);
            }

            // Delete Image File Execution Options Settings
            using (RegistryKey registry = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options", true))
            {
                if (registry != null)
                {
                    registry.DeleteSubKey("SppExtComObj.exe");
                }
            }
        }

        /// <summary>
        /// Inject DLL Into Process
        /// </summary>
        /// <param name="pToBeInjected">Process ID of Target</param>
        /// <param name="sDllPath">Full Path to DLL File</param>
        /// <returns>True if DLL Injection was successful, False if it failed.</returns>
        private static bool BInject(uint pToBeInjected, string sDllPath)
        {
            IntPtr hndProc = NativeMethods.OpenProcess((0x2 | 0x8 | 0x10 | 0x20 | 0x400), 1, pToBeInjected);

            if (hndProc == IntPtr.Zero)
            {
                return false;
            }

            IntPtr lpLlAddress = NativeMethods.GetProcAddress(NativeMethods.GetModuleHandle("kernel32.dll"), "LoadLibraryW");

            if (lpLlAddress == IntPtr.Zero)
            {
                return false;
            }

            IntPtr lpAddress = NativeMethods.VirtualAllocEx(hndProc, IntPtr.Zero, (uint)Encoding.Unicode.GetByteCount(sDllPath), NativeMethods.AllocationType.Commit, NativeMethods.MemoryProtection.ReadWrite);

            if (lpAddress == IntPtr.Zero)
            {
                return false;
            }

            int bytesWritten;
            IntPtr pLibFullPathUnmanaged = Marshal.StringToHGlobalUni(sDllPath);
            if (!NativeMethods.WriteProcessMemory(hndProc, lpAddress, pLibFullPathUnmanaged, (uint)Encoding.Unicode.GetByteCount(sDllPath), out bytesWritten) || bytesWritten != Encoding.Unicode.GetByteCount(sDllPath))
            {
                return false;
            }

            IntPtr hndThread = NativeMethods.CreateRemoteThread(hndProc, IntPtr.Zero, 0, lpLlAddress, lpAddress, 0, IntPtr.Zero);
            if (hndThread == IntPtr.Zero)
            {
                return false;
            }

            if (NativeMethods.WaitForSingleObject(hndThread, 0xFFFFFFFF) != 0)
            {
                return false;
            }

            Marshal.FreeHGlobal(pLibFullPathUnmanaged);
            NativeMethods.VirtualFreeEx(hndProc, lpAddress, 0, NativeMethods.AllocationType.Release);
            NativeMethods.CloseHandle(hndProc);

            return IsInjected(pToBeInjected, sDllPath);
        }

        /// <summary>
        /// Eject DLL from Process
        /// </summary>
        /// <param name="pToBeEjected">Process ID of Target</param>
        /// <param name="sDllPath">Full Path to DLL File</param>
        /// <returns>True if DLL Ejection was successful, False if it failed.</returns>
        private static bool BEject(uint pToBeEjected, string sDllPath)
        {
            Process targetProcess = Process.GetProcessById((int)pToBeEjected);
            foreach (ProcessModule module in targetProcess.Modules)
            {
                // Get Loaded Module
                if (module.FileName == sDllPath)
                {
                    IntPtr hndProc = NativeMethods.OpenProcess((0x2 | 0x8 | 0x10 | 0x20 | 0x400), 1, pToBeEjected);

                    if (hndProc == IntPtr.Zero)
                    {
                        return false;
                    }

                    IntPtr lpFlAddress = NativeMethods.GetProcAddress(NativeMethods.GetModuleHandle("kernel32.dll"), "FreeLibrary");

                    if (lpFlAddress == IntPtr.Zero)
                    {
                        return false;
                    }

                    IntPtr hndThread = NativeMethods.CreateRemoteThread(hndProc, IntPtr.Zero, 0, lpFlAddress, module.BaseAddress, 0, IntPtr.Zero);
                    if (hndThread == IntPtr.Zero)
                    {
                        return false;
                    }

                    if (NativeMethods.WaitForSingleObject(hndThread, 0xFFFFFFFF) != 0)
                    {
                        return false;
                    }

                    NativeMethods.CloseHandle(hndProc);

                    return !IsInjected(pToBeEjected, sDllPath);
                }
            }
            return false;
        }

        /// <summary>
        /// Eject DLL from Process
        /// </summary>
        /// <param name="pToBeChecked">Process ID of Target</param>
        /// <param name="sDllPath">Full Path to DLL File</param>
        /// <returns>True if DLL is Injected, False if it is not Injected.</returns>
        private static bool IsInjected(uint pToBeChecked, string sDllPath)
        {
            Process targetProcess = Process.GetProcessById((int)pToBeChecked);
            foreach (ProcessModule module in targetProcess.Modules)
            {
                if (module.FileName == sDllPath)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Replace Bytes in DLL
        /// </summary>
        /// <param name="input">DLL as Byte Array</param>
        /// <param name="pattern">Byte Search Pattern</param>
        /// <param name="replacement">Byte Replace Pattern</param>
        /// <returns>DLL as Patched Byte Array</returns>
        private static byte[] Replace(byte[] input, byte[] pattern, byte[] replacement)
        {
            if (pattern.Length == 0)
            {
                return input;
            }

            // Size Fix
            if (replacement.Length < pattern.Length)
            {
                byte[] newReplacement = new byte[pattern.Length];
                for (int j = 0; j < pattern.Length; j++)
                {
                    if (j < replacement.Length)
                    {
                        newReplacement[j] = replacement[j];
                    }
                    else
                    {
                        newReplacement[j] = 0;
                    }
                }
                replacement = newReplacement;
            }

            List<byte> result = new List<byte>();

            int i;

            for (i = 0; i <= input.Length - pattern.Length; i++)
            {
                bool foundMatch = true;
                for (int j = 0; j < pattern.Length; j++)
                {
                    if (input[i + j] != pattern[j])
                    {
                        foundMatch = false;
                        break;
                    }
                }

                if (foundMatch)
                {
                    result.AddRange(replacement);
                    i += pattern.Length - 1;
                }
                else
                {
                    result.Add(input[i]);
                }
            }

            for (; i < input.Length; i++)
            {
                result.Add(input[i]);
            }

            return result.ToArray();
        }
    }

    /// <summary>
    /// Methods for using TAP Driver
    /// </summary>
    public static class KMSTAPDriver
    {
        #region TAP Driver Information
        /// <summary>
        /// Struct for TAP Adapter Information
        /// </summary>
        internal struct TAPDriver
        {
            internal string Description;
            internal string DeviceInstanceID;
            internal string GUID;
            internal int Index;
            internal int InterfaceIndex;
            internal string Name;

            internal void DeviceEnable()
            {
                DeviceHelper.EnableDevice(new Guid("{4d36e972-e325-11ce-bfc1-08002be10318}"), DeviceInstanceID);
            }

            internal void DeviceDisable()
            {
                DeviceHelper.DisableDevice(new Guid("{4d36e972-e325-11ce-bfc1-08002be10318}"), DeviceInstanceID);
            }

            internal void DeviceRemove()
            {
                DeviceHelper.RemoveDevice(new Guid("{4d36e972-e325-11ce-bfc1-08002be10318}"), DeviceInstanceID);
            }

            internal void DeviceRestart()
            {
                DeviceHelper.RestartDevice(new Guid("{4d36e972-e325-11ce-bfc1-08002be10318}"), DeviceInstanceID);
            }

            /// <summary>
            /// Get the IP Address of the TAP Adapter
            /// </summary>
            /// <returns>IP Address if the TAP Device has a valid IPv4 Configuratio</returns>
            internal string GetIPAddress()
            {
                // Get IP Address
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    // Get IP Properties
                    try
                    {
                        IPInterfaceProperties adapterProperties = ni.GetIPProperties();
                        IPv4InterfaceProperties p = adapterProperties.GetIPv4Properties();

                        // Check Interface Index
                        if (p.Index == InterfaceIndex)
                        {
                            foreach (UnicastIPAddressInformation ip in adapterProperties.UnicastAddresses)
                            {
                                if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                                {
                                    return ip.Address.ToString();
                                }
                            }
                        }
                    }
                    catch (NetworkInformationException)
                    {
                    }
                }
                return String.Empty;
            }

            /// <summary>
            /// Get the IP Subnet of the TAP Adapter
            /// </summary>
            /// <returns>IP Subnet Mask if the TAP Device has a valid IPv4 Configuration</returns>
            internal string GetIPSubnet()
            {
                // Get IP Subnet
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    try
                    {
                        // Get IP Properties
                        IPInterfaceProperties adapterProperties = ni.GetIPProperties();
                        IPv4InterfaceProperties p = adapterProperties.GetIPv4Properties();

                        // Check Interface Index
                        if (p.Index == InterfaceIndex)
                        {
                            foreach (UnicastIPAddressInformation ip in adapterProperties.UnicastAddresses)
                            {
                                if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                                {
                                    return ip.IPv4Mask.ToString();
                                }
                            }
                        }
                    }
                    catch (NetworkInformationException)
                    {
                    }
                }
                return String.Empty;
            }

            /// <summary>
            /// Get MediaStatus Value for this TAP Adapter
            /// </summary>
            /// <returns>True if MediaStatus is set to 1, False otherwise</returns>
            internal bool GetMediaStatus()
            {
                using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SYSTEM\\ControlSet001\\Control\\Class\\{4d36e972-e325-11ce-bfc1-08002be10318}", true))
                {
                    if (registry != null)
                    {
                        foreach (string subKeyName in registry.GetSubKeyNames())
                        {
                            // Skip Locked Properties Key
                            if (subKeyName == "Properties")
                            {
                                continue;
                            }

                            using (RegistryKey registrySubKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SYSTEM\\ControlSet001\\Control\\Class\\{4d36e972-e325-11ce-bfc1-08002be10318}" + Path.DirectorySeparatorChar + subKeyName, true))
                            {
                                if (registrySubKey != null)
                                {
                                    if (registrySubKey.GetValue("NetCfgInstanceId").ToString() == GUID)
                                    {
                                        if (registrySubKey.GetValue("MediaStatus").ToString() == "1")
                                        {
                                            return true;
                                        }
                                        
                                        return false;
                                    }
                                }
                            }
                        }
                    }
                }
                throw new ApplicationException("MediaStatus value not found!");
            }

            /// <summary>
            /// Set MediaStatus Value for this TAP Adapter
            /// </summary>
            /// <param name="enable">True to Set MediaStatus to 1, False to Set MediaStatus to 0</param>
            internal void SetMediaStatus(bool enable)
            {
                using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SYSTEM\\ControlSet001\\Control\\Class\\{4d36e972-e325-11ce-bfc1-08002be10318}", true))
                {
                    if (registry != null)
                    {
                        foreach (string subKeyName in registry.GetSubKeyNames())
                        {
                            // Skip Locked Properties Key
                            if (subKeyName == "Properties")
                            {
                                continue;
                            }

                            using (RegistryKey registrySubKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SYSTEM\\ControlSet001\\Control\\Class\\{4d36e972-e325-11ce-bfc1-08002be10318}" + Path.DirectorySeparatorChar + subKeyName, true))
                            {
                                if (registrySubKey != null)
                                {
                                    if (registrySubKey.GetValue("NetCfgInstanceId").ToString() == GUID)
                                    {
                                        if (enable)
                                        {
                                            registrySubKey.SetValue("MediaStatus", "1");
                                        }
                                        else
                                        {
                                            registrySubKey.SetValue("MediaStatus", "0");
                                        }

                                        // Restart Service
                                        RestartTAPDeviceDriver();
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Saved Information to Identify Usable TAP Adapter
        /// </summary>
        private static TAPDriver _activeTAPDriver;
        #endregion

        /// <summary>
        /// Handle to Thread for TAP Mirror
        /// </summary>
        private static Thread _tapMirrorThread;

        /// <summary>
        /// Detect and Setup TAP Adapter
        /// </summary>
        public static void Initialize(string ip, string subnet)
        {
            // Clear Active TAP Adapter
            _activeTAPDriver = new TAPDriver();

            // Get Current TAP Adapters
            List<TAPDriver> currentTAPDrivers = GetTAPDrivers();

            // Run TAP Adapter Installer
            InstallTAPDeviceDriver();

            // Get New TAP Adapters
            List<TAPDriver> newTAPDrivers = GetTAPDrivers();
            foreach (TAPDriver newTAPDriver in newTAPDrivers)
            {
                if (currentTAPDrivers.Contains(newTAPDriver))
                {
                    continue;
                }
                _activeTAPDriver = newTAPDriver;
                break;
            }

            // Check Active TAP Adapter
            if (String.IsNullOrEmpty(_activeTAPDriver.GUID))
            {
                throw new ApplicationException("Failed to install TAP Adapter!");
            }

            // set IP Configuration
            SetTAPDeviceIPConfiguration(ip, subnet);
            Thread.Sleep(3000);
        }

        /// <summary>
        /// Start TAP Listener
        /// </summary>
        public static void Start()
        {
            StartTAPMirror();
            StopTAPMirror();
            StartTAPMirror();
            Thread.Sleep(10000);
        }

        /// <summary>
        /// Stop TAP Listener and Remove TAP Adapter
        /// </summary>
        public static void Unload()
        {
            StopTAPMirror();
            RemoveTAPDeviceDriver();
        }

        /// <summary>
        /// Get TAP Drivers as a List
        /// </summary>
        /// <returns>List of Information on All TAP Adapters</returns>
        /// <param name="enableAdapters">Use WMI to enable any disabled TAP Adapters for use and detection</param>
        private static List<TAPDriver> GetTAPDrivers(bool enableAdapters = true)
        {
            // List of TAP Adapters
            List<TAPDriver> tapDrivers = new List<TAPDriver>();

            // Enable All Disabled TAP Adapters
            if (enableAdapters)
            {
                try
                {
                    using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_NetworkAdapter WHERE (ServiceName = 'tap0901' OR ServiceName = 'visctap0901') AND NetEnabled = False"))
                    {
                        foreach (ManagementBaseObject queryObj in searcher.Get())
                        {
                            using (ManagementObject classInstance = new ManagementObject("root\\CIMV2", "Win32_NetworkAdapter.DeviceID='" + queryObj["DeviceID"] + "'", null))
                            {
                                classInstance.InvokeMethod("Enable", null, null);
                            }
                        }
                    }
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch (Exception)
                {
                }
            }

            // Get TAP Adapter Information
            foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
            {
                try
                {
                    // Get IP Properties
                    IPInterfaceProperties adapterProperties = ni.GetIPProperties();
                    IPv4InterfaceProperties p = adapterProperties.GetIPv4Properties();

                    // Check Interface Description
                    if (ni.Description.IndexOf("TAP-Windows Adapter V9", StringComparison.CurrentCultureIgnoreCase) >= 0 || ni.Description.IndexOf("Viscosity Virtual Adapter V9.1", StringComparison.CurrentCultureIgnoreCase) >= 0)
                    {
                        using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT PNPDeviceID, Index FROM Win32_NetworkAdapter WHERE InterfaceIndex = " + p.Index))
                        {
                            foreach (ManagementBaseObject queryObj in searcher.Get())
                            {
                                TAPDriver driver = new TAPDriver { Description = ni.Description, DeviceInstanceID = queryObj["PNPDeviceID"].ToString(), GUID = ni.Id, Index = Convert.ToInt32(queryObj["Index"]) ,InterfaceIndex = p.Index, Name = ni.Name};
                                tapDrivers.Add(driver);
                                break;
                            }
                        }
                    }
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch (Exception)
                {
                }
            }

            // Return Information
            return tapDrivers;
        }

        /// <summary>
        /// Get Number of TAP Adapters installed
        /// </summary>
        /// <returns>Number of TAP Adapters installed</returns>
        public static int GetTAPDeviceCount()
        {
            List<TAPDriver> tapDrivers = GetTAPDrivers();
            return tapDrivers.Count;
        }

        /// <summary>
        /// Get the IP Address of the Active TAP Adapter
        /// </summary>
        /// <returns>IP Address if the TAP Device has a valid IPv4 Configuratio</returns>
        public static string GetTAPDeviceIPAddress()
        {
            // Get TAP Adapter IP Address
            return _activeTAPDriver.GetIPAddress();
        }

        /// <summary>
        /// Get the IP Subnet of the Active TAP Adapter
        /// </summary>
        /// <returns>IP Subnet Mask if the TAP Device has a valid IPv4 Configuration</returns>
        public static string GetTAPDeviceIPSubnet()
        {
            // Get TAP Adapter IP Subnet
            return _activeTAPDriver.GetIPSubnet();
        }

        /// <summary>
        /// Get an IP Address within the TAP Device Network Subnet
        /// </summary>
        /// <returns>A random IP address within the TAP Network Subnet that is not the Network or Broadcast Address</returns>
        public static string GetTAPDeviceUsableIPAddress()
        {
            // Get TAP Device IP Address and Subnet
            string ip = GetTAPDeviceIPAddress();
            string subnet = GetTAPDeviceIPSubnet();

            // Get Network and Broadcast IP Addresses as Integers
            int startIP = BitConverter.ToInt32(CommonUtilities.GetNetworkAddress(IPAddress.Parse(ip), IPAddress.Parse(subnet)).GetAddressBytes().Reverse().ToArray(), 0);
            int currentIP = BitConverter.ToInt32(IPAddress.Parse(ip).GetAddressBytes().Reverse().ToArray(), 0);
            int endIP = BitConverter.ToInt32(CommonUtilities.GetBroadcastAddress(IPAddress.Parse(ip), IPAddress.Parse(subnet)).GetAddressBytes().Reverse().ToArray(), 0);

            // Generate Random IP Address Between Start and End IP Addresses and Exclude Current IP
            Random random = new Random();
            while (true)
            {
                int usableIP = random.Next(startIP + 1, endIP);
                if (usableIP != currentIP)
                {
                    return new IPAddress(BitConverter.GetBytes(usableIP).Reverse().ToArray()).ToString();
                }
            }
        }

        // TODO: Native
        /// <summary>
        /// Install TAP Adapter Driver
        /// </summary>
        private static void InstallTAPDeviceDriver()
        {
            // List of TAP Adapters
            List<TAPDriver> tapDrivers = GetTAPDrivers();

            // Installer Variables
            //const string installTAPAdapterFileName = "InstallTAPAdapter.exe";
            string installTAPAdapterFilePath = Environment.GetEnvironmentVariable("TEMP");

            if (tapDrivers.Count > 0)
            {
                foreach (TAPDriver tapDriver in tapDrivers)
                {
                    if (tapDriver.Description.IndexOf("TAP-Windows Adapter V9", StringComparison.CurrentCultureIgnoreCase) >= 0)
                    {
                        // Run TAP Adapter Installer Viscosity
                        //CommonUtilities.FileCreate(installTAPAdapterFileName, Resources.InstallTAPAdapterViscosity, installTAPAdapterFilePath);
                        //CommonUtilities.ExecuteCommand(installTAPAdapterFilePath + Path.DirectorySeparatorChar + installTAPAdapterFileName, false);

                        CommonUtilities.FileCreate("tapcert.cer", Resources.SparkLabs_Certificate, installTAPAdapterFilePath);
                        CommonUtilities.FileCreate("tapinstall.cmd", Resources.SparkLabs_Installer, installTAPAdapterFilePath);
                        if (Architecture.GetOSArch() == Architecture.X86)
                        {
                            CommonUtilities.FileCreate("visctap0901.cat", Resources.SparkLabs_Catalog_x86, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("visctap0901.inf", Resources.SparkLabs_INF_x86, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("visctap0901.sys", Resources.SparkLabs_Driver_x86, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tapinstall.exe", Resources.TAPInstall_x86, installTAPAdapterFilePath);
                        }
                        else
                        {
                            CommonUtilities.FileCreate("visctap0901.cat", Resources.SparkLabs_Catalog_x64, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("visctap0901.inf", Resources.SparkLabs_INF_x64, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("visctap0901.sys", Resources.SparkLabs_Driver_x64, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tapinstall.exe", Resources.TAPInstall_x64, installTAPAdapterFilePath);
                        }
                        CommonUtilities.ExecuteCommand(installTAPAdapterFilePath + Path.DirectorySeparatorChar + "tapinstall.cmd", false);
                        return;
                    }
                    if (tapDriver.Description.IndexOf("Viscosity Virtual Adapter V9.1", StringComparison.CurrentCultureIgnoreCase) >= 0)
                    {
                        // Run TAP Adapter Installer OpenVPN
                        //CommonUtilities.FileCreate(installTAPAdapterFileName, Resources.InstallTAPAdapter, installTAPAdapterFilePath);
                        //CommonUtilities.ExecuteCommand(installTAPAdapterFilePath + Path.DirectorySeparatorChar + installTAPAdapterFileName, false);

                        CommonUtilities.FileCreate("tapcert.cer", Resources.OpenVPN_Certificate, installTAPAdapterFilePath);
                        CommonUtilities.FileCreate("tapinstall.cmd", Resources.OpenVPN_Installer, installTAPAdapterFilePath);
                        if (Architecture.GetOSArch() == Architecture.X86)
                        {
                            CommonUtilities.FileCreate("tap0901.cat", Resources.OpenVPN_Catalog_x86, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tap0901.inf", Resources.OpenVPN_INF_x86, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tap0901.sys", Resources.OpenVPN_Driver_x86, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tapinstall.exe", Resources.TAPInstall_x86, installTAPAdapterFilePath);
                        }
                        else
                        {
                            CommonUtilities.FileCreate("tap0901.cat", Resources.OpenVPN_Catalog_x64, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tap0901.inf", Resources.OpenVPN_INF_x64, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tap0901.sys", Resources.OpenVPN_Driver_x64, installTAPAdapterFilePath);
                            CommonUtilities.FileCreate("tapinstall.exe", Resources.TAPInstall_x64, installTAPAdapterFilePath);
                        }
                        CommonUtilities.ExecuteCommand(installTAPAdapterFilePath + Path.DirectorySeparatorChar + "tapinstall.cmd", false);
                        return;
                    }
                }
            }

            // Run TAP Adapter Installer
            //CommonUtilities.FileCreate(installTAPAdapterFileName, Resources.InstallTAPAdapter, installTAPAdapterFilePath);
            //CommonUtilities.ExecuteCommand(installTAPAdapterFilePath + Path.DirectorySeparatorChar + installTAPAdapterFileName, false);
            //CommonUtilities.FileDelete(installTAPAdapterFilePath + Path.DirectorySeparatorChar + installTAPAdapterFileName);

            CommonUtilities.FileCreate("tapcert.cer", Resources.OpenVPN_Certificate, installTAPAdapterFilePath);
            CommonUtilities.FileCreate("tapinstall.cmd", Resources.OpenVPN_Installer, installTAPAdapterFilePath);
            if (Architecture.GetOSArch() == Architecture.X86)
            {
                CommonUtilities.FileCreate("tap0901.cat", Resources.OpenVPN_Catalog_x86, installTAPAdapterFilePath);
                CommonUtilities.FileCreate("tap0901.inf", Resources.OpenVPN_INF_x86, installTAPAdapterFilePath);
                CommonUtilities.FileCreate("tap0901.sys", Resources.OpenVPN_Driver_x86, installTAPAdapterFilePath);
                CommonUtilities.FileCreate("tapinstall.exe", Resources.TAPInstall_x86, installTAPAdapterFilePath);
            }
            else
            {
                CommonUtilities.FileCreate("tap0901.cat", Resources.OpenVPN_Catalog_x64, installTAPAdapterFilePath);
                CommonUtilities.FileCreate("tap0901.inf", Resources.OpenVPN_INF_x64, installTAPAdapterFilePath);
                CommonUtilities.FileCreate("tap0901.sys", Resources.OpenVPN_Driver_x64, installTAPAdapterFilePath);
                CommonUtilities.FileCreate("tapinstall.exe", Resources.TAPInstall_x64, installTAPAdapterFilePath);
            }
            CommonUtilities.ExecuteCommand(installTAPAdapterFilePath + Path.DirectorySeparatorChar + "tapinstall.cmd", false);
        }

        /// <summary>
        /// Remove TAP Adapter Driver
        /// </summary>
        private static void RemoveTAPDeviceDriver()
        {
           _activeTAPDriver.DeviceRemove();
        }

        /// <summary>
        /// Restart TAP Adapter Driver
        /// </summary>
        public static void RestartTAPDeviceDriver()
        {
            _activeTAPDriver.DeviceRestart();
        }

        /// <summary>
        /// Set TAP Adapter to use an IP and Subnet Address
        /// </summary>
        /// <param name="ip">IP Address to Use</param>
        /// <param name="subnet">IP Subnet to Use</param>
        public static void SetTAPDeviceIPConfiguration(string ip, string subnet)
        {
            // Check MediaStatus
            bool mediaStatus = _activeTAPDriver.GetMediaStatus();
            if (!mediaStatus)
            {
                _activeTAPDriver.SetMediaStatus(true);
            }

            // Used IP Check
            List<TAPDriver> tapDrivers = GetTAPDrivers();
            foreach (TAPDriver tapDriver in tapDrivers)
            {
                if (tapDriver.GetIPAddress() == ip)
                {
                    try
                    {
                        // Enable DHCP
                        using (ManagementObject classInstance = new ManagementObject("root\\CIMV2", "Win32_NetworkAdapterConfiguration.Index='" + tapDriver.Index + "'", null))
                        {
                            classInstance.InvokeMethod("EnableDHCP", null, null);
                        }

                        // Restart Device
                        tapDriver.DeviceRestart();
                    }
                    // ReSharper disable once EmptyGeneralCatchClause
                    catch (Exception)
                    {
                    }
                }
            }

            // Set Static IP and Force Static IP in Network Properties
            /*
            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT SettingID FROM Win32_NetworkAdapterConfiguration WHERE InterfaceIndex = " + _activeTAPDriver.InterfaceIndex))
                {
                    foreach (ManagementBaseObject queryObj in searcher.Get())
                    {
                        using (RegistryKey regKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\Tcpip\Parameters\Interfaces\" + queryObj["SettingID"], true))
                        {
                            if (regKey != null)
                            {
                               regKey.SetValue("EnableDHCP", 0);
                               regKey.SetValue("IPAddress", new[] {ip}, RegistryValueKind.MultiString);
                               regKey.SetValue("SubnetMask", new[] {subnet}, RegistryValueKind.MultiString);
                            }
                        }
                    }
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }
            */
            

            // Set IP Configuration
            using (ManagementObject classInstance2 = new ManagementObject("root\\CIMV2", "Win32_NetworkAdapterConfiguration.Index='" + _activeTAPDriver.Index + "'", null))
            {
                // Enable Static
                using (ManagementBaseObject inParams = classInstance2.GetMethodParameters("EnableStatic"))
                {
                    inParams["IPAddress"] = new[] {ip};
                    inParams["SubnetMask"] = new[] { subnet };
                    classInstance2.InvokeMethod("EnableStatic", inParams, null);
                }
            }
        }

        /// <summary>
        /// Start the TAP Mirror Thread
        /// </summary>
        private static void StartTAPMirror()
        {
            _tapMirrorThread = new Thread(TAPMirrorListener);
            _tapMirrorThread.Start();
        }

        /// <summary>
        /// Stop the TAP Mirror Thread
        /// </summary>
        private static void StopTAPMirror()
        {
            if (_tapMirrorThread != null && _tapMirrorThread.IsAlive)
            {
                _tapMirrorThread.Abort();
            }
        }

        #region TAP Control Constants and Variables
        private const uint MethodBuffered = 0;
        private const uint FileAnyAccess = 0;
        private const uint FileDeviceUnknown = 0x00000022;

        static FileStream _tap;
        static int _bytesRead;
        static SafeFileHandle _tapHandle;

        private const int FileAttributeSystem = 0x4;
        private const int FileFlagOverlapped = 0x40000000;
        #endregion

        #region TAP Control Functions
        private static uint CtlCode(uint deviceType, uint function, uint method, uint access)
        {
            return ((deviceType << 16) | (access << 14) | (function << 2) | method);
        }

        private static uint TAPControlCode(uint request, uint method)
        {
            return CtlCode(FileDeviceUnknown, request, method, FileAnyAccess);
        }

        private static void TAPMirrorListener()
        {
            // Get TAP Adapter
            TAPDriver tapDriver = _activeTAPDriver;

            // Get TAP Device IP Address and Subnet
            string ip = GetTAPDeviceIPAddress();
            string subnet = GetTAPDeviceIPSubnet();

            // Check MediaStatus
            bool mediaStatus = tapDriver.GetMediaStatus();
            if (!mediaStatus)
            {
                tapDriver.SetMediaStatus(true);
            }

            // Setup Handle for TAP Adapter Network I/O
            const string usermodeDeviceSpace = "\\\\.\\Global\\";
            string devGuid = tapDriver.GUID;
            _tapHandle = NativeMethods.CreateFile(usermodeDeviceSpace + devGuid + ".tap", FileAccess.ReadWrite, FileShare.ReadWrite, 0, FileMode.Open, FileAttributeSystem | FileFlagOverlapped, IntPtr.Zero);
            int len;
            IntPtr pstatus = Marshal.AllocHGlobal(4);
            Marshal.WriteInt32(pstatus, 1);
            NativeMethods.DeviceIoControl(_tapHandle, TAPControlCode(6, MethodBuffered) /* TAP_IOCTL_SET_MEDIA_STATUS */, pstatus, 4, pstatus, 4, out len, IntPtr.Zero);
            IntPtr ptun = Marshal.AllocHGlobal(12);
            Marshal.WriteInt32(ptun, 0, BitConverter.ToInt32(IPAddress.Parse(ip).GetAddressBytes(), 0)); // TAP IP as Bytes
            Marshal.WriteInt32(ptun, 4, BitConverter.ToInt32(CommonUtilities.GetNetworkAddress(IPAddress.Parse(ip), IPAddress.Parse(subnet)).GetAddressBytes(), 0)); // TAP Network IP as Bytes
            Marshal.WriteInt32(ptun, 8, unchecked(BitConverter.ToInt32(IPAddress.Parse(subnet).GetAddressBytes(), 0))); // TAP Subnet IP as Bytes
            NativeMethods.DeviceIoControl(_tapHandle, TAPControlCode(10, MethodBuffered) /* TAP_IOCTL_CONFIG_TUN */, ptun, 12, ptun, 12, out len, IntPtr.Zero);
            using (_tap = new FileStream(_tapHandle, FileAccess.ReadWrite, 10000, true))
            {
                byte[] buf = new byte[10000];

                while (true)
                {
                    // Synchronous Read
                    _bytesRead = _tap.Read(buf, 0, 10000);

                    // Reverse IPv4 addresses and send back to tun
                    for (int i = 0; i < 4; ++i)
                    {
                        byte tmp = buf[12 + i]; buf[12 + i] = buf[16 + i]; buf[16 + i] = tmp;
                    }

                    // Synchronous Write
                    _tap.Write(buf, 0, _bytesRead);
                }
            }
            // ReSharper disable once FunctionNeverReturns
        }
        #endregion
    }

    /// <summary>
    /// Methods for using WinDivert Client
    /// </summary>
    public static class KMSWinDivert
    {
        /// <summary>
        /// Handle to Thread for WinDivert Client
        /// </summary>
        private static Thread _winDivertThread;

        /// <summary>
        /// Log any Output from the WinDivert Loader
        /// </summary>
        public static string Log { get; private set; }

        /// <summary>
        /// Save WinDivert IP Address and Subnet
        /// </summary>
        private static string _ipAddress;
        private static string _ipSubnet;

        /// <summary>
        /// Save IP Forwarding Route for Later Deletion
        /// </summary>
        private static NativeMethods.MibIPForwardRow _route;

        /// <summary>
        /// WinDivert Path and File Names
        /// </summary
        private static string winDivertLoaderFileName = "WinDivert Loader.exe";
        private static string winDivertLoaderFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\Microsoft Toolkit\\WinDivert-KMS";

        /// <summary>
        /// Get an IP Address within the WinDivert Client Network Subnet
        /// </summary>
        /// <returns>A random IP address within the WinDivert Client Network Subnet that is not the Network or Broadcast Address</returns>
        private static string GetWinDivertClientUsableIPAddress()
        {
            // Get WinDivert Client IP Address and Subnet
            string ip = _ipAddress;
            string subnet = _ipSubnet;

            // Get Network and Broadcast IP Addresses as Integers
            int startIP = BitConverter.ToInt32(CommonUtilities.GetNetworkAddress(IPAddress.Parse(ip), IPAddress.Parse(subnet)).GetAddressBytes().Reverse().ToArray(), 0);
            int currentIP = BitConverter.ToInt32(IPAddress.Parse(ip).GetAddressBytes().Reverse().ToArray(), 0);
            int endIP = BitConverter.ToInt32(CommonUtilities.GetBroadcastAddress(IPAddress.Parse(ip), IPAddress.Parse(subnet)).GetAddressBytes().Reverse().ToArray(), 0);

            // Generate Random IP Address Between Start and End IP Addresses and Exclude Current IP
            Random random = new Random();
            while (true)
            {
                int usableIP = random.Next(startIP + 1, endIP);
                if (usableIP != currentIP)
                {
                    return new IPAddress(BitConverter.GetBytes(usableIP).Reverse().ToArray()).ToString();
                }
            }
        }

        /// <summary>
        /// Start the WinDivert Client Thread
        /// </summary>
        /// <param name="ip">IP Address to Use</param>
        /// <param name="subnet">IP Subnet to Use</param>
        public static void StartWinDivertClient(string ip, string subnet)
        {
            // Save IP Address and Subnet
            _ipAddress = ip;
            _ipSubnet = subnet;

            // Clear Log
            Log = string.Empty;

            // Create Route
            _route = new NativeMethods.MibIPForwardRow { dwForwardProto = (uint)NativeMethods.NlRouteProtocol.MibIPProtoNetmgmt, dwForwardDest = BitConverter.ToUInt32(IPAddress.Parse(ip).GetAddressBytes(), 0), dwForwardMask = 4294967295, dwForwardIfIndex = 1, dwForwardNextHop = BitConverter.ToUInt32(IPAddress.Parse(GetWinDivertClientUsableIPAddress()).GetAddressBytes(), 0), dwForwardMetric1 = 51 };
            // Get Software Loopback Interface Interface Index and Metric
            foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
            {
                try
                {
                    // Get IP Properties
                    IPInterfaceProperties adapterProperties = ni.GetIPProperties();
                    IPv4InterfaceProperties p = adapterProperties.GetIPv4Properties();

                    // Check Interface Description
                    if (ni.Description.IndexOf("Software Loopback Interface 1", StringComparison.CurrentCultureIgnoreCase) >= 0)
                    {
                        using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT Metric1 FROM Win32_IP4RouteTable WHERE InterfaceIndex = " + p.Index))
                        {
                            foreach (ManagementBaseObject queryObj in searcher.Get())
                            {
                                _route.dwForwardMetric1 = Convert.ToUInt32(queryObj["Metric1"]);
                                break;
                            }
                        }
                        _route.dwForwardIfIndex = (uint)p.Index;
                        break;
                    }
                }
                // ReSharper disable once EmptyGeneralCatchClause
                catch (Exception)
                {
                }
            }
            NativeMethods.CreateIpForwardEntry(ref _route);

            // Create Windows Defender Exclusion
            if (OSVersion.GetWindowsNumber() >= 10.0)
            {
                CommonUtilities.ExecuteCommand("powershell.exe Add-MpPreference -ExclusionPath " + CommonUtilities.EscapePath(winDivertLoaderFilePath) + " -Force", true);
            }

            // Start Thread
            _winDivertThread = new Thread(WinDivertClientRunner);
            _winDivertThread.Start();

            // Check if Running
            for (int i = 0; i < 5; i++)
            {
                Process[] pname = Process.GetProcessesByName("WinDivert Loader");
                if (pname.Length == 0)
                {
                    Thread.Sleep(1000);
                }
                else
                {
                    return;
                }
            }
            Log += "Failed to start WinDivert Loader!" + Environment.NewLine;
        }

        /// <summary>
        /// Stop the WinDivert Client Thread
        /// </summary>
        public static void StopWinDivertClient()
        {
            // Stop WinDivert Loader
            CommonUtilities.KillProcess("WinDivert Loader");
            Thread.Sleep(3000);

            try
            {
                // Stop Service
                //Services.StopService("WinDivert1.4");
                //Services.UninstallService("WinDivert1.4");
                CommonUtilities.ExecuteCommand("sc stop WinDivert1.4", true);
                CommonUtilities.ExecuteCommand("sc delete WinDivert1.4", true);
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }

            try
            {
                // Remove Windows Defender Exclusion
                if (OSVersion.GetWindowsNumber() >= 10.0)
                {
                    string cmd = CommonUtilities.EscapePath(winDivertLoaderFilePath);
                    CommonUtilities.ExecuteCommand("powershell.exe Remove-MpPreference -ExclusionPath " + cmd + " -Force", true);
                }

                CommonUtilities.FolderDelete(winDivertLoaderFilePath);
            }
            catch (Exception)
            {
                Log += "Failed to remove WinDivert Directory." + Environment.NewLine;
            }

            // Delete Route
            NativeMethods.DeleteIpForwardEntry(ref _route);
        }

        /// <summary>
        /// Run WinDivert Loader without Blocking Main Thread
        /// </summary>
        private static void WinDivertClientRunner()
        {
            // Run WinDivert Loader
            CommonUtilities.FileCreate(winDivertLoaderFileName, Resources.WinDivert_Loader, winDivertLoaderFilePath);
            Result result = CommonUtilities.ExecuteCommand(CommonUtilities.EscapePath(winDivertLoaderFilePath + Path.DirectorySeparatorChar + winDivertLoaderFileName) + _ipAddress, true);
            Log += result.Output;
        }
    }
}