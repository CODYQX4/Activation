using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Management;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Common;
using Keys;
using KMSEmulator;
using License;
using Microsoft.Win32;
using ProductDetection;

namespace Activation
{
    /// <summary>
    /// Group of methods for Checking Activation Status and Performing Activation
    /// </summary>
    public static class ActivationState
    {
        /// <summary>
        /// Attempt Activation on each Microsoft Office License in a List
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="minimalOutput">Reduce the Amount of Output During Activation</param>
        /// <param name="kmsServer">KMS Host to Connect To</param>
        /// <param name="kmsPort">KMS Port to Connect To</param>
        /// <param name="kmsPid">KMS PID to apply to KMSEmulator</param>
        /// <param name="kmsHwid">KMS Hardware ID to apply to KMSEmulator</param>
        /// <param name="useKMSEmulator">Start a KMSEmulator Process</param>
        /// <param name="removeKMSConnection">Remove KMS Host and Port after Activation</param>
        /// <param name="killProcessOnPort">Force Start KMSEmulator by Killing Other Processes usign the KMS Port</param>
        /// <param name="useDLLInjection">Use DLL Injection to Force Localhost KMS Activation</param>
        /// <param name="useTAPAdapter">Use TAP Adapter to Force Localhost KMS Activation</param>
        /// <param name="useWinDivert">Use WinDivert Client to Force Localhost KMS Activation</param>
        /// <param name="localHostBypassIPAddress">IP Address of TAP Adapter NIC or WinDivert Client</param>
        /// <param name="localHostBypassIPSubnet">Subnet Mask for TAP Adapter or WinDivert Client Network</param>
        /// <returns>Activation Result of all Microsoft Office Licenses</returns>
        public static string AttemptActivationOffice(LicenseList licenses, bool minimalOutput = false, string kmsServer = "127.0.0.2", int kmsPort = 1688, string kmsPid = "RandomKMSPID", string kmsHwid = "364F463A8863D35F", bool useKMSEmulator = true, bool removeKMSConnection = false, bool killProcessOnPort = false, bool useDLLInjection = false, bool useTAPAdapter = false, bool useWinDivert = false, string localHostBypassIPAddress = "10.3.0.1", string localHostBypassIPSubnet = "255.255.255.0")
        {
            // Call Windows if Microsoft Office uses Windows Licensing Services
            if (OfficeVersion.IsOfficeSPP())
            {
                return AttemptActivation(licenses, "SoftwareLicensingProduct.ID=", minimalOutput, kmsServer, kmsPort, kmsPid, kmsHwid, useKMSEmulator, removeKMSConnection, killProcessOnPort, useDLLInjection, useTAPAdapter, useWinDivert, localHostBypassIPAddress, localHostBypassIPSubnet);
            }
            return AttemptActivation(licenses, "OfficeSoftwareProtectionProduct.ID=", minimalOutput, kmsServer, kmsPort, kmsPid, kmsHwid, useKMSEmulator, removeKMSConnection, killProcessOnPort, useDLLInjection, useTAPAdapter, useWinDivert, localHostBypassIPAddress, localHostBypassIPSubnet);
        }

        /// <summary>
        /// Attempt Activation on each Microsoft Windows License in a List
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="minimalOutput">Reduce the Amount of Output During Activation</param>
        /// <param name="kmsServer">KMS Host to Connect To</param>
        /// <param name="kmsPort">KMS Port to Connect To</param>
        /// <param name="kmsPid">KMS PID to apply to KMSEmulator</param>
        /// <param name="kmsHwid">KMS Hardware ID to apply to KMSEmulator</param>
        /// <param name="useKMSEmulator">Start a KMSEmulator Process</param>
        /// <param name="removeKMSConnection">Remove KMS Host and Port after Activation</param>
        /// <param name="killProcessOnPort">Force Start KMSEmulator by Killing Other Processes usign the KMS Port</param>
        /// <param name="useDLLInjection">Use DLL Injection to Force Localhost KMS Activation</param>
        /// <param name="useTAPAdapter">Use TAP Adapter to Force Localhost KMS Activation</param>
        /// <param name="useWinDivert">Use WinDivert Client to Force Localhost KMS Activation</param>
        /// <param name="localHostBypassIPAddress">IP Address of TAP Adapter NIC or WinDivert Client</param>
        /// <param name="localHostBypassIPSubnet">Subnet Mask for TAP Adapter or WinDivert Client Network</param>
        /// <returns>Activation Result of all Microsoft Windows Licenses</returns>
        public static string AttemptActivationWindows(LicenseList licenses, bool minimalOutput = false, string kmsServer = "127.0.0.2", int kmsPort = 1688, string kmsPid = "RandomKMSPID", string kmsHwid = "364F463A8863D35F", bool useKMSEmulator = true, bool removeKMSConnection = false, bool killProcessOnPort = false, bool useDLLInjection = false, bool useTAPAdapter = false, bool useWinDivert = false, string localHostBypassIPAddress = "10.3.0.1", string localHostBypassIPSubnet = "255.255.255.0")
        {
            return AttemptActivation(licenses, "SoftwareLicensingProduct.ID=", minimalOutput, kmsServer, kmsPort, kmsPid, kmsHwid, useKMSEmulator, removeKMSConnection, killProcessOnPort, useDLLInjection, useTAPAdapter, useWinDivert, localHostBypassIPAddress, localHostBypassIPSubnet);
        }

        /// <summary>
        /// Attempt Activation on each License in a List
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="minimalOutput">Reduce the Amount of Output During Activation</param>
        /// <param name="wmiInfo">WMI Provider and associated data to attempt Activation</param>
        /// <param name="kmsServer">KMS Host to Connect To</param>
        /// <param name="kmsPort">KMS Port to Connect To</param>
        /// <param name="kmsPid">KMS PID to apply to KMSEmulator</param>
        /// <param name="kmsHwid">KMS Hardware ID to apply to KMSEmulator</param>
        /// <param name="useKMSEmulator">Start a KMSEmulator Process</param>
        /// <param name="removeKMSConnection">Remove KMS Host and Port after Activation</param>
        /// <param name="killProcessOnPort">Force Start KMSEmulator by Killing Other Processes usign the KMS Port</param>
        /// <param name="useDLLInjection">Use DLL Injection to Force Localhost KMS Activation</param>
        /// <param name="useTAPAdapter">Use TAP Adapter to Force Localhost KMS Activation</param>
        /// <param name="useWinDivert">Use WinDivert Client to Force Localhost KMS Activation</param>
        /// <param name="localHostBypassIPAddress">IP Address of TAP Adapter NIC or WinDivert Client</param>
        /// <param name="localHostBypassIPSubnet">Subnet Mask for TAP Adapter or WinDivert Client Network</param>
        /// <returns>Activation Result of all Licenses</returns>
        private static string AttemptActivation(LicenseList licenses, string wmiInfo, bool minimalOutput = false, string kmsServer = "127.0.0.2", int kmsPort = 1688, string kmsPid = "RandomKMSPID", string kmsHwid = "364F463A8863D35F", bool useKMSEmulator = true, bool removeKMSConnection = false, bool killProcessOnPort = false, bool useDLLInjection = false, bool useTAPAdapter = false, bool useWinDivert = false, string localHostBypassIPAddress = "10.3.0.1", string localHostBypassIPSubnet = "255.255.255.0")
        {
            using (StringWriter output = new StringWriter())
            {
                // Show Activation Errors if No Licenses or Keys Exist
                if (licenses.GetListUnlicensed().Count == 0 && licenses.GetListLicensed().Count == 0)
                {
                    return LicenseErrorCode.ErrBroken;
                }
                if (licenses.GetListUnlicensed().Count > 0 && licenses.GetListLicensed().Count == 0)
                {
                    return LicenseErrorCode.ErrKeyless;
                }

                // Get Firewall Parameters
                string programName = CommonUtilities.EscapePath(System.Reflection.Assembly.GetEntryAssembly().GetName().Name);
                string programLocation = CommonUtilities.EscapePath(Process.GetCurrentProcess().MainModule.FileName);

                // Delete Block Firewall Rules
                CommonUtilities.ExecuteCommand("netsh advfirewall firewall delete rule name=all program=" + programLocation, true);
                CommonUtilities.ExecuteCommand("netsh advfirewall firewall delete rule name=all localport=" + kmsPort, true);

                // Add Allow Firewall Rules
                CommonUtilities.ExecuteCommand("netsh advfirewall firewall add rule name=" + programName + " dir=in program=" + programLocation + " localport=" + kmsPort + " protocol=TCP action=allow remoteip=any", true);
                CommonUtilities.ExecuteCommand("netsh advfirewall firewall add rule name=" + programName + " dir=out program=" + programLocation + " localport=" + kmsPort + " protocol=TCP action=allow remoteip=any", true);

                // Setup Localhost Bypass for KMS V6
                if (OSVersion.GetWindowsNumber() >= 6.3 && ((licenses is LicenseListOffice && OfficeVersion.GetOfficeNumber() >= 15) || (licenses is LicenseListWindows)))
                {
                    // Disable KMS Online Validation
                    KMSConnection.DisableKMSGenuineChecks();

                    // Use Localhost Bypass for Loopback IP Addresses or Machine Name
                    if (Regex.IsMatch(kmsServer.ToLower(), @"^(127(\.\d+){1,3}|[0:]+1|localhost)$") || String.Compare(kmsServer, Environment.MachineName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        if (useDLLInjection)
                        {
                            // Disable Other Bypasses
                            useTAPAdapter = false;
                            useWinDivert = false;

                            try
                            {
                                KMSDLLInjection.Initialize(localHostBypassIPAddress);
                                /*
                                if (!KMSDLLInjection.LoadDLL())
                                {
                                    output.WriteLine("Failed to inject LocalHost Bypass DLL.");
                                }
                                else
                                {
                                    kmsServer = localHostBypassIPAddress;
                                }
                                */
                                KMSDLLInjection.LoadDLLIFEO();
                                kmsServer = localHostBypassIPAddress;
                            }
                            catch (Exception ex)
                            {
                                output.WriteLine("Failed to inject LocalHost Bypass DLL.");
                                output.WriteLine(ex.Message);
                            }                            
                        }
                        else if (useTAPAdapter)
                        {
                            // Disable Other Bypasses
                            useWinDivert = false;

                            // Check Installed TAP Adapter Count
                            if (KMSTAPDriver.GetTAPDeviceCount() > 1)
                            {
                                output.WriteLine("WARNING: You have more than 1 TAP Adapter installed.");
                            }

                            // Initialize TAP Driver
                            KMSTAPDriver.Initialize(localHostBypassIPAddress, localHostBypassIPSubnet);

                            // Check and Set IP Address
                            string tapIP = KMSTAPDriver.GetTAPDeviceIPAddress();
                            string tapSubnet = KMSTAPDriver.GetTAPDeviceIPSubnet();
                            if (tapIP == string.Empty || Regex.IsMatch(tapIP, @"^(169\.254\.([0,1]?[0-9]{1,2}|2[0-4][0-9]|25[0-5])\.([0,1]?[0-9]{1,2}|2[0-4][0-9]|25[0-5]))$"))
                            {
                                output.WriteLine("Failed to set IP Address on TAP Adapter. It may already be in use.");
                            }           
                            else if (tapIP != localHostBypassIPAddress || tapSubnet != localHostBypassIPSubnet)
                            {
                                KMSTAPDriver.SetTAPDeviceIPConfiguration(localHostBypassIPAddress, localHostBypassIPSubnet);
                                output.WriteLine("Failed to update IP Address Configuration on TAP Adapter.");
                            }
                            else
                            {
                                // Set KMS Server to Valid TAP IP Address
                                kmsServer = KMSTAPDriver.GetTAPDeviceUsableIPAddress();

                                // Start TAP Listener
                                KMSTAPDriver.Start();
                            }
                        }
                        else if (useWinDivert)
                        {
                            // Start WinDivert
                            KMSWinDivert.StartWinDivertClient(localHostBypassIPAddress, localHostBypassIPSubnet);
                            kmsServer = localHostBypassIPAddress;
                        }
                    }
                    else
                    {
                        // Disable Localhost Bypass
                        useDLLInjection = false;
                        useTAPAdapter = false;
                        useWinDivert = false;
                    }
                }
                else
                {
                    // Disable Localhost Bypass
                    useDLLInjection = false;
                    useTAPAdapter = false;
                    useWinDivert = false;
                }

                // Set KMS Server
                try
                {
                    if (licenses is LicenseListOffice)
                    {
                        KMSConnection.SetKMSHostOffice(licenses, kmsServer);
                    }
                    else if (licenses is LicenseListWindows)
                    {
                        KMSConnection.SetKMSHostWindows(licenses, kmsServer);
                    }
                }
                catch (Exception ex)
                {
                    output.WriteLine("Failed to set KMS Host!");
                    output.WriteLine(ex.Message);
                }

                // Set KMS Port
                try
                {
                    if (licenses is LicenseListOffice)
                    {
                        KMSConnection.SetKMSPortOffice(licenses, kmsPort.ToString(CultureInfo.InvariantCulture));
                    }
                    else if (licenses is LicenseListWindows)
                    {
                        KMSConnection.SetKMSPortWindows(licenses, kmsPort.ToString(CultureInfo.InvariantCulture));
                    }
                }
                catch (Exception ex)
                {
                    output.WriteLine("Failed to set KMS Port");
                    output.WriteLine(ex.Message);
                }

                // Start KMSEmulator
                if (useKMSEmulator)
                {
                    try
                    {
                        // Set KMS Server Settings
                        KMSServerSettings settings = new KMSServerSettings {KillProcessOnPort = killProcessOnPort, Port = kmsPort};

                        // Handle KMS PID Generation
                        if (kmsPid == "ReuseKMSPID")
                        {
                            // Check Licenses for Existing KMS PID
                            foreach (LicenseInstance license in licenses.GetListLicensed())
                            {
                                if (!String.IsNullOrWhiteSpace(license.KMSServerExtendedPID))
                                {
                                    // Found Existing KMS PID
                                    kmsPid = license.KMSServerExtendedPID;
                                    settings.DefaultKMSPID = license.KMSServerExtendedPID;
                                    settings.GenerateRandomKMSPID = false;
                                    break;
                                }
                            }

                            // Did Not Find Existing KMS PID
                            if (kmsPid == "ReuseKMSPID")
                            {
                                // Generate Random KMS PID
                                settings.GenerateRandomKMSPID = true;
                            }
                        }
                        else if (kmsPid == "RandomKMSPID")
                        {
                            // Generate Random KMS PID
                            settings.GenerateRandomKMSPID = true;
                        }
                        else
                        {
                            // Use Static KMS PID
                            settings.GenerateRandomKMSPID = false;

                            // Set Static KMS PID if it is not Default KMS PID
                            if (kmsPid != "DefaultKMSPID")
                            {
                                settings.DefaultKMSPID = kmsPid;
                            }
                        }

                        // Handle KMS HWID
                        settings.DefaultKMSHWID = kmsHwid;

                        // Handle Client Count
                        if (licenses is LicenseListOffice || OSVersion.IsWindowsServer())
                        {
                            settings.CurrentClientCount = 5;
                        }
                        else
                        {
                            settings.CurrentClientCount = 25;
                        }

                        // Start KMS Server
                        KMSServer.Start(null, settings);
                    }
                    catch (SocketException)
                    {
                        output.WriteLine("Failed to start KMS Emulator!");
                        output.WriteLine("KMS Port may be in use.");
                    }
                    catch (Exception ex)
                    {
                        output.WriteLine("Failed to start KMS Emulator!");
                        output.WriteLine(ex.Message);
                    }
                }

                // Attempt Activation on Each License
                if (minimalOutput == false)
                {
                    output.WriteLine("---Processing--------------------------");
                    output.WriteLine("----------------------------------------");
                }

                bool firstLine = true;
                foreach (LicenseInstance license in licenses.GetListLicensed())
                {
                    // Add Extra Line if Needed
                    if (firstLine)
                    {
                        firstLine = false;
                    }
                    else
                    {
                        output.WriteLine();
                    }

                    if (minimalOutput == false)
                    {
                        output.WriteLine("Installed product key detected - attempting to activate the following product:");
                        output.WriteLine("Name: " + license.LicenseName);
                        output.WriteLine("Description: " + license.LicenseDescription);
                        output.WriteLine("Family: " + license.LicenseFamily);
                        output.WriteLine("SKU ID: " + license.SKUID);
                        output.WriteLine("Last 5 characters of installed product key: " + license.PartialProductKey);
                    }
                    else
                    {
                        output.WriteLine("Attempting to Activate " + license.LicenseFamily);
                    }

                    using (ManagementObject classInstance = new ManagementObject("root\\CIMV2", wmiInfo + "'" + license.SKUID + "'", null))
                    {
                        string errorCode = string.Empty;
                        for (int i = 0; i < 10; i++)
                        {
                            try
                            {
                                classInstance.InvokeMethod("Activate", null, null);
                                output.WriteLine("<Product activation successful>");
                                output.Write("----------------------------------------");
                                errorCode = string.Empty;
                                break;
                            }
                            catch (COMException ex)
                            {
                                // Get Activation Error
                                errorCode = "0x" + ex.ErrorCode.ToString("X8");

                                if (errorCode == "0xC004F074" && (useDLLInjection || useTAPAdapter || useWinDivert))
                                {
                                    if (useDLLInjection)
                                    {
                                        // Kill KMS Connection Broker
                                        CommonUtilities.KillProcess("SppExtComObj");
                                    }
                                    continue;
                                }
                                break;
                            }
                        }

                        // Show Activation Error
                        if (errorCode != string.Empty)
                        {
                            output.WriteLine("ERROR CODE: " + errorCode);
                            output.WriteLine("ERROR TEXT: " + LicenseErrorCode.GetErrorDescription(errorCode));
                            if (errorCode == "0xC004F059" || errorCode == "0xC004F035")
                            {
                                output.WriteLine("WARNING: It will be impossible to activate via KMS due to OEM BIOS issues.");
                            }
                            output.WriteLine("<Product activation failed>");
                            output.Write("----------------------------------------");
                        }
                    }
                }

                // Stop LocalHost Bypass
                if (useDLLInjection)
                {
                    /*
                    if (!KMSDLLInjection.UnloadDLL())
                    {
                        output.Write(Environment.NewLine + "Failed to eject LocalHost Bypass DLL.");
                    }
                    */
                    try
                    {
                        KMSDLLInjection.UnloadDLLIFEO();
                    }
                    catch (Exception ex)
                    {
                        output.Write(Environment.NewLine + "Failed to eject LocalHost Bypass DLL.");
                        output.Write(Environment.NewLine + ex.Message);
                    }
                }
                else if (useTAPAdapter)
                {
                    KMSTAPDriver.Unload();
                }
                else if (useWinDivert)
                {
                    KMSWinDivert.StopWinDivertClient();
                }

                // Stop KMSEmulator
                if (useKMSEmulator)
                {
                    try
                    {
                        KMSServer.Stop();
                    }
                    catch (Exception ex)
                    {
                        output.WriteLine("Failed to stop KMS Emulator!");
                        output.WriteLine(ex.Message);
                    }
                }

                // Remove KMS Server
                if (removeKMSConnection)
                {
                    // Remove KMS Host
                    try
                    {
                        if (licenses is LicenseListOffice)
                        {
                            KMSConnection.RemoveKMSHostOffice(licenses);
                        }
                        else if (licenses is LicenseListWindows)
                        {
                            KMSConnection.RemoveKMSHostWindows(licenses);
                        }
                    }
                    catch (Exception ex)
                    {
                        output.WriteLine("Failed to remove KMS Host from registry!");
                        output.WriteLine(ex.Message);
                    }

                    // Remove KMS Port
                    try
                    {
                        if (licenses is LicenseListOffice)
                        {
                            KMSConnection.RemoveKMSPortOffice(licenses);
                        }
                        else if (licenses is LicenseListWindows)
                        {
                            KMSConnection.RemoveKMSPortWindows(licenses);
                        }
                    }
                    catch (Exception ex)
                    {
                        output.WriteLine("Failed to remove KMS Port from registry!");
                        output.WriteLine(ex.Message);
                    }
                }

                // Delete Allow Firewall Rules
                CommonUtilities.ExecuteCommand("netsh advfirewall firewall delete rule name=" + programName, true);

                // Read WinDivert Log
                if (useWinDivert && KMSWinDivert.Log != string.Empty)
                {
                    return KMSWinDivert.Log + output;
                }

                return output.ToString();
            }
        }

        /// <summary>
        /// View the Status and Information of all Licenses
        /// </summary>
        /// <param name="licenses">List of all LicenseInstance</param>
        /// <param name="showCMID">Display Client Machine ID Once</param>
        /// <param name="showUnlicensed">Show Licenses with no installed Product Keys</param>
        /// <returns>String Representation of all Licenses</returns>
        public static string CheckActivation(LicenseList licenses, bool showCMID = false, bool showUnlicensed = false)
        {
            using (StringWriter output = new StringWriter())
            {
                // Show Activation Errors if No Licenses or Keys Exist
                if (licenses.GetListLicensed().Count == 0 && licenses.GetListUnlicensed().Count > 0 && !showUnlicensed)
                {
                    return LicenseErrorCode.ErrKeyless;
                }
                if (licenses.GetListLicensed().Count == 0 && licenses.GetListUnlicensed().Count == 0)
                {
                    return LicenseErrorCode.ErrBroken;
                }
                // Show CMID
                if (showCMID)
                {
                    output.WriteLine("---------------------------------------");
                    output.WriteLine("CMID: " + licenses.GetCMID());
                }
                // Show Active Licenses
                foreach (LicenseInstance license in licenses.GetListLicensed())
                {
                    output.WriteLine("---------------------------------------");
                    output.WriteLine(license.ToString());
                }
                // Show Inactive Licenses
                if (showUnlicensed)
                {
                    foreach (LicenseInstance license in licenses.GetListUnlicensed())
                    {
                        output.WriteLine("---------------------------------------");
                        output.WriteLine(license.ToString());
                    }
                }
                output.Write("---------------------------------------");

                return output.ToString();
            }
        }
    }

    /// <summary>
    /// Group of Methods for Automating KMS Activation and Preparation
    /// </summary>
    public static class EZActivator
    {
        /// <summary>
        /// Install KMS Keys Automatically and Perform Activation of All KMS Licenses
        /// </summary>
        /// <param name="licenses">List of all Licenses</param>
        /// <param name="minimalOutput">Reduce the Amount of Output During Activation</param>
        /// <param name="kmsServer">KMS Host to Connect To</param>
        /// <param name="kmsPort">KMS Port to Connect To</param>
        /// <param name="kmsPid">KMS PID to apply to KMSEmulator</param>
        /// <param name="kmsHwid">KMS Hardware ID to apply to KMSEmulator</param>
        /// <param name="useKMSEmulator">Start a KMSEmulator Process</param>
        /// <param name="removeKMSConnection">Remove KMS Host and Port after Activation</param>
        /// <param name="killProcessOnPort">Force Start KMSEmulator by Killing Other Processes usign the KMS Port</param>
        /// <param name="useDLLInjection">Use DLL Injection to Force Localhost KMS Activation</param>
        /// <param name="useTAPAdapter">Use TAP Adapter to Force Localhost KMS Activation</param>
        /// <param name="useWinDivert">Use WinDivert Client to Force Localhost KMS Activation</param>
        /// <param name="localHostBypassIPAddress">IP Address of TAP Adapter NIC or WinDivert Client</param>
        /// <param name="localHostBypassIPSubnet">Subnet Mask for TAP Adapter or WinDivert Client Network</param>
        /// <returns>Activation Result of all KMS Licenses</returns>
        public static string RunEZActivator(LicenseList licenses, bool minimalOutput = true, string kmsServer = "127.0.0.2", int kmsPort = 1688, string kmsPid = "RandomKMSPID", string kmsHwid = "364F463A8863D35F", bool useKMSEmulator = true, bool removeKMSConnection = false, bool killProcessOnPort = false, bool useDLLInjection = false, bool useTAPAdapter = false, bool useWinDivert = false, string localHostBypassIPAddress = "10.3.0.1", string localHostBypassIPSubnet = "255.255.255.0")
        {
            // Show Activation Errors if No Licenses Exist
            if (licenses.GetListUnlicensed().Count == 0 && licenses.GetListLicensed().Count == 0)
            {
                return LicenseErrorCode.ErrBroken;
            }
            // Show Activation Errors if No KMS Licenses Exist
            if (licenses.GetListKMS().Count == 0)
            {
                return LicenseErrorCode.ErrNoKMS;
            }

            using (StringWriter output = new StringWriter())
            {
                // Activated Windows Check
                if (licenses is LicenseListWindows && licenses.GetListPermanentlyActivated().Count > 0)
                {
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Windows is already permanently activated.");
                    output.WriteLine("----------------------------------------");
                    return output.ToString();
                }

                // Get All Possible KMS Keys
                KeyList keys = new KeyList();
                if (licenses is LicenseListOffice)
                {
                    keys = KeyBase.GetApplicableKeysList(OfficeVersion.GetOfficeName());
                }
                else if (licenses is LicenseListWindows)
                {
                    keys = KeyBase.GetApplicableKeysList(OSVersion.GetWindowsName());
                }

                // Remove Trial Keys
                bool removedTrialKeys = false;
                output.WriteLine("----------------------------------------");
                output.WriteLine("Removing Any Trial/Grace Keys.");
                foreach (LicenseInstance licenseKeys in licenses.GetListLicensed())
                {
                    if (licenseKeys.LicenseDescription.ToUpper().Contains("TRIAL") || licenseKeys.LicenseDescription.ToUpper().Contains("GRACE"))
                    {
                        if (licenses is LicenseListOffice)
                        {
                            KeyInstaller.UnInstallKeyByKeyOffice(licenseKeys.PartialProductKey);
                            output.WriteLine("Removed Key for: " + licenseKeys.LicenseFamily + " (" + licenseKeys.PartialProductKey + ").");
                            removedTrialKeys = true;
                        }
                        else if (licenses is LicenseListWindows)
                        {
                            KeyInstaller.UnInstallKeyByKeyOffice(licenseKeys.PartialProductKey);
                            output.WriteLine("Removed Key for: " + licenseKeys.LicenseFamily + " (" + licenseKeys.PartialProductKey + ").");
                            removedTrialKeys = true;
                        }
                    }
                }
                if (removedTrialKeys)
                {
                    licenses.Refresh();
                }

                // Install Uninstalled KMS Keys by SKUID Match
                bool installedKMSKeys = false;
                output.WriteLine("----------------------------------------");
                output.WriteLine("Installing Any Matching Volume Keys.");
                if (licenses is LicenseListWindows && OSVersion.GetWindowsNumber() >= 10 && OSVersion.GetWindowsBuildNumber() >= 14393)
                {
                    // Determine KMS Key
                    if (licenses.GetListLicensed().Count > 0)
                    {
                        // SKU Match Existing Key
                        foreach (LicenseInstance license in licenses.GetListLicensed())
                        {
                            string editionId = license.LicenseFamily;
                            try
                            {
                                // Check Key
                                bool installKey = true;
                                List<string> skuidListMatched = keys.GetSKUIDs(editionId);
                                foreach (LicenseInstance licensed in licenses.GetListLicensed())
                                {
                                    if (skuidListMatched.Contains(licensed.SKUID))
                                    {
                                        installKey = false;
                                        break;
                                    }
                                }

                                // Install Key
                                if (installKey)
                                {
                                    // Get All SKUIDs
                                    List<string> skuidListAll = new List<string>();
                                    foreach (LicenseInstance licensed in licenses.GetListFull())
                                    {
                                        skuidListAll.Add(licensed.SKUID);
                                    }

                                    // Get Matched Key
                                    foreach (string skuid in skuidListMatched)
                                    {
                                        if (skuidListAll.Contains(skuid))
                                        {
                                            output.WriteLine("Installing " + keys.GetProductName(skuid) + " KMS Key (" + keys.GetProductKey(skuid) + ").");
                                            KeyInstaller.InstallKeyWindows(keys.GetProductKey(skuid));
                                            output.WriteLine("<Product key installation successful>");
                                            installedKMSKeys = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                output.WriteLine(ex.Message);
                            }
                        }
                    }
                    else
                    {
                        // Edition ID Check
                        using (RegistryKey registrySubKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", true))
                        {
                            if (registrySubKey != null && registrySubKey.GetValue("EditionID") != null)
                            {
                                string editionId = registrySubKey.GetValue("EditionID").ToString();
                                try
                                {
                                    if (keys.GetProductKey(editionId) != string.Empty)
                                    {
                                        output.WriteLine("Installing " + keys.GetProductName(editionId) + " KMS Key (" + keys.GetProductKey(editionId) + ").");
                                        KeyInstaller.InstallKeyWindows(keys.GetProductKey(editionId));
                                        output.WriteLine("<Product key installation successful>");
                                        installedKMSKeys = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    output.WriteLine(ex.Message);
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (LicenseInstance license in licenses.GetListUnlicensed())
                    {
                        if (keys.GetSKUIDs().Contains(license.SKUID))
                        {
                            try
                            {
                                output.WriteLine("Installing " + keys.GetProductName(license.SKUID) + " KMS Key (" + keys.GetProductKey(license.SKUID) + ").");
                                if (licenses is LicenseListOffice)
                                {
                                    KeyInstaller.InstallKeyOffice(keys.GetProductKey(license.SKUID));
                                    output.WriteLine("<Product key installation successful>");
                                    installedKMSKeys = true;
                                }
                                else if (licenses is LicenseListWindows)
                                {
                                    KeyInstaller.InstallKeyWindows(keys.GetProductKey(license.SKUID));
                                    output.WriteLine("<Product key installation successful>");
                                    installedKMSKeys = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                output.WriteLine(ex.Message);
                            }
                        }
                    }
                }
                output.WriteLine("----------------------------------------");
                if (installedKMSKeys)
                {
                    licenses.Refresh();
                }

                // Perform Activation
                if (licenses is LicenseListOffice)
                {
                    output.WriteLine("Attempting To Activate Microsoft Office");
                    output.Write(ActivationState.AttemptActivationOffice(licenses, minimalOutput, kmsServer, kmsPort, kmsPid, kmsHwid, useKMSEmulator, removeKMSConnection, killProcessOnPort, useDLLInjection, useTAPAdapter, useWinDivert, localHostBypassIPAddress, localHostBypassIPSubnet));
                }
                else if (licenses is LicenseListWindows)
                {
                    output.WriteLine("Attempting To Activate Microsoft Windows");
                    output.Write(ActivationState.AttemptActivationWindows(licenses, minimalOutput, kmsServer, kmsPort, kmsPid, kmsHwid, useKMSEmulator, removeKMSConnection, killProcessOnPort, useDLLInjection, useTAPAdapter, useWinDivert, localHostBypassIPAddress, localHostBypassIPSubnet));
                }

                return output.ToString();
            }
        }
    }
}