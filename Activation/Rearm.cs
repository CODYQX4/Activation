using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Common;
using License;
using ProductDetection;

namespace Activation
{
    /// <summary>
    /// Group of methods supporting Microsoft Office and Windows Rearm
    /// </summary>
    public static class RearmBase
    {
        /// <summary>
        /// Get the Lowest Remaining Grace Period in Days
        /// </summary>
        /// <param name="licenses">List of Licenses to Check</param>
        /// <returns>Number of Days Remaining for the License that will expire first</returns>
        public static int LowestGraceDays(LicenseList licenses)
        {
            // Store Lowest Grace Period
            int lowestDays = 30;

            // Get List of All Licensed Licenses
            List<LicenseInstance> licenseFilter = licenses.GetListLicensed();

            // Remove Permanently Activated Licenses
            foreach (LicenseInstance license in licenses.GetListPermanentlyActivated())
            {
                if (licenseFilter.Contains(license))
                {
                    licenseFilter.Remove(license);
                }
            }

            // Check which license has the lowest remaining Grace Period
            foreach (LicenseInstance license in licenseFilter)
            {
                if (license.RemainingGraceDays >= lowestDays)
                {
                    continue;
                }
                lowestDays = Convert.ToInt32(license.RemainingGraceDays);
            }
            return lowestDays;
        }
    }

    /// <summary>
    /// Group of Methods For Rearming Microsoft Office and Obtaining Remaining Rearm Count
    /// </summary>
    public static class RearmOffice
    {
        /// <summary>
        /// Path to OSPPREARM.exe
        /// </summary>
        private static string _ospprearm = string.Empty;

        /// <summary>
        /// Rearm Microsoft Office
        /// </summary>
        /// <returns>Output of Rearm Result and Any Errors</returns>
        public static string Rearm()
        {
            try
            {
                using (StringWriter output = new StringWriter())
                {
                    // Get Rearm Tool Path
                    if (String.IsNullOrWhiteSpace(_ospprearm))
                    {
                        _ospprearm = GetRearmToolPath();
                    }

                    // Rearm with SKUIDs if Microsoft Office uses Windows Licensing Services
                    if (OfficeVersion.IsOfficeSPP())
                    {
                        LicenseList licenses = new LicenseListOffice();
                       
                        output.WriteLine("---Processing--------------------------");
                        output.Write("----------------------------------------");
                        foreach (LicenseInstance license in licenses.GetListLicensed())
                        {
                            output.WriteLine();
                            output.WriteLine("Installed license detected - attempting to rearm the following product:");
                            output.WriteLine("Name: " + license.LicenseName);
                            output.WriteLine("Description: " + license.LicenseDescription);
                            output.WriteLine("SKU ID: " + license.SKUID);
                            Result result = CommonUtilities.ExecuteCommand(_ospprearm + " {" + license.SKUID + "}", true, true);

                            if (result.HasError)
                            {
                                // Show Rearm Error
                                string errorCode = Regex.Match(result.Error, "[0][x][A-Fa-f0-9]{8}").ToString().ToUpper().Replace("X", "x");
                                output.WriteLine("ERROR CODE: " + errorCode);
                                output.WriteLine("ERROR TEXT: " + LicenseErrorCode.GetErrorDescription(errorCode));
                                output.Write("<Product rearm failed>" + Environment.NewLine + "----------------------------------------");
                            }
                            else
                            {
                                output.Write("<Product rearm successful>" + Environment.NewLine + "----------------------------------------");
                            }
                        }
                    }
                    // Rearm without SKUIDs
                    else
                    {
                        Result result = CommonUtilities.ExecuteCommand(_ospprearm, true, true);

                        if (result.HasError)
                        {
                            output.WriteLine("<Microsoft Office rearm failed.>");
                            // Show Rearm Error
                            string errorCode = Regex.Match(result.Error, "[0][x][A-Fa-f0-9]{8}").ToString().ToUpper().Replace("X", "x");
                            output.WriteLine("ERROR CODE: " + errorCode);
                            output.WriteLine("ERROR TEXT: " + LicenseErrorCode.GetErrorDescription(errorCode));
                        }
                        else
                        {
                            output.WriteLine("<Microsoft Office rearm successful.>");
                        }
                    }
                    return output.ToString();
                }
            }
            catch (COMException ex)
            {
                return "<Microsoft Office rearm failed>" + Environment.NewLine + LicenseErrorCode.GetErrorDescription("0x" + ex.ErrorCode.ToString("X8"));
            }
        }
       
        /// <summary>
        /// Obtain the remaining number of Rearms for Microsoft Office
        /// </summary>
        /// <returns>Remaining Rearm Count Number</returns>
        public static int GetRearmCount()
        {
            // Rearm Count
            int rearmCount;

            // Microsoft Office Application ID
            Guid officeAppID = Guid.Empty;

            // Initialize Microsoft Office Application Specific Information
            if (OfficeVersion.GetOfficeNumber() == 14)
            {
                officeAppID = new Guid("59a52881-a989-479d-af46-f275c6370663");
            }
            else if (OfficeVersion.GetOfficeNumber() >= 15)
            {
                officeAppID = new Guid("0ff1ce15-a989-479d-af46-f275c6370663");
            }

            // Target the desired Software Licensing Service
            if (!OfficeVersion.IsOfficeSPP())
            {
                // Load Library
                IntPtr osppDllHandle = NativeMethods.LoadLibrary(GetOSPPCPath());

                // Handles
                IntPtr osppHandle = IntPtr.Zero;

                // Open Handle to Microsoft Office Software Licensing Service
                NativeMethods.OSPPOpen(ref osppHandle);

                // Get Microsoft Office Remaining Rearm Count from OSPP // TODO: Broken?
                uint descSize = 4;
                byte[] descBuffer = new byte[descSize];
                NativeMethods.SLDATATYPE slDataType = NativeMethods.SLDATATYPE.SL_DATA_DWORD;
                NativeMethods.OSPPGetApplicationInformation(osppHandle, ref officeAppID, "RemainingRearmCount", ref slDataType, ref descSize, ref descBuffer);
                rearmCount = descBuffer[0];

                // Close Handle to Microsoft Office Software Licensing Service
                NativeMethods.OSPPClose(osppHandle);
                NativeMethods.FreeLibrary(osppDllHandle);
            }
            else
            {
                // Handles
                IntPtr sppHandle = IntPtr.Zero;

                // Open Handle to Microsoft Windows Software Licensing Service
                NativeMethods.SPPOpen(ref sppHandle);

                // Get Microsoft Office Remaining Rearm Count from SPP
                uint descSize = 4;
                byte[] descBuffer = new byte[descSize];
                NativeMethods.SLDATATYPE slDataType = NativeMethods.SLDATATYPE.SL_DATA_DWORD;
                NativeMethods.SPPGetApplicationInformation(sppHandle, ref officeAppID, "RemainingRearmCount", ref slDataType, ref descSize, ref descBuffer);
                rearmCount = descBuffer[0];

                // Close Handle to Microsoft Windows Software Licensing Service
                NativeMethods.SPPClose(sppHandle);
            }

            // Return Rearm Count
            return rearmCount;
        }

        /// <summary>
        /// Obtain Path to OSPPREARM.exe for Microsoft Office
        /// </summary>
        /// <returns>Path to OSPPREARM.exe</returns>
        private static string GetRearmToolPath()
        {
            switch (OfficeVersion.GetOfficeNumber())
            {
                case 14:
                    switch (Architecture.GetOfficeArch())
                    {
                        case Architecture.X86:
                            {
                                return CommonUtilities.EscapePath(Environment.ExpandEnvironmentVariables("%CommonProgramFiles%\\microsoft shared\\OfficeSoftwareProtectionPlatform\\OSPPREARM.exe"));
                            }
                        case Architecture.WOW:
                            {
                                return CommonUtilities.EscapePath(Environment.ExpandEnvironmentVariables("%CommonProgramFiles(x86)%\\microsoft shared\\OfficeSoftwareProtectionPlatform\\OSPPREARM.exe"));
                            }
                        case Architecture.X64:
                            {
                                return CommonUtilities.EscapePath(Environment.ExpandEnvironmentVariables("%CommonProgramW6432%\\microsoft shared\\OfficeSoftwareProtectionPlatform\\OSPPREARM.exe"));
                            }
                        default:
                            {
                                return string.Empty;
                            }
                    }
                case 15:
                case 16:
                    return CommonUtilities.EscapePath(OfficeVersion.GetInstallationPath() + "OSPPREARM.exe");
                default:
                    throw new ApplicationException("Unsupported Microsoft Office Edition!");
            }
        }

        /// <summary>
        /// Obtain Path to OSPPC.dll for Microsoft Office
        /// </summary>
        /// <returns>Path to OSPPC.dll</returns>
        private static string GetOSPPCPath()
        {
            switch (OfficeVersion.GetOfficeNumber())
            {
                case 14:
                case 15:
                case 16:
                    switch (Architecture.GetOfficeArch())
                    {
                        case Architecture.X86:
                            {
                                return Environment.ExpandEnvironmentVariables("%CommonProgramFiles%\\microsoft shared\\OfficeSoftwareProtectionPlatform\\osppc.dll");
                            }
                        case Architecture.WOW:
                        case Architecture.X64:
                            {
                                return Environment.ExpandEnvironmentVariables("%CommonProgramW6432%\\microsoft shared\\OfficeSoftwareProtectionPlatform\\osppc.dll");
                            }
                        default:
                            {
                                return string.Empty;
                            }
                    }
                default:
                    throw new ApplicationException("Unsupported Microsoft Office Edition!");
            }
        }
    }

    /// <summary>
    /// Group of Methods For Rearming Microsoft Windows and Obtaining Remaining Rearm Count
    /// </summary>
    public static class RearmWindows
    {
        /// <summary>
        /// Obtain the remaining number of Rearms for Microsoft Windows
        /// </summary>
        /// <returns>Remaining Rearm Count Number</returns>
        public static int GetRearmCount()
        {
            int rearmCount = 0;
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\CIMV2", "SELECT RemainingWindowsReArmCount FROM SoftwareLicensingService"))
            {
                foreach (ManagementBaseObject queryObj in searcher.Get())
                {
                    rearmCount = Convert.ToInt32(queryObj["RemainingWindowsReArmCount"]);
                }
            }
            return rearmCount;
        }

        /*
        // TODO: Fix 1000=232 Bug
        /// <summary>
        /// Obtain the remaining number of Rearms for Microsoft Windows via DLL
        /// </summary>
        /// <returns>Remaining Rearm Count Number</returns>
        public static int GetRearmCountDLL()
        {
            // Microsoft Windows Application ID
            Guid windowsAppID = new Guid("55c92734-d682-4d71-983e-d6ec3f16059f");

            // Handles
            IntPtr sppHandle = IntPtr.Zero;

            // Open Handle to Microsoft Windows Software Licensing Service
            NativeMethods.SPPOpen(ref sppHandle);

            // Get Microsoft Office Remaining Rearm Count from SPP
            uint descSize = 4;
            byte[] descBuffer = new byte[descSize];
            NativeMethods.SLDATATYPE slDataType = NativeMethods.SLDATATYPE.SL_DATA_DWORD;
            NativeMethods.SPPGetApplicationInformation(sppHandle, ref windowsAppID, "RemainingRearmCount", ref slDataType, ref descSize, ref descBuffer);

            // Close Handle to Microsoft Windows Software Licensing Service
            NativeMethods.SPPClose(sppHandle);

            // Return Rearm Count
            return descBuffer[0];
        }
        */

        /// <summary>
        /// Rearm Microsoft Windows
        /// </summary>
        /// <returns>Output of Rearm Result and Any Errors</returns>
        public static string Rearm()
        {
            try
            {
                // Perform Rearm
                using (ManagementObject classInstance = new ManagementObject(@"root\CIMV2", "SoftwareLicensingService.Version='" + OSVersion.GetSPPSVCVersion() + "'", null))
                {
                    classInstance.InvokeMethod("ReArmWindows", null, null);
                }

                // Restart SPPSVC
                Services.StopSPPSVC();
                Services.StartSPPSVC();

                // On Windows 7 and Earlier, we don't need to reboot.
                if (OSVersion.GetWindowsNumber() < 6.2)
                {
                    return "<Microsoft Windows rearm successful.>";
                }
                return "<Microsoft Windows rearm successful.>" + Environment.NewLine + "You need to reboot for the rearm to take effect or Windows to work properly.";
            }
            catch (COMException ex)
            {
                return "<Microsoft Windows rearm failed>" + Environment.NewLine + LicenseErrorCode.GetErrorDescription("0x" + ex.ErrorCode.ToString("X8"));
            }
        }
    }
}