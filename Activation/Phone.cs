using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using License;
using ProductDetection;

namespace Activation
{
    /// <summary>
    /// Methods for Performing Phone Activation and reusing previous Activation Codes
    /// </summary>
    public static class Phone
    {
        /// <summary>
        /// Name of the File to Save Valid Confirmation IDs to
        /// </summary>
        public const string PhoneFileName = "PhoneActivation.xml";

        /// <summary>
        /// Check all Licenses for eligible Microsoft Office Phone Activation and prompt for the Activation Code
        /// </summary>
        /// <param name="licenses">List of All Microsoft Office Licenses</param>
        /// <param name="window">Active GUI Window</param>
        /// <returns>Activation Result of all Microsoft Office Licenses</returns>
        public static string PhoneActivationQueryOffice(LicenseList licenses, IWin32Window window)
        {
            if (licenses is LicenseListWindows)
            {
                throw new ApplicationException("Cannot perform Microsoft Office Phone Activation using Microsoft Windows LicenseList!");
            }
            return PhoneActivationQuery(licenses, window);
        }

        /// <summary>
        /// Check all Licenses for eligible Microsoft Windows Phone Activation and prompt for the Activation Code
        /// </summary>
        /// <param name="licenses">List of All Microsoft Windows Licenses</param>
        /// <param name="window">Active GUI Window</param>
        /// <returns>Activation Result of all Microsoft Windows Licenses</returns>
        public static string PhoneActivationQueryWindows(LicenseList licenses, IWin32Window window)
        {
            if (licenses is LicenseListOffice)
            {
                throw new ApplicationException("Cannot perform Microsoft Windows Phone Activation using Microsoft Office LicenseList!");
            }
            return PhoneActivationQuery(licenses, window);
        }

        /// <summary>
        /// Check all Licenses for eligible Phone Activation and prompt for the Activation Code
        /// </summary>
        /// <param name="licenses">List of All Licenses</param>
        /// <param name="window">Active GUI Window</param>
        /// <returns>Activation Result of all Licenses</returns>
        private static string PhoneActivationQuery(LicenseList licenses, IWin32Window window)
        {
            using (StringWriter output = new StringWriter())
            {
                // Get Unactivated Licenses
                List<LicenseInstance> licenseFilter = licenses.GetListUnactivated();

                // Remove KMS Licenses
                foreach (LicenseInstance license in licenses.GetListKMS())
                {
                    if (licenseFilter.Contains(license))
                    {
                        licenseFilter.Remove(license);
                    }
                }

                // Remove Unlicensed Licenses
                foreach (LicenseInstance license in licenses.GetListUnlicensed())
                {
                    if (licenseFilter.Contains(license))
                    {
                        licenseFilter.Remove(license);
                    }
                }

                // Activate Each License and Save ConfirmationID
                output.WriteLine("---Processing--------------------------");
                output.Write("----------------------------------------");

                // Return if No Licenses are Applicable
                if (licenseFilter.Count == 0)
                {
                    output.WriteLine();
                    output.Write("<No applicable products detected>" + Environment.NewLine + "----------------------------------------");
                    return output.ToString();
                }

                // Phone Activate All Licenses
                foreach (LicenseInstance license in licenseFilter)
                {
                    // Set Activation Success Flag
                    bool success = false;

                    // Get Confirmation ID
                    string currentCID = AskForConfirmationID(license.OfflineInstallationID, window);
                    if (String.IsNullOrWhiteSpace(currentCID))
                    {
                        output.WriteLine();
                        output.Write("Blank or invalid Confirmation ID!" + Environment.NewLine + "----------------------------------------");
                        continue;
                    }

                    output.WriteLine();
                    output.WriteLine("Installed product key detected - attempting to activate the following product:");
                    output.WriteLine("Name: " + license.LicenseName);
                    output.WriteLine("Description: " + license.LicenseDescription);
                    output.WriteLine("SKU ID: " + license.SKUID);
                    output.WriteLine("Last 5 characters of installed product key: " + license.PartialProductKey);

                    if (licenses is LicenseListOffice)
                    {
                        success = PhoneActivationOffice(currentCID, license.OfflineInstallationID, license.SKUID);
                    }
                    else if (licenses is LicenseListWindows)
                    {
                        success = PhoneActivationWindows(currentCID, license.OfflineInstallationID, license.SKUID);
                    }

                    // Check if Successful Activation
                    if (success)
                    {
                        SaveConfirmationID(license.OfflineInstallationID, currentCID);
                        output.Write("<Product activation successful>" + Environment.NewLine + "----------------------------------------");
                    }
                    else
                    {
                        output.Write("<Product activation failed>" + Environment.NewLine + "----------------------------------------");
                    }
                }

                return output.ToString();
            }
        }

        /// <summary>
        /// Attempt Phone Activation on an individual Microsoft Office License
        /// </summary>
        /// <param name="confirmationID">Activation Code obtained from Phone Activation</param>
        /// <param name="offlineInstallationID">OfflineInstallationID of the License</param>
        /// <param name="skuid">SKUID of the License</param>
        /// <returns>True if Activation was successfull, False if it failed</returns>
        private static bool PhoneActivationOffice(string confirmationID, string offlineInstallationID, string skuid)
        {
            // Call Windows if Microsoft Office uses Windows Licensing Services
            if (OfficeVersion.IsOfficeSPP())
            {
                return PhoneActivationWindows(confirmationID, offlineInstallationID, skuid);
            }
            return PhoneActivation(confirmationID, offlineInstallationID, "OfficeSoftwareProtectionProduct.ID=" + "'" + skuid + "'", "SELECT ID, LicenseStatus FROM OfficeSoftwareProtectionProduct", skuid);
        }

        /// <summary>
        /// Attempt Phone Activation on an individual Microsoft Windows License
        /// </summary>
        /// <param name="confirmationID">Activation Code obtained from Phone Activation</param>
        /// <param name="offlineInstallationID">OfflineInstallationID of the License</param>
        /// <param name="skuid">SKUID of the License</param>
        /// <returns>True if Activation was successfull, False if it failed</returns>
        private static bool PhoneActivationWindows(string confirmationID, string offlineInstallationID, string skuid)
        {
            return PhoneActivation(confirmationID, offlineInstallationID, "SoftwareLicensingProduct.ID=" + "'" + skuid + "'", "SELECT ID, LicenseStatus FROM SoftwareLicensingProduct", skuid);
        }

        /// <summary>
        /// Attempt Phone Activation on an individual Microsoft Office License
        /// </summary>
        /// <param name="confirmationID">Activation Code obtained from Phone Activation</param>
        /// <param name="offlineInstallationID">OfflineInstallationID of the License</param>
        /// <param name="wmiInfo1">WMI Provider and associated data to attempt Phone Activation</param>
        /// <param name="wmiInfo2">WMI Provider and associated data to get all license instances</param>
        /// <param name="skuid">SKUID of the License</param>
        /// <returns>True if Activation was successful, False if it failed</returns>
        private static bool PhoneActivation(string confirmationID, string offlineInstallationID, string wmiInfo1, string wmiInfo2, string skuid)
        {
            // Attempt Activation
            using (ManagementObject classInstance = new ManagementObject(@"root\CIMV2", wmiInfo1, null))
            {
                using (ManagementBaseObject inParams = classInstance.GetMethodParameters("DepositOfflineConfirmationId"))
                {
                    inParams["ConfirmationId"] = confirmationID;
                    inParams["InstallationId"] = offlineInstallationID;

                    // Execute the method and obtain the return values.
                    classInstance.InvokeMethod("DepositOfflineConfirmationId", inParams, null);
                }
            }

            // Check Successful Activation
            int licStatus = 0;
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\CIMV2", wmiInfo2))
            {
                foreach (ManagementBaseObject queryObj in searcher.Get())
                {
                    if (skuid != Convert.ToString(queryObj["ID"]))
                    {
                        continue;
                    }
                    licStatus = Convert.ToInt32(queryObj["LicenseStatus"]);
                    break;
                }
            }

            // 1 is Activated
            if (licStatus != 1)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Obtain a working Confirmation ID for a given OfflineInstallation by prompting the user or reusing a previously used Confirmation ID
        /// </summary>
        /// <param name="oid">OfflineInstallationID to obtain a Confirmation ID for</param>
        /// <param name="window">Active GUI Window</param>
        /// <returns>Confirmation ID that will Activate the given OfflineInstallationID</returns>
        private static string AskForConfirmationID(string oid, IWin32Window window)
        {
            // Attempt to Load Saved CID
            string cid = LoadConfirmationID(oid);

            // Get New CID
            if (String.IsNullOrWhiteSpace(cid))
            {
                // Show CID Prompt and Get From Form
                using (AskForCID askForCID = new AskForCID(oid))
                {
                    // TODO Fix Threading
                    askForCID.ShowDialog(window);   
                    cid = AskForCID.ConfirmationID;
                }
            }
            return cid;
        }

        /// <summary>
        /// Load a previously used and valid Confirmation ID
        /// </summary>
        /// <param name="oid">OfflineInstallationID to Activate</param>
        /// <returns>Confirmation ID that will Activate the given OfflineInstallationID</returns>
        private static string LoadConfirmationID(string oid)
        {
            // Open an XMLReader on the XML File
            string fileName = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + PhoneFileName;
            if (File.Exists(fileName) == false)
            {
                using (XmlWriter writer = XmlWriter.Create(fileName))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("PhoneActivationEntries");
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
            XmlDocument phoneDoc = new XmlDocument();
            phoneDoc.Load(fileName);

            // Get a list of all XmlNodes with the Tag PhoneActivationEntry
            foreach (XmlNode node in phoneDoc.GetElementsByTagName("PhoneActivationEntry"))
            {
                if (node.FirstChild.InnerText == oid)
                {
                    return node.LastChild.InnerText;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Save a working Confirmation ID for later reactivation
        /// </summary>
        /// <param name="oid">OfflineInstallationID that was Activated</param>
        /// <param name="cid">Valid Confirmation ID that Activated the given OfflineInstallationID</param>
        private static void SaveConfirmationID(string oid, string cid)
        {
            // Open a FileStream on the XML File
            using (FileStream phoneFile = new FileStream(Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + PhoneFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                // Load the XML Document
                XmlDocument phoneDoc = new XmlDocument();
                phoneDoc.Load(phoneFile);

                // Get a list of all XmlNodes with the TAG PhoneActivationEntry
                XmlNodeList idList = phoneDoc.GetElementsByTagName("PhoneActivationEntry");

                foreach (XmlNode node in idList)
                {
                    if (node.FirstChild.InnerText == oid)
                    {
                        return;
                    }
                }

                // Create a new XMLElement
                XmlElement newEntry = phoneDoc.CreateElement("PhoneActivationEntry");

                // Create a new XMLElement
                XmlElement oidElement = phoneDoc.CreateElement("OfflineInstallationID");
                oidElement.InnerText = oid;
                // Append the element as an child
                newEntry.AppendChild(oidElement);

                // Create a new XMLElement
                XmlElement cidElement = phoneDoc.CreateElement("ConfirmationID");
                cidElement.InnerText = cid;
                // Append the element as an child
                newEntry.AppendChild(cidElement);

                // Insert the newly created XMLElement into the XMLDocument before the LastChild
                if (phoneDoc.DocumentElement != null)
                {
                    phoneDoc.DocumentElement.InsertBefore(newEntry, phoneDoc.DocumentElement.LastChild);
                }

                // Save the Document
                phoneFile.Position = 0;
                phoneDoc.Save(phoneFile);
            }
        }

        /// <summary>
        /// Check if a Confirmation ID is in valid format
        /// </summary>
        /// <param name="cid">Confirmation ID to validate</param>
        /// <returns>True if the Confirmation ID is Valid, False if it is invalid</returns>
        public static bool IsValidCID(string cid)
        {
            Regex validCID = new Regex("^([0-9]{6})-([0-9]{6})-([0-9]{6})-([0-9]{6})-([0-9]{6})-([0-9]{6})-([0-9]{6})-([0-9]{6})$");
            return validCID.Match(cid).Success;
        }
    }
}