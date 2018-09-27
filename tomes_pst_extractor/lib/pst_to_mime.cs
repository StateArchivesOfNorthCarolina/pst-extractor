/* pst_to_mime.cs

Author: Jeremy Gibson and Nitin Arora, State Archives of North Carolina.
Version: 0.0.2

This application uses MailBee .NET objects library (https://afterlogic.com/mailbee-net/email-components) 
to extract EML files a source PST file.

Note: This code requires a MailBee license key (look for "string licenseKey =", below).

*** WARNING!! *** You MUST NOT distribute source code with a valid license key value. But you can
distribute a compiled executable per the MailBee documentation 
(https://afterlogic.com/mailbee-net/docs/keys.html):
    "If you redistribute the application to third-parties without the source code, embedding the
    key in the code may be the preferred method of hiding the license key from end users."

Features:

    - Skips processing of folders that start with the string "delete" (case insensitive).
    - Names EML files with the ItemId from the PST's Messaging Application Programming Interface (MAPI).
    - Legalizes output folder/file path names.

Logging notes:
    - All output should start with one of the following strings: "DEBUG: ", "INFO: ", "WARNING: ", 
    "ERROR: " , or "CRITICAL: ".
    - This will allow the Python wrapper ("../pst_extractor.py") to interpret the correct logging
    level.
    - The only exception to this is "INFO: " in that omitting this string from "info"-level output
    will imply the output to be at the "info"-level.
*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using MailBee;

namespace TOMES_PST_Extractor {

    class TOMES_PST_Extractor {

        static void Main(string[] args) {
            /* Parses positional command line arguments and converts a PST file to MIME.

            Args:
                - accountName: An identifier for the email account.
                - pstFile: The PST file to convert.
                - outputPath: The directory in which to place the MIME data.

            Example:
                `[mono] pst_to_mime.exe sample ../../tests/sample_files/sample.pst ../../tests/sample_files`
            */

            // !!! DO NOT DISTRIBUTE SOURCE CODE WITH THIS VALUE FILLED OUT. !!!
            string licenseKey = "License_Key_Goes_Here";
            
            // set MailBee license key.
            MailBee.Global.LicenseKey = licenseKey;
            
            // get command line arguments.
            if (args.Length < 3) {
                string err = "Missing required argument(s).";
                Console.WriteLine("ERROR: " + err);
                System.Environment.Exit(1);
            }
            string accountName = args[0];
            string pstFile = args[1];
            string outputPath = args[2];
            
            // set EML output folder.
            string mimePath = outputPath + "/" + accountName;
            Console.WriteLine("DEBUG: Setting MIME folder: " + mimePath);
            
            // create a MailBee PST object and folder collection; extract PST to EML.
            Console.WriteLine("Extracting MIME data from PST file: " + pstFile);
            try {
                MailBee.Outlook.PstReader pstObject = new MailBee.Outlook.PstReader(pstFile);
                MailBee.Outlook.PstFolderCollection folderObjects = 
                    pstObject.GetPstRootFolders(true);
                List<int> foldersToProcess = GetFoldersToProcess(folderObjects);
                ProcessFolders(mimePath, foldersToProcess, folderObjects);
                System.Environment.Exit(0);
            }
            catch (Exception error) {
                Console.WriteLine("ERROR: " + error.Message);
                System.Environment.Exit(error.HResult);
            }
        }

        private static List<int> GetFoldersToProcess(MailBee.Outlook.PstFolderCollection 
			folderObjects) {
            /* Loops through folder names in the PST file and determines which folders to process.

            Args:
                - folderObjects: The MailBee PST folder objects.
            
            Returns:
                 list: Each item is an identifier for each folder to process.
            */
            
            Console.WriteLine("Collecting folders to process.");
            
            List<int> foldersToProcess = new List<int>();
            foreach (MailBee.Outlook.PstFolder folder in folderObjects) {

                string folderName = NormalizeFolderName(folder.Name);
                
                // skip folders starting with "delete" (case insensitive).
                bool skipFolder = folder.SafeShortName.ToLower().StartsWith("delete");
                if (skipFolder == true) {
                    Console.WriteLine("Skipping message folder: " + folderName);
                }
                else {
                    Console.WriteLine("Adding message folder: " + folderName);
                    foldersToProcess.Add(folder.PstID);
                }
            }

            return foldersToProcess;
        }

        private static void ProcessFolders(string mimePath, List<int> foldersToProcess, 
			MailBee.Outlook.PstFolderCollection folderObjects) {
            /* Loops through the PST @folderObjects and extracts messages if the PST folder is in 
            @foldersToProcess.

            Args:
                - mimePath: The folder in which to extract the PST.
                - foldersToProcess: A list of identifiers of folders to process.
                - folderObjects: The MailBee PST folder objects.
            
            Returns:
                null
            */

            Console.WriteLine("Processing folders.");
            
            // loop through each PST folder.
            int totalFromPST = 0;
            foreach (MailBee.Outlook.PstFolder folder in folderObjects) {

                // skip folders not in @foldersToProcess.
                if (foldersToProcess.Contains(folder.PstID) == false) {
                    continue;
                }

                // normalize the folder name; create the output path.
                string folderName = NormalizeFolderName(folder.Name);
                string folderPath = mimePath + "/" + folderName;
                Console.WriteLine("Processing folder: " + folderPath);
                Directory.CreateDirectory(folderPath);
                
                // extract messages from the @folder and write them to @folderPath.
                try {
                    totalFromPST += ProcessMessages(folderPath, folder);
                }
                catch (Exception error) {
                    Console.WriteLine("ERROR: " + error.Message);
                } 
            }

            Console.WriteLine("Total messages extracted from PST: " + totalFromPST);
            return;
        }

        private static int ProcessMessages (string folderPath, MailBee.Outlook.PstFolder folder) {
            /* Extracts messages from the PST @folder and writes them to @folderPath.

            Args:
                - folderPath: The path into which to extract a PST folder's messages.
                - folder: The MailBee PST folder object.
            
            Returns
                int: The number of messages extracted from @folder.
            */
            
            Console.WriteLine("Total messages being processed: " + folder.Items.Count);

            // loop through messages and extract them.
            int totalFromFolder = 0;
            foreach (MailBee.Outlook.PstItem item in folder.Items) {

                // skip non-messages; see: https://docs.microsoft.com/en-us/office/vba/outlook/concepts/forms/item-types-and-message-classes
                string messageClass = (string)item.PstFields["MessageClass"];
                if (messageClass != "IPM.Note") {
                    Console.WriteLine("Omitting non-message item with id: " + item.PstID);
                    continue;
                }

                // extract the message and write it to file as an EML.
                MailBee.Mime.MailMessage message = item.GetAsMailMessage();
                string emlPath = folderPath + "/" + item.PstID + ".eml";
                Console.WriteLine("Writing EML file: " + emlPath);
                message.SaveMessage(emlPath);
                totalFromFolder++;
            }

            Console.WriteLine("Total messages extracted from folder: " + totalFromFolder);
            return totalFromFolder;
        }
    
        private static string NormalizeFolderName(string folderName) {
            /* Replaces invalid path characters with an underscore and replaces back slashes with
            forward slashes.
            
            Args:
                - folderName: The folder path to normalize.
                
            Returns:
                string: The normalized folder name.
            */

            // explicitly replace backslashes with forward slashes.
            // Note: backslashes create odd folder name characters in Linux; also forward slashes
            // are legitimate separators in Windows.
            folderName = folderName.Replace(@"\", "/");

            // replace invalid path characters with an underscore.
            // based on: https://stackoverflow.com/a/23182807
            folderName = string.Join("_", folderName.Split(Path.GetInvalidPathChars()));
            
            return folderName;
        }
    }
}
