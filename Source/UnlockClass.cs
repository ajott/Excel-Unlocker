using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Xml;

namespace ExcelUnlockerVisual {

    class UnlockClass {

        public static int Unlock(string filePath, bool overWrite, bool unlockVBA, IProgress<int> progress, IProgress<int>consoleProg) {
            // Initial variable declarations
            string homeDirectory = "";
            string hopefulNewName = "";
            int currentProgress = 0;
            int worksheetProgress = 0;

            // Gets the filename from the provided file path
            string fileName = Path.GetFileName(filePath);

            // Creates what we hope will be the new name of the file
            if (overWrite) {
                hopefulNewName = fileName;
            } else {
                hopefulNewName = "unlocked_" + fileName;
            }

            // Saves the home directory to a string
            homeDirectory = Path.GetDirectoryName(filePath);


            // Creates a GUID for the temporary directory, and creates C:\Temp\... filepath
            string directoryID = Guid.NewGuid().ToString();
            string directoryPath = "C:\\Temp\\" + directoryID;
            currentProgress += 5;
            progress.Report(currentProgress);


            // Gets the file extension from the filename
            string fileExtension = Path.GetExtension(filePath);
            

            // Uses the C:\Temp\... path from earlier, and makes the working directory
            Directory.CreateDirectory(directoryPath);
            currentProgress += 5;
            progress.Report(currentProgress);

            // Copies the workbook to the working directory as a .zip
            File.Copy(filePath, directoryPath + "\\workBook.zip");
            currentProgress += 5;
            progress.Report(currentProgress);

            // Extracts the .zip contents
            ZipFile.ExtractToDirectory(directoryPath + "\\workBook.zip", directoryPath + "\\workBook");
            currentProgress += 5;
            progress.Report(currentProgress);

            if (unlockVBA) {
                if (File.Exists(directoryPath + "\\workBook\\xl\\vbaProject.bin")) {
                    try {
                        // Reads vbaProject.bin
                        byte[] buf = File.ReadAllBytes(directoryPath + "\\workBook\\xl\\vbaProject.bin");

                        // Encodes the binary as hex, which allows us to edit the three hex couplets below
                        var str = new SoapHexBinary(buf).ToString();


                        // Find the VBA protection key ("DPB") and replaces it with a nothing key ("DBx")
                        // This causes the VBA editor to go through recovery and delete protection
                        str = str.Replace("445042", "444278");

                        // Writes the hex back to binary into the vbaProject.bin file
                        File.WriteAllBytes(directoryPath + "\\workBook\\xl\\vbaProject.bin", SoapHexBinary.Parse(str).Value);
                    }
                    catch {
                        return 2;
                    }
                }
            }
            currentProgress += 10;
            progress.Report(currentProgress);

            // Removes workbook-level structure protection
            RemoveWorkbookProtection(directoryPath + "\\workBook\\xl\\workbook.xml");

            // Searches in the decompiled workbook's xl directory for worksheets, and
            string[] worksheets = Directory.GetFiles(directoryPath + "\\workBook\\xl\\worksheets");

            worksheetProgress = 50 / worksheets.Length;

            // Calls the RemoveSheetProtection method for each of them
            foreach (string worksheet in worksheets) {
                RemoveSheetProtection(worksheet);
                currentProgress += worksheetProgress;
                progress.Report(currentProgress);
            }            

            // Recompiles the workbook with the newly unprotected sheets
            ZipFile.CreateFromDirectory(directoryPath + "\\workBook", directoryPath + "\\Book1Mod.zip");
            currentProgress += 10;
            progress.Report(currentProgress);

            // Copies the .zip back as an Excel workbook to the home directory
            try {
                File.Copy(directoryPath + "\\Book1Mod.zip", homeDirectory + "\\" + hopefulNewName, overWrite);
                progress.Report(100);
                consoleProg.Report(0);
            }
            catch {
                Directory.Delete(directoryPath, true);
                progress.Report(100);
                consoleProg.Report(1);
                return 1;
            }
            

            // Deletes the C:\Temp\... directory.
            Directory.Delete(directoryPath, true);

            return 0;

        }

        static void RemoveSheetProtection(string fileName) {
            // Initializes our document, and loads it from the filename
            var doc = new System.Xml.XmlDocument();
            doc.Load(fileName);

            XmlNodeList protections = doc.GetElementsByTagName("sheetProtection");
            try {
                foreach (XmlNode element in protections) {
                    element.ParentNode.RemoveChild(element);
                }
            }
            catch {

            }
            doc.Save(fileName);
        }

        static void RemoveWorkbookProtection(string fileName) {
            var docWorkbook = new System.Xml.XmlDocument();
            docWorkbook.Load(fileName);
            XmlNodeList protections = docWorkbook.GetElementsByTagName("workbookProtection");
            try {
                foreach (XmlNode element in protections) {
                    element.ParentNode.RemoveChild(element);
                }
            }
            catch {

            }
            docWorkbook.Save(fileName);
        }

        static int IndexOfValidFileName(string[] args) {
            foreach (string argument in args) {
                if (Path.GetExtension(argument) == ".xlsx" || Path.GetExtension(argument) == ".xlsm" || Path.GetExtension(argument) == ".xlam") {
                    return Array.IndexOf(args, argument);
                }
            }
            return -1;
        }
    }

}
