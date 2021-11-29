using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ExcelDataReader;


namespace FormatFile.Common
{
    public class GenerateXML
    {
        public static bool HandleFile(string file, string delimiter, bool normalize, bool removeoriginalfile, int headerrow, bool genericcolumnname, string suffix, bool autoheader, string tablename, string procedurename)
        {
            string FileNameExtension = Path.GetExtension(file);
            string FileName = Path.GetFileNameWithoutExtension(file);
            string FullPath = Path.GetDirectoryName(file);

            Console.WriteLine("FileName: {0}", FileName);
            Console.WriteLine("File-type: {0}", FileNameExtension);
            Console.WriteLine("Full path: {0}", FullPath);

            if (FileNameExtension == ".xls" || FileNameExtension == ".xlsx")
            {
                Console.WriteLine("File is Excel-format - converting Excel to CSV");
                string csvoutput = Path.ChangeExtension(file, ".csv");
                delimiter = "|";
                ExcelFileHelper.SaveAsCsv(file, csvoutput, delimiter);
                if (removeoriginalfile)
                {
                    Console.WriteLine("Deletes Excel file");
                    File.Delete(file);
                }
            }

            if (delimiter == "tab")
            {
                Console.WriteLine("Delimter set to TAB - converting TAB to pipe '|'");
                string text = "";
                using (StreamReader sr = new StreamReader(file))
                {
                    int i = 0;
                    do
                    {
                        i++;
                        string line = sr.ReadLine();
                        if (line != "")
                        {
                            line = line.Replace("\t", "|");
                            text = text + line + Environment.NewLine;
                        }
                    } while (sr.EndOfStream == false);
                }
                //change delimiter and filename
                delimiter = "|";
                if (!removeoriginalfile)
                {
                    FileName += "_tabConverted";
                    Console.WriteLine("New filename: {0}", FileName);
                }
                File.WriteAllText(FullPath + "\\" + FileName + ".csv", text);
            }

            //if (normalize)
            //{
            //    int r = 0;
            //    string headerline = File.ReadLines(FullPath + "\\" + FileName + ".csv").First(); // gets the first line from file.
            //    int delimitercount = (headerline.Length - headerline.Replace(delimiter, "").Length);
            //    Console.WriteLine("Normalizing file");
            //    string text = "";
            //    using (StreamReader sr = new StreamReader(FullPath + "\\" + FileName + ".csv"))
            //    {
            //        int i = 0;
            //        do
            //        {
            //            i++;
            //            string line = sr.ReadLine();
            //            if (line != "")
            //            {
            //                int linedelimiters = line.Length - line.Replace(delimiter, "").Length;
            //                if (linedelimiters < delimitercount)
            //                {
            //                    r++;
            //                    int missingdelimiters = delimitercount - linedelimiters;
            //                    line += string.Concat(Enumerable.Repeat(delimiter, missingdelimiters));
            //                }
            //                text = text + line + Environment.NewLine;
            //            }
            //        } while (sr.EndOfStream == false);
            //    }
            //    if (r > 0)
            //    {
            //        Console.WriteLine("Found {0} rows to be normalized", r);
            //        File.WriteAllText(FullPath + "\\" + FileName + ".csv", text);
            //    }
            //}


            Console.WriteLine("Processing CSV-file and generating XML-formatfile");
            var doc = GenerateXML.GenerateXMLFile(FullPath, "\\" + FileName + ".csv", delimiter, normalize, headerrow, genericcolumnname, suffix, autoheader);
            StreamWriter xmlfile = File.CreateText(FullPath + "\\formatfile_" + FileName + ".xml");
            doc.Save(xmlfile);
            xmlfile.Close();

            if (GenerateFormatFile.Properties.Settings.Default.loadToSQL)
            {
                Console.WriteLine("LoadToSQL is activated - loading file to SQL Server");
                LoadToSQL.CSVFile(
                    GenerateFormatFile.Properties.Settings.Default.servername
                    , GenerateFormatFile.Properties.Settings.Default.databasename
                    , GenerateFormatFile.Properties.Settings.Default.username
                    , GenerateFormatFile.Properties.Settings.Default.password
                    , string.IsNullOrEmpty(suffix) ? FullPath + "\\" + FileName + ".csv" : FullPath + "\\" + FileName + (delimiter == "tab" ? "_tabConverted" : "") + suffix + ".csv"
                    , FullPath + "\\formatfile_" + FileName + ".xml"
                    , tablename
                    , procedurename);
            }

            //Console.WriteLine("Delete CSV file");
            //File.Delete(FullPath + "\\" + FileName + ".csv");
            //Console.WriteLine("Delete Format-file");
            //File.Delete(FullPath + "\\formatfile_" + FileName + ".xml");

            return true;
        }

        public static XDocument GenerateXMLFile(string path, string filename, string delimiter, bool normalize, int headerrow, bool genericcolumnname, string suffix, bool autoheader)
        {
            var reader = new StreamReader(File.OpenRead(path + filename));
            //string[] values = reader.ReadLine().Replace(@"""", @"").Split(';');
            string[] values = new string[0];
            string[] valuestmp = new string[0];
            int ColumnCount = 0;


            if (!normalize || !autoheader)
            {
                for (int i = 1; i <= headerrow; i++)
                {
                    values = reader.ReadLine().Replace(@"""", @"").Split(delimiter.ToCharArray()[0]);
                }
            }
            else
            {
                string pattern = "/[^,]|[^(?:!(\".*?\"(,(?!$))?)))]/*,|(\".*?\"(,(?!$))?)";

                while (reader.Peek() >= 0)
                {
                    //valuestmp = reader.ReadLine().Replace(@"""", @"").Split(delimiter.ToCharArray()[0]);
                    valuestmp = Regex.Matches(reader.ReadLine() + ",", pattern).Cast<Match>().Select(m => m.Value).ToArray();

                    if (ColumnCount >= valuestmp.Length)
                    {
                        continue;
                    }
                    values = valuestmp;
                    ColumnCount = values.Length;
                }
            }

            ColumnCount = values.Length;

            var ns = XNamespace.Get("http://schemas.microsoft.com/sqlserver/2004/bulkload/format");
            var nsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");

            XDocument doc =
                  new XDocument(
                    new XElement(ns + "BCPFORMAT",
                        new XAttribute(XNamespace.Xmlns + "xsi", nsi),
                            new XElement("RECORD",
                            values.Select((v, index) =>
                                new XElement("FIELD",
                                    new XAttribute("ID", index + 1),
                                    new XAttribute(nsi + "type", "CharTerm"),
                                    new XAttribute("TERMINATOR", index == values.Length - 1 ? "\\r\\n" : delimiter),
                                    new XAttribute("MAX_LENGTH", "510"),
                                    new XAttribute("COLLATION", "")
                                    )
                                )
                            )
                        , new XElement("ROW",
                            values.Select((v, index) =>
                                new XElement("COLUMN",
                                    new XAttribute("SOURCE", index + 1),
                                    new XAttribute("NAME", string.IsNullOrEmpty((v.ToString()).Trim()) | genericcolumnname ? "C" + (index + 1) : (v.ToString()).Trim()),
                                    new XAttribute(nsi + "type", "SQLNVARCHAR")
                                )
                            )
                        )
                    )
                );
            reader.Close();

            if (normalize)
            {
                Console.WriteLine("Normalizing file");

                int counter = 0;
                bool overwrite = string.IsNullOrEmpty(suffix);
                string guid = System.Guid.NewGuid().ToString().Replace("X", "");
                suffix = overwrite ? guid : suffix;
                //Note: path has not tail \, filename has leading \
                string tempFile = "\\" + System.IO.Path.GetFileNameWithoutExtension(path + filename) + suffix + System.IO.Path.GetExtension(path + filename);

                using (StreamWriter writer = new StreamWriter(path + tempFile))
                {
                    reader = new StreamReader(File.OpenRead(path + filename));
                    string lineToWrite = null;
                    string pattern = delimiter + "|(\".*? \"(" + delimiter + "(?!$))?)"; // delimiter;
                    int appendCount = 0;
                    Regex rgx = new Regex(pattern);

                    while (reader.Peek() >=0)
                    {
                      lineToWrite = reader.ReadLine();
                      if (counter < headerrow || autoheader)
                        {
                            appendCount = rgx.Matches(lineToWrite).Count;

                            if (appendCount < ColumnCount - 1)
                            {
                                lineToWrite += new string(delimiter.ToCharArray()[0], ColumnCount - appendCount - 1);
                            }
                        }

                        writer.WriteLine(lineToWrite);
                        counter++;
                    }
                }

                if (overwrite)
                {
                    System.IO.File.Delete(path + filename);
                    System.IO.File.Move(path + tempFile, path + filename);
                }
            }
                
            return doc;
        }
    }
    public class ExcelFileHelper
    {
        public static bool SaveAsCsv(string excelFilePath, string destinationCsvFilePath, string delimiter)
        {

            using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                IExcelDataReader reader = null;
                if (excelFilePath.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (excelFilePath.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                if (reader == null)
                    return false;

                var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                });

                var csvContent = string.Empty;
                int row_no = 0;
                while (row_no < ds.Tables[0].Rows.Count)
                {
                    var arr = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        arr.Add(Regex.Replace(ds.Tables[0].Rows[row_no][i].ToString(), @"(\n)", string.Empty));
                    }
                    row_no++;
                    csvContent += string.Join(delimiter, arr) + "\r\n";
                }
                StreamWriter csv = new StreamWriter(destinationCsvFilePath, false);
                csv.Write(csvContent);
                csv.Close();
                csv.Dispose();
                return true;
            }
        }
    }

    public class LoadToSQL
    {
        public static bool CSVFile(string servername, string database, string username, string password, string filename, string formatfile, string tablename, string procedurename)
        {
            if (!string.IsNullOrEmpty(procedurename))
            {
                string connectionstring = @"server=" + servername + ";Database=" + database + ";";
                if (string.IsNullOrEmpty(password))
                {
                    connectionstring += "Integrated Security = true";
                }
                else
                {
                    connectionstring += "User=" + username + ";Password=" + password;
                }

                string query = @"exec " + procedurename + " '" + filename + "', '" + formatfile + "';";
                Console.WriteLine(query);

                using (SqlConnection connection = new SqlConnection(connectionstring))
                {
                    SqlCommand command = new SqlCommand(query, connection);
                    connection.Open();
                    command.ExecuteReader();
                    connection.Close();
                }
            }
            else
            {
                string command = "bcp.exe \"" + tablename + "\" in \"" + filename + "\" -S " + servername;
                if (string.IsNullOrEmpty(password))
                {
                    command += " -T";
                }
                else
                {
                    command += " -U " + username + " -P " + password;
                }
                command += " -d \"" + database + "\" -f \"" + formatfile + "\"";

                try
                {
                    // create the ProcessStartInfo using "cmd" as the program to be run,
                    // and "/C " as the parameters.
                    // Incidentally, /C tells cmd that we want it to execute the command that follows,
                    // and then exit.
                    System.Diagnostics.ProcessStartInfo procStartInfo = new System.Diagnostics.ProcessStartInfo("cmd", "/C " + command);
                    // The following commands are needed to redirect the standard output.
                    // This means that it will be redirected to the Process.StandardOutput StreamReader.
                    procStartInfo.RedirectStandardError = true;
                    procStartInfo.RedirectStandardOutput = true;
                    procStartInfo.UseShellExecute = false;
                    // Do not create the black window.
                    procStartInfo.CreateNoWindow = true;
                    // Now we create a process, assign its ProcessStartInfo and start it
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.StartInfo = procStartInfo;
                    proc.Start();
                    // Get the output into a string
                    string result = proc.StandardOutput.ReadToEnd();
                    // Display the command output.
                    Console.WriteLine(result);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString()); 
                }
            }

            return true;
        }
    }
}