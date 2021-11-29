using System;
using System.IO;
using NDesk.Options;
using FormatFile.Common;
using System.Linq;

namespace GenerateFormatFile
{
    class Program
    {
        static void Main(string[] args)
        {

            string file = "";
            string delimiter = "";
            bool normalize = false;
            bool showHelp = false;
            string servername = ""; //GenerateFormatFile.Properties.Settings.Default.servername;
            string username = ""; //GenerateFormatFile.Properties.Settings.Default.username;
            string password = ""; //GenerateFormatFile.Properties.Settings.Default.password;
            string database = ""; //GenerateFormatFile.Properties.Settings.Default.databasename;
            string tablename = "";
            bool LoadToSQL = false; //GenerateFormatFile.Properties.Settings.Default.loadToSQL;

            bool removeoriginalfile = false;
            int headerrow = 1;
            bool genericcolumnname = false;
            string suffix = "";
            bool autoheader = false;
            string procedurename = "";

            var p = new OptionSet()
            {
                {"f|path=", "(needed) The filename/folderpath of the file or folder to be processed", (string v)=>file=v },
                {"d|delimiter=", "(needed) Defines the delimiter for the fiels", (string v)=>delimiter=v },
                {"n|normalize:", "(optional) Normalize file if rows have different number of delimiters", (bool v)=>normalize = v != null },
                {"h|headerrow=", "(optional) The header row, default is 1", v=>headerrow = Int32.Parse(v) },
                {"a|autoheader:", "(optional) When set to TRUE, normalize file will ignore any header row specified and evaluate all rows to determine the max number of columns. Default is FALSE.", (bool v)=>autoheader = v != null},
                {"g|genericcolumnname:", "(optional) Use generic column name, like C1,C2 etc. Useful when the headerrow value is not the desired column names. Default is FALSE", (bool v)=>genericcolumnname = v != null},
                {"r|removeoriginalfile:", "(optional) When processing Excel file, after convert the Excel file to CSV, also delete the original Excel file. Default is FALSE.", (bool v)=>removeoriginalfile = v != null },
                {"sf|suffix=", "(optional) When normalize file, keep original file as is and create a new file with the suffix appended to the file name. Default is overwrite existing one.", (string v)=>suffix=v},
                {"l|loadtoSQL:", "(optional) The filename/folderpath of the file or folder to be processed. Default is FALSE", (bool v)=>LoadToSQL = v != null },
                {"s|servername=", "(optional) SQL Server instance name, when used with option -l.", (string v)=>servername=v },
                {"u|username=", "(optional) Username to connect to SQL instance, when used with option -l.", (string v)=>username=v },
                {"p|password=", "(optional) Password. If not specified, will use Integrated Security for the connection.", (string v)=>password=v },
                {"db|database=", "(optional) Database to load the data, when used with option -l.", (string v)=>database=v },
                {"t|tablename=", "(optional) Table name to load the data into using BCP, when used with option -l.", (string v)=>tablename=v },
                {"sp|procedurename=", "(optional) Stored procedure name used for loading the data, when used with option -l. The stored procedure will have two parameter, first one for CSV data file name and second one for the format file name. The store procedure can then using your preferred method to load/manipulate the data.", (string v)=>procedurename=v },
                {"?|help", "Show this message and end", v=>showHelp = v != null },
            };

            try
            {
                p.Parse(args);

                if(showHelp)
                {
                    ShowHelp(p);
                    return;
                }
                //Check file and folder for existance
                if (string.IsNullOrWhiteSpace(file))
                {
                    throw new OptionException("Path or filename cannot be blank or empty", file);
                }

                //Check file and folder for existance
                if (string.IsNullOrWhiteSpace(delimiter))
                {
                    throw new OptionException("Delimiter cannot be blank or empty", delimiter);
                }

                //Check SQL config
                if (LoadToSQL)
                {
                    if (string.IsNullOrEmpty(servername))
                        throw new OptionException("",servername);

                    if (string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                        throw new OptionException("", username);

                    if (string.IsNullOrEmpty(database))
                        throw new OptionException("", database);

                    if (string.IsNullOrEmpty(tablename) && string.IsNullOrEmpty(procedurename))
                        throw new Exception("Need to specify either table name or stored process name.");
                }

                GenerateFormatFile.Properties.Settings.Default.servername = servername;
                GenerateFormatFile.Properties.Settings.Default.databasename = database;
                GenerateFormatFile.Properties.Settings.Default.username = username;
                GenerateFormatFile.Properties.Settings.Default.password = password;
                GenerateFormatFile.Properties.Settings.Default.loadToSQL = LoadToSQL;
                Console.WriteLine(LoadToSQL);
                if (File.Exists(file))
                {
                    GenerateXML.HandleFile(file, delimiter, normalize, removeoriginalfile, headerrow, genericcolumnname, suffix, autoheader, tablename, procedurename);
                }
                else if (Directory.Exists(file))
                {
                    string[] fileEntries = Directory.GetFiles(file);

                    foreach (string filename in fileEntries.Where(i => i.EndsWith(".csv") | i.EndsWith(".xlsx") | i.EndsWith(".xls")))
                    {
                        GenerateXML.HandleFile(filename, delimiter, normalize, removeoriginalfile, headerrow, genericcolumnname, suffix, autoheader, tablename, procedurename);
                    }
                }
                else
                {
                    throw new OptionException("{0}' is not a valid file or path.", file);
                }

            }
            catch (OptionException e)
            {
                Console.Write("GenerateFormatFile: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `GenerateFormatFile --help' for more information");
                Console.WriteLine("Path to the file: {0}", Path.GetDirectoryName(file));
                return;
            }

        }
        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Use: GenerateFormatFile [OPTIONS]");
            Console.WriteLine("Generates a predefined formatfile and can load data directly to SQL Server.");
            Console.WriteLine("© Brian Bønk - 2018");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out, 2);
            Console.WriteLine();
            Console.WriteLine("The generated format file name is original file name with prefix formatfile_");
            Console.WriteLine("When load Excel file, Excel file will be converted to CSV for loading and can use -r option to remove original Excel file after convert.");
            Console.WriteLine("If CSV using TAB delimiter, TAB delimiter will be converted to pipe '|', and a new file with _tabConverted appended to the original name will be created if -r option is not specified.");
       }
    }
}
