/***************************************************************************\
Module Name:  Program.cs
Project:      SSISConnectionBuilder
Copyright (c) Wesley D. Brown

SSISConnectionBuilder is a simple command line tool to ease the burden of
building flat file connectors for SSIS.

Exept where noted or, code is owned by another author, this source is subject 
to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

Excel Data Reader
http://exceldatareader.codeplex.com/
 
NDesk.Options
http://www.ndesk.org/Options

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.Common;
using System.Text;
using System.Collections.Generic;
using NDesk.Options;
using Excel;

namespace SSISConnectionBuilder
{
    class ConnectionBuilderMain
    {
        public static string SchemaFile;
        public static string Delimiter;
        public static string PackageName;
        public static string CSVFileName;
        public static string ConnectionName;
        public static bool isUnicode = false;

        static void Main(string[] args)
        {
            /*
             Options:
                  -s, --schemafile[=VALUE]   Your excel schema defenition file.
                  -d, --delimiter[=VALUE]    The column seperator you wish to use, usually a
                                               comma or pipe.
                  -p, --packagename[=VALUE]  Name of the dtsx file that will have your
                                               connection in.
                  -c, --csvfilename[=VALUE]  Name of the csv file that your connection will
                                               use.
                  -u, --unicode              csv file is unicode.
                  -?, -h, --help             show this message and exit
             */
            var parseerr = ParseCommandLine(args);
            if (parseerr == 1)
                return;

            int unicodeFlag = 0;
            if (isUnicode)
                unicodeFlag = 1;
            else
                unicodeFlag = 0;

            Dictionary<string, string> options = new Dictionary<string, string>(){
                {"RCNAME", "John Smith"},
                {"RCOMPUTER", "Virtual"},
                {"RDATE", DateTime.Now.ToString()},
                {"RPNAME", PackageName},
                {"RConnName", ConnectionName},
                {"RConnDescription", "CSV File connection"},
                {"RCDelimited", "Delimited"},
                {"RConnSkip", "0"},
                {"RConnUnicode", unicodeFlag.ToString()},
                {"RConnHeaderRowDelimiter", DelimiterToHex("\r\n")},
                {"RConnColumnNamesInFirstDataRow", "-1"},
                {"RConnRowDelimiter", DelimiterToHex("\r\n")},
                {"RConnDataRowsToSkip", "0"},
                {"RConnTextQualifier", "\""},
                {"RCConnCodePage", "1252"},
                {"RConnConnectionString", CSVFileName}
            };

            Dictionary<string, string> ColOptions = new Dictionary<string, string>(){};

            PackageBuilder pg = new PackageBuilder();

            StringBuilder PackageXML = new StringBuilder();

            ////build table to hold the SSIS DataTypes Enum mappings
            DataTable SSISDataTypesEnum = new DataTable();
            SSISDataTypesEnum.TableName = "SSISDataTypesEnum";

            DataColumn[] SSISDataTypesEnumkeys = new DataColumn[2];
            DataColumn SSISDataTypesEnumcolumn;

            //set primary key
            SSISDataTypesEnumcolumn = new DataColumn();
            SSISDataTypesEnumcolumn.DataType = System.Type.GetType("System.String");
            SSISDataTypesEnumcolumn.ColumnName = "SSISDataType";
            SSISDataTypesEnum.Columns.Add(SSISDataTypesEnumcolumn);
            SSISDataTypesEnumkeys[0] = SSISDataTypesEnumcolumn;

            SSISDataTypesEnum.PrimaryKey = SSISDataTypesEnumkeys;

            SSISDataTypesEnum.Columns.Add("NumericValue", Type.GetType("System.String"));

            //read in data type conversions
            using (Stream stream = typeof(ConnectionBuilderMain).Assembly.GetManifestResourceStream("SSISConnectionBuilder.SSISDataTypesEnum.xml"))
            {
                SSISDataTypesEnum.ReadXml(stream);
            }


            //build table to hold the SSIS to SQL Server data type mappings
            DataTable SQLServerToSSISDataTypes = new DataTable();
            SQLServerToSSISDataTypes.TableName = "SQLServerToSSISDataTypes";

            DataColumn[] keys = new DataColumn[2];
            DataColumn column;

            //set primary key
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "SQLServerDataType";
            SQLServerToSSISDataTypes.Columns.Add(column);
            keys[0] = column;
            SQLServerToSSISDataTypes.PrimaryKey = keys;
            SQLServerToSSISDataTypes.Columns.Add("SSISDataType", Type.GetType("System.String"));
            SQLServerToSSISDataTypes.Columns.Add("SSISExpression", Type.GetType("System.String"));

            //read in data type conversions
            using (Stream stream = typeof(ConnectionBuilderMain).Assembly.GetManifestResourceStream("SSISConnectionBuilder.SQLServerToSSISDataTypes.xml"))
            {
                SQLServerToSSISDataTypes.ReadXml(stream);
            }

            StringBuilder pkg = new StringBuilder();
            pkg.Append(pg.PackageHeader(options));
            pkg.Append(pg.ConnectionManagerHeader(options));

            var filePath = SchemaFile;
            FileStream excelStream;
            try
            {
                excelStream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            }
            catch
            {
                Console.WriteLine("cannot open file, exiting.");
                return;
            }
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(excelStream);
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            int rc = 2;
            foreach (DataRow row in result.Tables[0].Rows)
            {
                if (row[0].ToString().Length == 0)
                    break;

                var dataType = row[1].ToString();

                //check for varchar or nvarchar max
                if (dataType == "varchar" && Convert.ToInt32(row[2].ToString()) == -1)
                {
                    dataType = "text";
                }
                if (dataType == "nvarchar" && Convert.ToInt32(row[2].ToString()) == -1)
                {
                    dataType = "ntext";
                }
                ColOptions = new Dictionary<string, string>(){
                    {"RColColumnType", "Delimited"},
                    {"RColColumnWidth", "0"},
                    {"RColMaximumWidth", row[2].ToString()},
                    {"RColDataType", findSSISTypeEnum(findSSISType(dataType,SQLServerToSSISDataTypes),SSISDataTypesEnum)},
                    {"RColDataPrecision", row[2].ToString()},
                    {"RColDataScale", row[3].ToString()},
                    {"RColObjectName", row[0].ToString()},
                    {"RColDescription", ""},
                    {"RColSQLDataType", dataType}
                };

                //if this is the last column set the delimiter to control linefeed instead of your row delimiter
                if (result.Tables[0].Rows.Count == rc)
                {
                    Console.WriteLine(DelimiterToHex("\r\n"));
                    Console.WriteLine("Last Column");
                    ColOptions.Add("RColColumnDelimiter", DelimiterToHex("\r\n"));
                }
                else
                {
                    Console.WriteLine("Column");
                    ColOptions.Add("RColColumnDelimiter", DelimiterToHex(","));
                }
                rc++;
                pkg.Append(pg.FlatFileColumn(ColOptions));
            }

            excelReader.Close();

            pkg.Append(pg.ConnectionManagerEnd(options));
            pkg.Append(pg.PackageEnd(options));

            using (StreamWriter outfile = new StreamWriter(PackageName))
            {
                outfile.Write(pkg);
            }

            //write to file system
            //app.SaveToXml(pkg, p, null);
            Console.WriteLine("Done, Any Key To Continue:");
            Console.ReadKey();
        }

        public static String DelimiterToHex(String data)
        {
            String output = String.Empty;
            foreach (Char c in data)
            {
                output += "_x00" + ((int)c).ToString("X2") + "_";
            }
            return output;
        }

        private static string findSSISType(string sqltype, DataTable dt)
        {
            DataRow foundRow = dt.Rows.Find(sqltype);
            string retdt = "";
            if (foundRow != null)
            {
                retdt = foundRow[2].ToString();
            }
            return retdt;
        }

        private static string findSSISTypeEnum(string sqltype, DataTable dt)
        {
            DataRow foundRow = dt.Rows.Find(sqltype);
            string retdt = "";
            if (foundRow != null)
            {
                retdt = foundRow[1].ToString();
            }
            return retdt;
        }

        static public int ParseCommandLine(string[] args)
        {
            bool showHelp = false;
            var p = new OptionSet
                        {
              { "s:|schemafile:", "Your excel schema defenition file.",
              v => SchemaFile = v },

              { "d:|delimiter:", "The column seperator you wish to use, usually a comma or pipe.",
              v => Delimiter = v},

             { "p:|packagename:", "Name of the dtsx file that will have your connection in.",
              v => PackageName = v},

             { "c:|csvfilename:", "Name of the csv file that your connection will use.",
              v => CSVFileName = v},

              { "n:|connectionname:", "The name connection will use inside the package.",
              v => ConnectionName = v},

             { "u|unicode", "csv file is unicode.", 
                 v => isUnicode = v  != null },

             { "?|h|help",  "show this message and exit", 
              v => showHelp = v != null },
            };

            try
            {
                p.Parse(args);
            }

            catch (OptionException e)
            {
                Console.Write("SSISConnectionBuilder Error: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `SSISConnectionBuilder --help' for more information.");
                return 1;
            }

            if (args.Length == 0)
            {
                ShowHelp("Error: please specify some commands....", p);
                return 1;
            }


            if (SchemaFile == null ||
                Delimiter == null ||
                PackageName == null ||
                CSVFileName == null ||
                ConnectionName == null
                && !showHelp
            )
            {
                ShowHelp("You must specifiy all commands execpt u|unicode", p);
                return 1;
            }


            if (showHelp)
            {
                ShowHelp(p);
                return 1;
            }
            return 0;
        }

        static void ShowHelp(string message, OptionSet p)
        {
            Console.WriteLine(message);
            Console.WriteLine("Usage: SSISConnectionBuilder [OPTIONS]");
            Console.WriteLine("Build an SSIS connector dynamically.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage: SSISConnectionBuilder [OPTIONS]");
            Console.WriteLine("Build an SSIS connector dynamically.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }
    }
}
