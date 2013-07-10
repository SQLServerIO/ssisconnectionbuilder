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
using Microsoft.SqlServer.Dts.Runtime;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;
using Microsoft.SqlServer.Dts.Pipeline.Wrapper;
using NDesk.Options;
using Excel;

namespace SSISConnectionBuilder
{
    class Program
    {
        public static string SchemaFile;
        public static string Delimiter;
        public static string PackageName;
        public static string CSVFileName;
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

            //build table to hold the SSIS to SQL Server data type mappings
            DataTable dt = new DataTable();
            dt.TableName = "SQLServerToSSISDataTypes";

            DataColumn[] keys = new DataColumn[2];
            DataColumn column;

            //set primary key
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName= "SQLServerDataType";
            dt.Columns.Add(column);
            keys[0] = column;
            dt.PrimaryKey = keys;
            dt.Columns.Add("SSISDataType", Type.GetType("System.String"));
            dt.Columns.Add("SSISExpression", Type.GetType("System.String"));

            //read in data type conversions
            using (Stream stream = typeof(Program).Assembly.GetManifestResourceStream("SQLServerToSSISDataTypes.xml"))
            {
                dt.ReadXml(stream);
            }

            string pkg = PackageName;

            Microsoft.SqlServer.Dts.Runtime.Application app = new Microsoft.SqlServer.Dts.Runtime.Application();
            Microsoft.SqlServer.Dts.Runtime.Package p = new Microsoft.SqlServer.Dts.Runtime.Package();

            ConnectionManager ConMgr;
            ConMgr = p.Connections.Add("FlatFile");

            IDTSConnectionManagerFlatFile100 connectionFlatFile = (IDTSConnectionManagerFlatFile100)ConMgr.InnerObject;
            ConMgr.ConnectionString = CSVFileName;
            ConMgr.Name = "CSV";
            ConMgr.Description = "CSV File connection";
            connectionFlatFile.LocaleID = 1033;
            connectionFlatFile.CodePage = 1252;
            connectionFlatFile.Unicode = isUnicode;
            connectionFlatFile.RowDelimiter = "\r\n";
            connectionFlatFile.TextQualifier = "\"";
            connectionFlatFile.HeaderRowDelimiter = "\r\n";
            connectionFlatFile.ColumnNamesInFirstDataRow = true;
            connectionFlatFile.Format = "Delimited";

            /*
            how to set values directly to the ConMgr
            ConMgr.Properties["AlwaysCheckForRowDelimiters"].SetValue(ConMgr, false);
            ConMgr.Properties["CodePage"].SetValue(ConMgr, 1252);
            ConMgr.Properties["ColumnNamesInFirstDataRow"].SetValue(ConMgr, true);
            ConMgr.Properties["Columns"].SetValue(ConMgr, "");
            ConMgr.Properties["ConnectionString"].SetValue(ConMgr, "");
            ConMgr.Properties["CreationName"].SetValue(ConMgr, "");
            ConMgr.Properties["DataRowsToSkip"].SetValue(ConMgr, 0);
            ConMgr.Properties["Description"].SetValue(ConMgr, "");
            ConMgr.Properties["FileUsageType"].SetValue(ConMgr, "");
            ConMgr.Properties["Format"].SetValue(ConMgr, "Delimited");
            ConMgr.Properties["HeaderRowDelimiter"].SetValue(ConMgr, "\r\n");
            ConMgr.Properties["HeaderRowsToSkip"].SetValue(ConMgr, 0);
            ConMgr.Properties["ID"].SetValue(ConMgr, "");
            ConMgr.Properties["LocaleID"].SetValue(ConMgr, "");
            ConMgr.Properties["Name"].SetValue(ConMgr, "");
            ConMgr.Properties["ProtectionLevel"].SetValue(ConMgr, "");
            ConMgr.Properties["RowDelimiter"].SetValue(ConMgr, "\r\n");
            ConMgr.Properties["Scope"].SetValue(ConMgr, "");
            ConMgr.Properties["SupportsDTCTransactions"].SetValue(ConMgr, false);
            ConMgr.Properties["TextQualifier"].SetValue(ConMgr, "\"");
            ConMgr.Properties["Unicode"].SetValue(ConMgr, false);
            */

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
            
            IDTSConnectionManagerFlatFileColumn100 flatFileColumn = null;
            int rc = 1;
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

                flatFileColumn = (IDTSConnectionManagerFlatFileColumn100)connectionFlatFile.Columns.Add();
                flatFileColumn.ColumnType = "Delimited";
                flatFileColumn.DataType = findSSISType(dataType, dt);
                IDTSName100 columnName = (IDTSName100)flatFileColumn;
                columnName.Name = row[0].ToString();

                //Column	Type	Precision	Scale
                //Console.WriteLine(row[0].ToString());
                //Console.WriteLine(dataType);
                //Console.WriteLine(row[2].ToString());

                if (dataType == "smallint")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "int")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "bigint")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "tinyint")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "real")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "float")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "char")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                if (dataType == "nchar")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                if (dataType == "varchar")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                if (dataType == "nvarchar")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                if (dataType == "sql_variant")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                if (dataType == "xml")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                if (dataType == "date")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "bit")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "decimal")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = Convert.ToInt32(row[2].ToString()); ;
                    flatFileColumn.DataScale = Convert.ToInt32(row[3].ToString()); ;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "numeric")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = Convert.ToInt32(row[2].ToString()); ;
                    flatFileColumn.DataScale = Convert.ToInt32(row[3].ToString()); ;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "smallmoney")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "money")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "uniqueidentifier")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "binary")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "varbinary")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "timestamp")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "date")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "time(p)")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "datetime")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "smalldatetime")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "datetime2")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "datetimeoffset(p)")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = Convert.ToInt32(row[2].ToString());
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "image")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.MaximumWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = false;
                }

                if (dataType == "text")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                if (dataType == "ntext")
                {
                    flatFileColumn.ColumnWidth = 0;
                    flatFileColumn.DataPrecision = 0;
                    flatFileColumn.DataScale = 0;
                    flatFileColumn.TextQualified = true;
                }

                //if this is the last column set the delimiter to control linefeed instead of your row delimiter
                if (result.Tables[0].Rows.Count == rc)
                {
                    flatFileColumn.ColumnDelimiter = "\r\n";
                }
                else
                {
                    flatFileColumn.ColumnDelimiter = Delimiter;
                }
                rc++;
            }

            excelReader.Close();

            //write to file system
            app.SaveToXml(pkg, p, null);
            Console.WriteLine("Done, Any Key To Continue:");
            Console.ReadKey();
        }

        private static DataType findSSISType(string sqltype,DataTable dt)
        {
            DataRow foundRow = dt.Rows.Find(sqltype);
            DataType retdt = new DataType();
            if (foundRow != null)
            {
                retdt = (DataType) Enum.Parse(typeof(DataType),foundRow[2].ToString());
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

             { "u|unicode", "csv file is unicode.", v => isUnicode = v  != null },

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
                CSVFileName == null
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

        public string GetResourceTextFile(string filename)
        {
            string result = string.Empty;

            using (Stream stream = this.GetType().Assembly.
                       GetManifestResourceStream("assembly.folder." + filename))
            {
                using (StreamReader sr = new StreamReader(stream))
                {
                    result = sr.ReadToEnd();
                }
            }
            return result;
        }
    }
}