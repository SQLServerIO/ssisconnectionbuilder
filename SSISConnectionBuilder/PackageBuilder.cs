using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSISConnectionBuilder
{
    class PackageBuilder
    {
        /*
         * tokens that can be replaced.
         * RCNAME
         * RCOMPUTER
         * RDATE
         * RGUID(X)
         * RPNAME
         * RConnName
         * RConnDescription
         * RCDelimited
         * RConnSkip
         * RConnHeaderRowDelimiter
         * RConnRowDelimiter
         * RConnDataRowsToSkip
         * RConnTextQualifier
         * RCConnCodePage
         * RColColumnType
         * RColColumnDelimiter
         * RColColumnWidth
         * RColMaximumWidth
         * RColDataType
         * RColDataPrecision
         * RColDataScale
         * RColTextQualified
         * RColObjectName
         * RColDescription
         * RConnConnectionString
         */

        public string PackageHeader(Dictionary<string, string> options)
        {
            StringBuilder PackageHeader = new StringBuilder();
            PackageHeader.Append(@"<?xml version=""1.0""?><DTS:Executable xmlns:DTS=""www.microsoft.com/SqlServer/Dts"" DTS:ExecutableType=""SSIS.Package.2"">");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""PackageFormatVersion"">3</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""VersionComments""></DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""CreatorName"">RCNAME</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""CreatorComputerName"">RCOMPUTER</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""CreationDate"" DTS:DataType=""7"">RDATE</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""PackageType"">0</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""ProtectionLevel"">1</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""MaxConcurrentExecutables"">-1</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""PackagePriorityClass"">0</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""VersionMajor"">1</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""VersionMinor"">0</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""VersionBuild"">0</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""VersionGUID"">{RGUID}</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""EnableConfig"">0</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""CheckpointFileName""></DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""SaveCheckpoints"">0</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""CheckpointUsage"">0</DTS:Property>");
            PackageHeader.Append("\r\n");
            PackageHeader.Append(@"<DTS:Property DTS:Name=""SuppressConfigurationWarnings"">0</DTS:Property>""");
            PackageHeader.Append("\r\n");
            foreach (string key in options.Keys)
            {
                PackageHeader = PackageHeader.Replace(key, options[key]);
            }

            PackageHeader = PackageHeader.Replace("RGUID", Guid.NewGuid().ToString());

            return PackageHeader.ToString();
        }

        public string ConnectionManagerHeader(Dictionary<string, string> options)
        {
            StringBuilder ConnectionManagerHeader = new StringBuilder();
            ConnectionManagerHeader.Append(@"<DTS:ConnectionManager>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""DelayValidation"">0</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""ObjectName"">RConnName</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""DTSID"">{RGUID}</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""Description"">RConnDescription</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""CreationName"">FLATFILE</DTS:Property><DTS:ObjectData><DTS:ConnectionManager>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""FileUsageType"">0</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""Format"">RCDelimited</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""LocaleID"">1033</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""Unicode"">RConnUnicode</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""HeaderRowsToSkip"">RConnSkip</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""HeaderRowDelimiter"" xml:space=""preserve"">RConnHeaderRowDelimiter</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""ColumnNamesInFirstDataRow"">RConnColumnNamesInFirstDataRow</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""RowDelimiter"" xml:space=""preserve"">RConnRowDelimiter</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""DataRowsToSkip"">RConnDataRowsToSkip</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""TextQualifier"">RConnTextQualifier</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");
            ConnectionManagerHeader.Append(@"<DTS:Property DTS:Name=""CodePage"">RCConnCodePage</DTS:Property>");
            ConnectionManagerHeader.Append("\r\n");

            foreach (string key in options.Keys)
            {
                ConnectionManagerHeader = ConnectionManagerHeader.Replace(key, options[key]);
            }

            ConnectionManagerHeader = ConnectionManagerHeader.Replace("RGUID", Guid.NewGuid().ToString());

            return ConnectionManagerHeader.ToString();

        }

        public string FlatFileColumn(Dictionary<string, string> options)
        {
            //Console.WriteLine("Column Name:{0} Column Type:{1} Column Width:{2}", row[0].ToString(), dataType, row[2].ToString());

            var dataType = options["RColSQLDataType"];

            StringBuilder FlatFileColumn = new StringBuilder();

            //#region mapdatatype
            if (dataType == "smallint" || dataType == "int" || dataType == "bigint" || dataType == "tinyint" || dataType == "real" || 
                dataType == "float" || dataType == "date" || dataType == "bit" || dataType == "smallmoney" || dataType == "binary" || dataType == "varbinary")
            {
                FlatFileColumn.Append(@"<DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnType"">RColColumnType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnDelimiter"" xml:space=""preserve"">RColColumnDelimiter</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""MaximumWidth"">RColMaximumWidth</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataType"">RColDataType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataPrecision"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataScale"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""TextQualified"">-1</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ObjectName"">RColObjectName</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DTSID"">{RGUID}</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""Description""></DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""CreationName""></DTS:Property></DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
            }

            if (dataType == "char" || dataType == "nchar" || dataType == "varchar" || dataType == "nvarchar" || dataType == "sql_variant" || dataType == "xml")
            {
                FlatFileColumn.Append(@"<DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnType"">RColColumnType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnDelimiter"" xml:space=""preserve"">RColColumnDelimiter</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""MaximumWidth"">RColMaximumWidth</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataType"">RColDataType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataPrecision"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataScale"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""TextQualified"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ObjectName"">RColObjectName</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DTSID"">{RGUID}</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""Description""></DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""CreationName""></DTS:Property></DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
            }


            if (dataType == "decimal" || dataType == "numeric")
            {
                FlatFileColumn.Append(@"<DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnType"">RColColumnType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnDelimiter"" xml:space=""preserve"">RColColumnDelimiter</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""MaximumWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataType"">RColDataType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataPrecision"">RColDataPrecision</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataScale"">RColDataScale</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""TextQualified"">-1</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ObjectName"">RColObjectName</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DTSID"">{RGUID}</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""Description""></DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""CreationName""></DTS:Property></DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
            }

            if (dataType == "money" || dataType == "uniqueidentifier" || dataType == "timestamp" || dataType == "date" || dataType == "datetime" || 
                dataType == "smalldatetime" || dataType == "datetime2" || dataType == "image" || dataType == "text" || dataType == "ntext")
            {
                FlatFileColumn.Append(@"<DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnType"">RColColumnType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnDelimiter"" xml:space=""preserve"">RColColumnDelimiter</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""MaximumWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataType"">RColDataType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataPrecision"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataScale"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""TextQualified"">-1</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ObjectName"">RColObjectName</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DTSID"">{RGUID}</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""Description""></DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""CreationName""></DTS:Property></DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
            }

            if (dataType == "time(p)" || dataType == "datetimeoffset(p)")
            {
                FlatFileColumn.Append(@"<DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnType"">RColColumnType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnDelimiter"" xml:space=""preserve"">RColColumnDelimiter</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ColumnWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""MaximumWidth"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataType"">RColDataType</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataPrecision"">RColDataPrecision</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DataScale"">0</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""TextQualified"">-1</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""ObjectName"">RColObjectName</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""DTSID"">{RGUID}</DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""Description""></DTS:Property>");
                FlatFileColumn.Append("\r\n");
                FlatFileColumn.Append(@"<DTS:Property DTS:Name=""CreationName""></DTS:Property></DTS:FlatFileColumn>");
                FlatFileColumn.Append("\r\n");
            }
            //#endregion

            foreach (string key in options.Keys)
            {
                FlatFileColumn = FlatFileColumn.Replace(key, options[key]);
            }

            FlatFileColumn = FlatFileColumn.Replace("RGUID", Guid.NewGuid().ToString());
            return FlatFileColumn.ToString();
        }

        public string ConnectionManagerEnd(Dictionary<string, string> options)
        {
            StringBuilder ConnectionManagerEnd = new StringBuilder(@"<DTS:Property DTS:Name=""ConnectionString"">RConnConnectionString</DTS:Property></DTS:ConnectionManager></DTS:ObjectData></DTS:ConnectionManager>");

            foreach (string key in options.Keys)
            {
                ConnectionManagerEnd = ConnectionManagerEnd.Replace(key, options[key]);
            }

            ConnectionManagerEnd = ConnectionManagerEnd.Replace("RGUID", Guid.NewGuid().ToString());

            return ConnectionManagerEnd.ToString();

        }

        public string PackageEnd(Dictionary<string, string> options)
        {
            StringBuilder PackageEnd = new StringBuilder();
            PackageEnd.Append(@"<DTS:Property DTS:Name=""LastModifiedProductVersion"">10.50.1600.1</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""ForceExecValue"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""ExecValue"" DTS:DataType=""3"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""ForceExecutionResult"">-1</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""Disabled"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""FailPackageOnFailure"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""FailParentOnFailure"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""MaxErrorCount"">1</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""ISOLevel"">1048576</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""LocaleID"">1033</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""TransactionOption"">1</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""DelayValidation"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:LoggingOptions>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""LoggingMode"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""FilterKind"">1</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""EventFilter"" DTS:DataType=""8""></DTS:Property></DTS:LoggingOptions>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""ObjectName"">{RGUID1}</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""DTSID"">{RGUID2}</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""Description""></DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""CreationName"">SSIS.Package.2</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"<DTS:Property DTS:Name=""DisableEventHandlers"">0</DTS:Property>");
            PackageEnd.Append("\r\n");
            PackageEnd.Append(@"</DTS:Executable>");
            foreach (string key in options.Keys)
            {
                PackageEnd = PackageEnd.Replace(key, options[key]);
            }

            PackageEnd = PackageEnd.Replace("RGUID1", Guid.NewGuid().ToString());
            PackageEnd = PackageEnd.Replace("RGUID2", Guid.NewGuid().ToString());

            return PackageEnd.ToString();
        }

        public static String DelimiterToHex(String data)
        {
            String output = String.Empty;
            foreach (Char c in data)
            {
                output += "_x00"+((int)c).ToString("X2")+"_";
            }
            return output;
        }
    }
}
