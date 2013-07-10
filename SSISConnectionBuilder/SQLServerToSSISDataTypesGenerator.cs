/***************************************************************************\
Module Name:  SQLServerToSSISDataTypesGenerator.cs
Project:      SSISConnectionBuilder
Copyright (c) Wesley D. Brown

SSISConnectionBuilder is a simple command line tool to ease the burden of
building flat file connectors for SSIS.

Exept where noted or, code is owned by another author, this source is subject 
to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/
using System;
using System.IO;
using System.Data;
using System.Data.Common;
namespace SSISConnectionBuilder
{
    class SQLServerToSSISDataTypes
    {
        static void GenerateXML()
        {
            DataTable dt = new DataTable();
            DataColumn[] keys = new DataColumn[2];
            DataColumn column;

            //set primary key
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "SQLServerDataType";
            dt.Columns.Add(column);
            keys[0] = column;
            dt.PrimaryKey = keys;
            dt.Columns.Add("SSISDataType", Type.GetType("System.String"));
            dt.Columns.Add("SSISExpression", Type.GetType("System.String"));

            DataRow dr;
            /*
            unknown or not mapped to a SQL Server data type
            dr["SSISDataType"] = "database time";
            dr["SSISExpression"] ="DT_DBTIME";
            dr["SQLServerDataType"] ="";
            dt.Rows.Add(dr);
            file timestamp,(DT_FILETIME),
            dr["SSISDataType"] = "numeric";
            dr["SSISExpression"] ="DT_NUMERIC";
            dr["SQLServerDataType"] ="decimal";
            dt.Rows.Add(dr);
            two-byte unsigned integer,(DT_UI2),
            four-byte unsigned integer,(DT_UI4),
            eight-byte unsigned integer,(DT_UI8),
            dr["SSISDataType"] = "single-byte signed integer";
            dr["SSISExpression"] ="DT_I1";
            dr["SQLServerDataType"] ="";
            dt.Rows.Add(dr);
            */
            dr = dt.NewRow();
            dr["SSISDataType"] = "two-byte signed integer";
            dr["SSISExpression"] = "DT_I2";
            dr["SQLServerDataType"] = "smallint";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "four-byte signed integer";
            dr["SSISExpression"] = "DT_I4";
            dr["SQLServerDataType"] = "int";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "eight-byte signed integer";
            dr["SSISExpression"] = "DT_I8";
            dr["SQLServerDataType"] = "bigint";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "single-byte unsigned integer";
            dr["SSISExpression"] = "DT_UI1";
            dr["SQLServerDataType"] = "tinyint";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "float";
            dr["SSISExpression"] = "DT_R4";
            dr["SQLServerDataType"] = "real";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "double-precision";
            dr["SSISExpression"] = "DT_R8";
            dr["SQLServerDataType"] = "float";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "string";
            dr["SSISExpression"] = "DT_STR";
            dr["SQLServerDataType"] = "varchar";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "string";
            dr["SSISExpression"] = "DT_STR";
            dr["SQLServerDataType"] = "char";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "Unicode text stream";
            dr["SSISExpression"] = "DT_WSTR";
            dr["SQLServerDataType"] = "nchar";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "Unicode text stream";
            dr["SSISExpression"] = "DT_WSTR";
            dr["SQLServerDataType"] = "nvarchar";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "Unicode text stream";
            dr["SSISExpression"] = "DT_WSTR";
            dr["SQLServerDataType"] = "sql_variant";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "Unicode text stream";
            dr["SSISExpression"] = "DT_WSTR";
            dr["SQLServerDataType"] = "xml";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "Boolean";
            dr["SSISExpression"] = "DT_BOOL";
            dr["SQLServerDataType"] = "bit";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "numeric";
            dr["SSISExpression"] = "DT_NUMERIC";
            dr["SQLServerDataType"] = "numeric";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "decimal";
            dr["SSISExpression"] = "DT_DECIMAL";
            dr["SQLServerDataType"] = "decimal";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "currency";
            dr["SSISExpression"] = "DT_CY";
            dr["SQLServerDataType"] = "smallmoney";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "currency";
            dr["SSISExpression"] = "DT_CY";
            dr["SQLServerDataType"] = "money";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "unique identifier";
            dr["SSISExpression"] = "DT_GUID";
            dr["SQLServerDataType"] = "uniqueidentifier";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "byte stream";
            dr["SSISExpression"] = "DT_BYTES";
            dr["SQLServerDataType"] = "binary";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "byte stream";
            dr["SSISExpression"] = "DT_BYTES";
            dr["SQLServerDataType"] = "varbinary";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "byte stream";
            dr["SSISExpression"] = "DT_BYTES";
            dr["SQLServerDataType"] = "timestamp";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "database date";
            dr["SSISExpression"] = "DT_DBDATE";
            dr["SQLServerDataType"] = "date";
            dt.Rows.Add(dr);
            //duplicate I don't know what to do with this yet.
            //dr = dt.NewRow();
            //dr["SSISDataType"] = "date";
            //dr["SSISExpression"] ="DT_DATE";
            //dr["SQLServerDataType"] ="date";
            //dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "database time with precision";
            dr["SSISExpression"] = "DT_DBTIME2";
            dr["SQLServerDataType"] = "time(p)";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "database timestamp";
            dr["SSISExpression"] = "DT_DBTIMESTAMP";
            dr["SQLServerDataType"] = "datetime";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "database timestamp";
            dr["SSISExpression"] = "DT_DBTIMESTAMP";
            dr["SQLServerDataType"] = "smalldatetime";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "database timestamp with precision";
            dr["SSISExpression"] = "DT_DBTIMESTAMP2";
            dr["SQLServerDataType"] = "datetime2";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "database timestamp with timezone";
            dr["SSISExpression"] = "DT_DBTIMESTAMPOFFSET";
            dr["SQLServerDataType"] = "datetimeoffset(p)";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr = dt.NewRow();
            dr["SSISDataType"] = "image";
            dr["SSISExpression"] = "DT_IMAGE";
            dr["SQLServerDataType"] = "image";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            //expects a code page
            dr["SSISDataType"] = "text stream";
            dr["SSISExpression"] = "DT_TEXT";
            dr["SQLServerDataType"] = "text";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["SSISDataType"] = "Unicode string";
            dr["SSISExpression"] = "DT_NTEXT";
            dr["SQLServerDataType"] = "ntext";
            dt.Rows.Add(dr);
            dr = dt.NewRow();

            dt.TableName = "SQLServerToSSISDataTypes";
            dt.WriteXml(@"SQLServerToSSISDataTypes.xml");
            dt.WriteXmlSchema(@"SQLServerToSSISDataTypes.xlst");
        }
    }
}