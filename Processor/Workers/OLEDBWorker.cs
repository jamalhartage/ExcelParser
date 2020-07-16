using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace Processor.Workers
{
    public class OLEDBWorker : WorkerBase
    {
        public OLEDBWorker(string filePhysical) :
                base(filePhysical)
        {}

        public OLEDBWorker(string filePhysical, Dictionary<string, int> headerRowInfo) :
                base(filePhysical, headerRowInfo)
        {}

        public override DataSet GetData()
        {
            return getOLEDBWorker();
        }

        private DataSet getOLEDBWorker()
        {
            DataSet ds = datasetFromExcel_OLE();

            if ((ds == null) && (FileType == ExcelExt.XLSX))
            {
                FileType = ExcelExt.XLS;
                ds = datasetFromExcel_OLE();

                if (ds == null)
                    FileType = ExcelExt.Unknown;
            }
            else if ((ds == null) && (FileType == ExcelExt.XLS))
            {
                FileType = ExcelExt.XLSX;
                ds = datasetFromExcel_OLE();

                if (ds == null)
                    FileType = ExcelExt.Unknown;
            }

            return ds;
        }
        
        private DataSet datasetFromExcel_OLE()
        {
            DataSet ds = null;

            try
            {
                string connectionString = getConnectionString();

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;

                    // Get all Sheets in Excel File
                    // Sheets are not in Ordinal(idx) order
                    DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    if (dtSheet != null && dtSheet.Rows.Count > 0)
                    {
                        ds = new DataSet();

                        // Loop through all Sheets to get data
                        for(int t = 0; t < dtSheet.Rows.Count; t++)
                        {
                            DataRow dr = dtSheet.Rows[t];

                            string sheetName = dr["TABLE_NAME"].ToString();
                            string modifiedSheetName = sheetName;

                            if (modifiedSheetName.StartsWith("'") && modifiedSheetName.EndsWith("$'"))
                                modifiedSheetName = modifiedSheetName.Substring(1, modifiedSheetName.Length - 2);

                            if (modifiedSheetName.EndsWith("$"))
                                modifiedSheetName = modifiedSheetName.Substring(0, modifiedSheetName.Length - 1);

                            // Get all rows from the Sheet
                            cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                            DataTable dt = new DataTable();
                            dt.TableName = modifiedSheetName;

                            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                            da.Fill(dt);

                            if (HasUserDefinedHeaderRow(modifiedSheetName))
                                if(GetHeaderRowIdx(modifiedSheetName) < dt.Rows.Count)
                                {
                                    for (int j = 0; j < dt.Columns.Count; j++)
                                    {
                                        object value = dt.Rows[GetHeaderRowIdx(modifiedSheetName)][j];

                                        if (value != null && value != DBNull.Value && !String.IsNullOrWhiteSpace(value.ToString()))
                                        {
                                            dt.Columns[j].ColumnName = value.ToString();
                                        }
                                        else
                                        {
                                            dt.Columns[j].ColumnName = "_BlankCol_" + j.ToString();
                                        }
                                    }

                                    for (int i = 0; i <= GetHeaderRowIdx(modifiedSheetName); i++)
                                        dt.Rows[i].Delete();

                                    dt.AcceptChanges();
                                }

                            ds.Tables.Add(dt);
                        }

                        cmd = null;
                        conn.Close();
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

            return ds;
        }
        
        private string getConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            if (FileType == ExcelExt.XLSX)
            {
                // XLSX - Excel 2007, 2010, 2012, 2013
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
                props["Extended Properties"] = "'Excel 12.0 XML;HDR=No;IMEX=1;'";
                props["Data Source"] = FilePhysical;
            }

            if (FileType == ExcelExt.XLS)
            {
                // XLS - Excel 2003 and Older
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "'Excel 8.0;HDR=NO;IMEX=1;'";
                props["Data Source"] = FilePhysical;
            }

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }
    }
}