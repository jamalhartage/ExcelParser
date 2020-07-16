using System.Collections.Generic;
using System.Data;
using Processor.Workers;

namespace Processor
{
    /*
        Created by:  Jamal Hartage

        [Description]
            App will open excel using NPOI or OLEDB and convert to dataset

        [Run Info]
            HeaderRowInfo Dictionary consist of -->> [key: Sheetname], [value: col of header row (headerIdx)]
            If header info is sent program will use headerIdx as
            Datatable column name and only import row below (headerIdx + 1)

            OLEDB data begins adding DataRows with first row of data found in sheet.
            ex. if data on excel sheet start at row 5. That row will be row 0 in DataTable
    */


    static public class ExcelProcessor
    {
        static public DataSet GetDataNPOI(string filename)
        {
            DataSet ds = null;

            NPOIWorker npoiWorker = new NPOIWorker(filename);
            ds = npoiWorker.GetData();

            return ds;
        }

        static public DataSet GetDataNPOI(string filename, Dictionary<string, int> headerRowInfo)
        {
            DataSet ds = null;

            NPOIWorker npoiWorker = new NPOIWorker(filename, headerRowInfo);
            ds = npoiWorker.GetData();

            return ds;
        }

        static public DataSet GetDataOLEDB(string filename)
        {
            DataSet ds = null;

            OLEDBWorker oledbWorker = new OLEDBWorker(filename);
            ds = oledbWorker.GetData();

            return ds;
        }

        static public DataSet GetDataOLEDB(string filename, Dictionary<string, int> headerRowInfo)
        {
            DataSet ds = null;

            OLEDBWorker oledbWorker = new OLEDBWorker(filename, headerRowInfo);
            ds = oledbWorker.GetData();

            return ds;
        }
    }
}