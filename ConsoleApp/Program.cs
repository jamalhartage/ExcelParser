using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Processor;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //Test files located in 'test_data' folder

            //[Normal Run]
            //-----------------------
            
            DataSet dsNPOIxlsx = ExcelProcessor.GetDataNPOI(getPhysicalFilename("test_data_1.xlsx"));
            DataSet dsNPOIxls = ExcelProcessor.GetDataNPOI(getPhysicalFilename("test_data_1.xls"));
            
            DataSet dsOLEDBxlsx = ExcelProcessor.GetDataOLEDB(getPhysicalFilename("test_data_1.xlsx"));
            DataSet dsOLEDBxls = ExcelProcessor.GetDataOLEDB(getPhysicalFilename("test_data_1.xls"));


            //[Populate column names from row (header row)]
            //-----------------------

            Dictionary<string, int> headersNPOI = new Dictionary<string, int>();
            headersNPOI.Add("Source Data", 0);
            headersNPOI.Add("Sample PivotTable Report", 4);  
            DataSet dsNPOIWithHeader = ExcelProcessor.GetDataNPOI(
                getPhysicalFilename("test_data_1.xlsx"),
                headersNPOI
            );

            Dictionary<string, int> headersOLEDB = new Dictionary<string, int>();
            headersOLEDB.Add("Source Data", 0);
            headersOLEDB.Add("Sample PivotTable Report", 1);
            DataSet dsOLEDBWithHeader = ExcelProcessor.GetDataOLEDB(
                getPhysicalFilename("test_data_1.xlsx"),
                headersOLEDB
            );


            //[Swapping file extensions (if file has wrong extension app will swap an retry)]
            //-----------------------

            DataSet dsNPOIxlsSwap = ExcelProcessor.GetDataNPOI(getPhysicalFilename("test_data_2_actually_xls.xlsx"));
            DataSet dsNPOIxlsxSwap = ExcelProcessor.GetDataNPOI(getPhysicalFilename("test_data_2_actually_xlsx.xls"));

            DataSet dsOLEDBxlsSwap = ExcelProcessor.GetDataOLEDB(getPhysicalFilename("test_data_2_actually_xls.xlsx"));
            DataSet dsOLEDBxlsxSwap = ExcelProcessor.GetDataOLEDB(getPhysicalFilename("test_data_2_actually_xlsx.xls"));
        }

        private static string getPhysicalFilename(string filename)
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string nameOfExe = AppDomain.CurrentDomain.FriendlyName;

            int rootIdx = baseDirectory.LastIndexOf(nameOfExe);

            string path = Path.Combine(baseDirectory.Substring(0, rootIdx), nameOfExe);
            path = Path.Combine(path, "test_data");

            return Path.Combine(path, filename);
        }
    }
}
