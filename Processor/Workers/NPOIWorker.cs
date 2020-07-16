using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Processor.Workers
{
    public class NPOIWorker : WorkerBase
    {
        public NPOIWorker(string filePhysical) :
                base(filePhysical)
        {}

        public NPOIWorker(string filePhysical, Dictionary<string, int> headerRowInfo) :
                base(filePhysical, headerRowInfo)
        {}

        public override DataSet GetData()
        {
            return datasetFromExcel_NPOI();
        }

        public IWorkbook GetWorkbook()
        {
            IWorkbook wb = getExcelWorkbookWorker();

            return wb;
        }

        private IWorkbook getExcelWorkbookWorker()
        {
            IWorkbook wb = null;

            if ((wb == null) && (FileType == ExcelExt.XLS))
            {
                wb = getWorkbookXLS();

                if (wb == null)
                {
                    FileType = ExcelExt.XLSX;
                    wb = getWorkbookXLSX();

                    if(wb == null)
                        FileType = ExcelExt.Unknown;
                }
            }
            else if ((wb == null) && (FileType == ExcelExt.XLSX))
            {
                wb = getWorkbookXLSX();

                if (wb == null)
                {
                    FileType = ExcelExt.XLS;
                    wb = getWorkbookXLS();

                    if(wb == null)
                        FileType = ExcelExt.Unknown;
                }
            }

            return wb;
        }

        private IWorkbook getWorkbookXLS()
        {
            IWorkbook wb = null;

            try
            {
                if(!String.IsNullOrWhiteSpace(FilePhysical))
                    if(File.Exists(FilePhysical))
                        using (FileStream fs = new FileStream(FilePhysical, FileMode.Open, FileAccess.Read))
                        {
                            fs.Position = 0;

                            wb = new HSSFWorkbook(fs);

                            fs.Close();
                        }
            }
            catch (Exception)
            {
                wb = null;
            }

            return wb;
        }

        private IWorkbook getWorkbookXLSX()
        {
            IWorkbook wb = null;

            try
            {
                if(!String.IsNullOrWhiteSpace(FilePhysical))
                    if(File.Exists(FilePhysical))
                        using (FileStream fs = new FileStream(FilePhysical, FileMode.Open, FileAccess.Read))
                        {
                            fs.Position = 0;

                            wb = new XSSFWorkbook(fs);

                            fs.Close();
                        }
            }
            catch (Exception)
            {
                wb = null;
            }

            return wb;
        }

        private DataSet datasetFromExcel_NPOI()
        {
            IWorkbook wb = GetWorkbook();

            if(wb == null)
                return null;

            DataSet ds = new DataSet();

            try
            {
                for (int q = 0; q < wb.NumberOfSheets; q++)
                {
                    ISheet sh = (ISheet)wb.GetSheetAt(q);
                    String sheetName = sh.SheetName;

                    DataTable dt = new DataTable(sheetName);
                    
                    int sheetRowCount = sh.LastRowNum;
                    int sheetColCount = 0;
                    for (int z = 0; z <= sheetRowCount; z++)
                    {
                        IRow sheetRow = sh.GetRow(z);

                        if (sheetRow == null)
                            continue;

                        int lastCellColIdx = (int)sheetRow.LastCellNum;

                        if (lastCellColIdx > sheetColCount)
                            sheetColCount = lastCellColIdx;
                    }

                    //get header row
                    IRow headerRow =
                        HasUserDefinedHeaderRow(sheetName) ? 
                            sh.GetRow(GetHeaderRowIdx(sheetName)) : 
                            null;

                    //add cols to dt
                    for (int j = 0; j < sheetColCount; j++)
                    {
                        string value = null;

                        if (headerRow != null)
                            value = getCellValue(headerRow.GetCell(j));

                        if (String.IsNullOrWhiteSpace(value))
                            value = "_BlankCol_" + j.ToString();

                        dt.Columns.Add(value, typeof(string));
                    }

                    int dataStartingRow =
                        HasUserDefinedHeaderRow(sheetName) ?
                        GetHeaderRowIdx(sheetName) + 1 :
                        0;

                    //add data
                    for (int z = dataStartingRow; z <= sheetRowCount; z++)
                    {
                        if (sh.GetRow(z) == null)
                        {
                            dt.Rows.Add(dt.NewRow());
                            continue;
                        }

                        // add row
                        DataRow row = dt.NewRow();

                        List<ICell> cells = sh.GetRow(z).Cells;

                        if (cells != null)
                        {
                            // write row value
                            foreach (ICell cell in sh.GetRow(z).Cells)
                            {
                                int cellIdx = cell.ColumnIndex;

                                row[cellIdx] = getCellValue(cell);
                            }
                        }

                        dt.Rows.Add(row);
                    }

                    dt.AcceptChanges();

                    ds.Tables.Add(dt);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

            return ds;
        }

        private string getCellValue(ICell cell)
        {
            if(cell == null)
                return null;

            string rtnVal = String.Empty;

            switch (cell.CellType)
            {
                case NPOI.SS.UserModel.CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try
                        {
                            ICellStyle style = cell.CellStyle;
                            string excelFormat = style.GetDataFormatString(); //returns format excel date is in 

                            double cellValue = double.Parse(cell.NumericCellValue.ToString());
                            
                            DateTime cellDateTime = DateTime.FromOADate(cellValue);


                            string rtnFormat = string.Empty;

                            if(cellValue >= 1 || cellValue < 0)
                                rtnFormat = rtnFormat + " MM/dd/yyyy";

                            if((cellValue % 1) != 0)
                                rtnFormat = rtnFormat + " hh:mm:ss tt";

                            rtnVal = cellDateTime.ToString(rtnFormat.Trim());
                        }
                        catch(Exception)
                        {
                            rtnVal = cell.NumericCellValue.ToString();
                        }
                    }
                    else
                    {
                        rtnVal = cell.NumericCellValue.ToString();
                    }
                    break;

                case NPOI.SS.UserModel.CellType.String:
                    rtnVal = cell.StringCellValue;
                    break;

                case NPOI.SS.UserModel.CellType.Boolean:
                    rtnVal = cell.BooleanCellValue.ToString();
                    break;

                case NPOI.SS.UserModel.CellType.Formula:
                    switch (cell.CachedFormulaResultType)
                    {
                        case NPOI.SS.UserModel.CellType.Numeric:
                            rtnVal = cell.NumericCellValue.ToString();
                            break;

                        case NPOI.SS.UserModel.CellType.String:
                            rtnVal = cell.StringCellValue;
                            break;

                        case NPOI.SS.UserModel.CellType.Boolean:
                            rtnVal = cell.BooleanCellValue.ToString();
                            break;

                        default:
                            rtnVal = cell.ToString();
                            break;
                    }
                    break;

                case NPOI.SS.UserModel.CellType.Blank:
                    rtnVal = String.Empty;
                    break;

                case NPOI.SS.UserModel.CellType.Unknown:
                    rtnVal = String.Empty;
                    break;

                default:
                    rtnVal = cell.ToString();
                    break;
            }

            return rtnVal;
        }
    }
}