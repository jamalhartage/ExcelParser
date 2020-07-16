using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Processor.Workers
{
    public enum ExcelExt { Unknown, XLS, XLSX };

    public abstract class WorkerBase
    {
        private string _filePhysical = null;
        private Dictionary<string, int> _headerRowInfo = null;
        private ExcelExt _fileType = ExcelExt.Unknown;

        public WorkerBase(string filePhysical) :
                this(filePhysical, null)
        {}
        
        public WorkerBase(string filePhysical, Dictionary<string, int> headerRowInfo)
        {
            if(File.Exists(filePhysical))
                _filePhysical = filePhysical;

            if(headerRowInfo != null && headerRowInfo.Count > 0)
            {
                _headerRowInfo = new Dictionary<string, int>();

                foreach(KeyValuePair<string, int> item in headerRowInfo)
                    _headerRowInfo.Add(item.Key.ToLower(), item.Value);
            }

            _fileType = ExcelExt.Unknown;
        }

        public abstract DataSet GetData();


        protected ExcelExt FileType
        {
             get
             {
                if(this._fileType == ExcelExt.Unknown)
                {
                    if(this.FilePhysical.EndsWith(".xlsx", StringComparison.Ordinal))
                        this._fileType = ExcelExt.XLSX;

                    if(this.FilePhysical.EndsWith(".xls", StringComparison.Ordinal))
                        this._fileType = ExcelExt.XLS;
                }

                if(this._fileType == ExcelExt.Unknown)
                    this._fileType = ExcelExt.XLS;

                return this._fileType;
            }

            set
            {
                this._fileType = value;
            }
        }

        protected bool HasUserDefinedHeaderRow(string sheetName)
        {
            if(this.GetHeaderRowIdx(sheetName) >= 0)
                return true;
            
            return false;
        }

        protected int GetHeaderRowIdx(string sheetName)
        {
            if(_headerRowInfo != null)
                if(sheetName != null)
                    if(_headerRowInfo.ContainsKey(sheetName.ToLower()))
                        return _headerRowInfo[sheetName.ToLower()];

            return -1;
        }

        protected string FilePhysical
        {
            get { return _filePhysical; }
        }
        
    }
}