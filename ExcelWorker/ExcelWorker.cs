using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Exel_XML.Classes
{
    class ExcelWorker
    {
        string _fileName;
        int _updateLinks;
        bool _readOnly;
        int _format;
        string _password;
        string _writeResPassword;
        bool _ignoreReadOnlyRecommended;
        XlPlatform _origin;
        string _delimiter;
        bool _editable;
        bool _notify;
        bool _converter;
        bool _addToMru;
        bool _local;
        bool _corruptLoad;


        //Sheet number
        int _workSheet;


        Application _objExcel;
        Workbook _objWorkBook;
        Worksheet _objWorkSheet;


        public int Rows
        {
            get { return _objWorkSheet.UsedRange.Rows.Count; }
        }
        public int Columns
        {
            get { return _objWorkSheet.UsedRange.Columns.Count; }
        }


        public string FileName
        {
            get { return _fileName; }          
        }


        public ExcelWorker(string FileName)
        {
            _objExcel = new Application();
            _workSheet = 1;
            _fileName = FileName;

            _updateLinks = 0;
            _readOnly = true;
            _format = 5;
            _password = "";
            _writeResPassword = "";
            _ignoreReadOnlyRecommended = true;
            _origin = XlPlatform.xlWindows;
            _delimiter = "";
            _editable = false;
            _notify = false;
            _converter = false;
            _addToMru = false;
            _local = false;
            _corruptLoad = false;
        }




        //https://msdn.microsoft.com/de-de/library/microsoft.office.interop.excel.workbooks.open(v=office.11).aspx
        public void Open(int UpdateLinks=0,
                            bool ReadOnly=true,
                            int Format=5,
                            string Password="",
                            string WriteResPassword="",
                            bool IgnoreReadOnlyRecommended=true,
                            XlPlatform Origin = XlPlatform.xlWindows,
                            string Delimiter="",
                            bool Editable=false,
                            bool Notify=false,
                            bool Converter=false,
                            bool AddToMru=false,
                            bool Local=false,
                            bool CorruptLoad=false)
        {
            try
            {
                
                _objWorkBook = _objExcel.Workbooks.Open(_fileName,_updateLinks, _readOnly, _format, _password, _writeResPassword, _ignoreReadOnlyRecommended, _origin, 
                                                        _delimiter, _editable, _notify, _converter, _addToMru, _local, _corruptLoad);

                _ReadWorkSheet();
                //return true;
            }
            catch(Exception e)
            {
                //throw new Exception(e.Message);
                //return false;
            }

        }


        public void Create()
        {
            //_objExcel = new Application();
            try
            {
                //Книга.
                _objWorkBook = _objExcel.Workbooks.Add(System.Reflection.Missing.Value);
                //Таблица.
                _objWorkSheet = (Worksheet)_objWorkBook.Worksheets.Item[1];
            }
            catch(Exception e)
            {
                throw new Exception(e.Message);
            }

        }

        private void _ReadWorkSheet()
        {
            _objWorkSheet = (Worksheet)_objWorkBook.Sheets[_workSheet];
            //_range = _objWorkSheet.UsedRange;
            
        }


        public object[] ReadRow(int row)
        {
            object[] values = new object[Columns];
            object[,] rowMatrix = GetRange(row,1,row,Columns);
            for (int cCnt = 1; cCnt <= Columns; cCnt++)
            {
                values[cCnt - 1] = rowMatrix[1, cCnt] ;//== null ? "" : Convert.ToString(rowMatrix[1, cCnt])
            }
            return values;
        }
        public object[,] GetRange(int minRow, int minColumn, int maxRow,int maxColumn )
        {
            Range range = _objWorkSheet.Range[_objWorkSheet.Cells[minRow, minColumn], _objWorkSheet.Cells[maxRow, maxColumn]];
            var matrixRow = (object[,])range.Value;


            return matrixRow;
        }
        
        public string ReadCell(int row, int column)
        {
            return Convert.ToString((_objWorkSheet.UsedRange.Cells[row, column] as Range).Value2 == null ? "" : (_objWorkSheet.UsedRange.Cells[row, column] as Range).Value2);
        }


        public void WriteCell(int row, int column,string value)
        {
            _objWorkSheet.Cells[row, column] = value;

        }


        public void WriteRow(int row, string[] values)
        {
            for (int cCnt = 1; cCnt <= Columns; cCnt++)
            {
                WriteCell(row,cCnt, values[cCnt-1]);
            }
        }


        public void Close()
        {
            //_objWorkBook.Close(false);
            _objExcel.Quit();
        }

        ~ExcelWorker()
        {
            Close();
        }
    }
}
