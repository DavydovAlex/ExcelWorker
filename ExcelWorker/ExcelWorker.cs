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
        string _fullName;
        int _workSheet;
        int _headerRow;

        int _rows;
        int _columns;


        Application _objExcel;
        Workbook _objWorkBook;
        Worksheet _objWorkSheet;
        Range _range;

        public string FileNameWithoutExtension
        {
            get { return Path.GetFileNameWithoutExtension(_fullName); }
        }

        public string FileName
        {
            get { return Path.GetFileName(_fullName); }
        }

        public string Extension
        {
            get { return Path.GetExtension(_fullName); }
        }

        public int HeaderRow
        {
            get { return _headerRow; }
            set { _headerRow = value; }
        }
        public int Rows
        {
            get { return _rows; }
        }

        public string FullName
        {
            get { return _fullName; }
            //set { _fullName = value; }
        }

        public Worksheet Worksheet
        {
            get { return _objWorkSheet; }
        }

        public Application ObjExcel
        {
            get { return _objExcel; }
            //set { _objExcel = value; }
        }

        public Workbook Workbook
        {
            get { return _objWorkBook; }
            //set { _objWorkBook = value; }
        }
        public ExcelWorker(string fullName, int workSheet = 1, int headerRow=1,bool readOnly=true)
        {
            _fullName = fullName;
            _workSheet = workSheet;
            _headerRow = headerRow;

            _OpenFile(readOnly);
            _rows = _range.Rows.Count;
            _columns = _range.Columns.Count;

        }
        public ExcelWorker()
        {

        }

        //https://msdn.microsoft.com/de-de/library/microsoft.office.interop.excel.workbooks.open(v=office.11).aspx
        public bool Open(string Filename,
                            int UpdateLinks=0,
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
                _objExcel = new Application();
                _objWorkBook = _objExcel.Workbooks.Open(Filename,UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, 
                                                        Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad);
                _ReadWorkSheet();
                return true;
            }
            catch
            {
                return false;
            }
            
        }


        private void _OpenFile(bool readOnly)
        {
            _objExcel = new Application();
            _objWorkBook = _objExcel.Workbooks.Open(Filename: _fullName, ReadOnly: readOnly);
            _ReadWorkSheet();
        }
        public void CreateFile()
        {
            _objExcel = new Application();

            //Книга.
            _objWorkBook = _objExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            _objWorkSheet = (Worksheet)_objWorkBook.Worksheets.Item[1];
        }

        private void _ReadWorkSheet()
        {
            _objWorkSheet = (Worksheet)_objWorkBook.Sheets[_workSheet];
            _range = _objWorkSheet.UsedRange;
            
        }


        public string[] ReadRow(int row)
        {
            string[] values = new string[_columns];
            for (int cCnt = 1; cCnt <= _columns; cCnt++)
            {
                values[cCnt - 1] = ReadCell(row,cCnt);
            }
            return values;
        }

        public string ReadCell(int row, int column)
        {
            return Convert.ToString((_range.Cells[row, column] as Range).Value2 == null ? "" : (_range.Cells[row, column] as Range).Value2);
        }
        public void WriteCell(int row, int column,string value)
        {
            _objWorkSheet.Cells[row, column] = value;
        }


        public string[] GetHeader()
        {
            return ReadRow(_headerRow);
        }


        public void CloseFile()
        {
            _objExcel.Quit();
        }

        ~ExcelWorker()
        {
            CloseFile();
        }
    }
}
