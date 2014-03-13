using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Common.Office
{
    public class ExcelSheet
    {
        public string sheetName;
        public List<ExcelSheetRow> rowList;
        public List<ExcelSheetColumn> columnList;

        public ExcelSheet(string fileName, int sheetIndex)
        {
            if (sheetIndex < 1) sheetIndex = 1;

            //操作Excel的变量
            Excel.Application xApp = null;
            Excel.Worksheet xSheet = null;
            Excel.Workbook xBook = null;

            try
            {
                //创建Application对象
                xApp = new Excel.ApplicationClass();
                xApp.Visible = false;
                //得到WorkBook对象, 可以用两种方式之一: 下面的是打开已有的文件
                xBook = xApp.Workbooks._Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                //获取选定的表单
                xSheet = (Excel.Worksheet)xBook.Sheets[sheetIndex];
                this.sheetName = xSheet.Name;

                //获取列数
                int m_nColumn_Count = xSheet.UsedRange.Columns.Count;
                //获取行数
                int m_nRow_Count = xSheet.UsedRange.Rows.Count;
                if (m_nRow_Count < 2)
                {
                    throw new ExcelHelperException("请至少有一行数据，第一行是抬头");
                }

                this.rowList = new List<ExcelSheetRow>();
                this.columnList = new List<ExcelSheetColumn>();

                for (int i = 1; i <= m_nColumn_Count; i++)
                {
                    ExcelSheetColumn column = new ExcelSheetColumn(i, GetValue(xSheet, 1, i).ToString());
                    this.columnList.Add(column);
                }

                // 循环取值
                for (int i = 2; i <= m_nRow_Count; i++)
                {
                    ExcelSheetRow row = new ExcelSheetRow(i);
                    this.rowList.Add(row);

                    foreach (ExcelSheetColumn column in this.columnList)
                    {
                        ExcelSheetCell cell = new ExcelSheetCell(row, column, GetValue(xSheet, i, column.index));
                        row.AddCell(cell);
                    }
                }
            }
            finally
            {
                xSheet = null;
                xBook = null;
                if (null != xApp)
                {
                    xApp.Quit(); //这一句是非常重要的，否则Excel对象不能从内存中退出
                    xApp = null;
                }
            }
        }

        public object GetValue(Excel.Worksheet sheet, int rowIndex, int columnIndex)
        {
            Excel.Range range = (Excel.Range)(sheet.Cells[rowIndex, columnIndex]);
            return range.Value2 != null ? range.Value2 : range.Text;
        }
    }

    public class ExcelSheetRow
    {
        public int index;
        public List<ExcelSheetCell> cellList;

        public ExcelSheetRow(int index)
        {
            this.index = index;
            this.cellList = new List<ExcelSheetCell>();
        }

        public void AddCell(ExcelSheetCell cell)
        {
            cellList.Add(cell);
        }

        public ExcelSheetCell GetCell(int columnIndex)
        {
            foreach (ExcelSheetCell cell in cellList)
            {
                if (columnIndex == cell.column.index)
                {
                    return cell;
                }
            }

            return null;
        }

        public ExcelSheetCell GetCell(string columnName)
        {
            foreach (ExcelSheetCell cell in cellList)
            {
                if (columnName.ToLower().Equals(cell.column.name.ToLower()))
                {
                    return cell;
                }
            }

            return null;
        }
    }

    public class ExcelSheetColumn
    {
        public int index;
        public string name;

        public ExcelSheetColumn(int index, string name)
        {
            this.index = index;
            this.name = name;
        }
    }

    public class ExcelSheetCell
    {
        public ExcelSheetRow row;
        public ExcelSheetColumn column;
        public object value;

        public ExcelSheetCell(ExcelSheetRow row, ExcelSheetColumn column, object value)
        {
            this.row = row;
            this.column = column;
            this.value = value;
        }

        public string GetString()
        {
            return Convert.ToString(value);
        }

        public int GetInt()
        {
            return Convert.ToInt32(value);
        }
    }

    public class ExcelHelperException : Exception
    {
        public ExcelHelperException(string msg) : base(msg) { }
    }

    public class ExcelHelper
    {
        

    }
}
