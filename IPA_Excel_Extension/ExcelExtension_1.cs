using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace IPA_Excel_Extension
{
    public class ExcelExtension_1
    {
        public static Excel.Application getExcelApplication()
        {
            Excel.Application oExApp = new Excel.Application();
            try
            {
                oExApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            }
            catch (Exception e)
            {

                if (oExApp == null)
                    throw new Exception("Unable to Find Excel Application.Launch Excel First", e)
                    {
                        Source = "ExcelExtension.getExcelApplication"
                    };
            }
            oExApp.Visible =true;
            return oExApp;
        }

        public static Excel.Workbook getWorkbookObject(string Path)
        {
            Excel.Workbook oWB = null;

            return oWB;
        }


        public static Excel.Workbook OpenExcelDocument(string Path)
        {
            Excel.Workbook oWB = null;

            return oWB;
        }


        public static Excel.Worksheet GetWorksheet(string Path,string SheetName)
        {
            Excel.Worksheet oWorksheet = null;
               
            return oWorksheet;
        }


        public static void SaveAs(string CurrentPath,string Newpath,Excel.XlFileFormat Format)
        {

        }


        public static void DeleteSheet(string Path, string SheetName)
        {

        }

        public static void NumberFormatRangeValues(string Path, string SheetName,string Range,string Format)
        {

        }

        public static void NumberFormatColumnValues(string Path, string SheetName, string ColumnName, string Format,int ColumnHeaderRow)
        {

        }

        public static void FormatRangeStyle(string Path, string SheetName, string Range, string Style)
        {

        }


        public static void FormatColumnStyle(string Path, string SheetName, string ColumnName, string Style,int ColumnHeaderRow)
        {

        }

        public static Excel.Range Find(string Path, string SheetName, string FindInRange, string FindAfterRange, string FindValue,Excel.XlFindLookIn LookIn,Excel.XlLookAt LookAt,Excel.XlSearchOrder SearchOrder,Excel.XlSearchDirection SerachDirection)
        {
            Excel.Range r = null;

            return r;
        }

        public static void CopyPasteRange()
        {

        }
        
        public static void FindAndReplace()
        {

        }

        public static void TextOfColumn()
        {

        }

        public static string FindColumnAddress()
        {
            string Address= string.Empty;

            return Address;
        }


        public static void VLookUpRange()
        {

        }


        public static long FindLastRow(string Path,string sheetName,string columnName,int columnHeaderRow)
        {
            long lastRow = 0;
            return lastRow;
        }

        public static void UnFilter(string Path,String SheetName)
        {

        }

        public static void FilterRange(string Path,String SheetName, Dictionary<string, string[]> ColumnNameWithValues, int ColumnHeaderRow,bool DisableCurrentFilter)
        {

        }

        public static void InsertColumnsBefore(string Path,string SheetName,int ColumnHeaderRow, string[] NewColumnNames,string ExistingColumnName)
        {

        }

        public static void CopyPasteColumns(string SourcePath,string SourceSheetName,string DestinationPath,string DestinationSheetName, string[] ColumnNames,int SourceColumnHeaderRow,int DestinationHeaderRow)
        {

        }

        public static long FindLastColumn(string Path,string SheetName,int ColumnHeaderRow)
        {
            return 0;
        }

        public static void MoveColumnBeforeColumn(string Path, string SheetName,string ColumntoMove,int ColumnHeaderRow,string ColumntoMoveBefore)
        {

        }
        public static void CloseExcelDocument(string Path)
        {

        }


        //public static long FindLastRow(string Path, string SheetName, string ColumnName, int ColumnHeaderRow)
        //{
        //    Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);
        //    string columnAddress;
        //    long lastRow = 0;

        //    if (string.IsNullOrEmpty(ColumnName))
        //    {
        //        columnAddress = "A";
        //    }
        //    else
        //    {
        //        columnAddress = ExcelExtension.FindColumnAddress(Path, SheetName, ColumnName, ColumnHeaderRow);
        //    }

        //    if (!string.IsNullOrEmpty(columnAddress))
        //    {
        //        lastRow = oWs.Cells[oWs.Rows.Count, columnAddress].End(Excel.XlDirection.xlUp).Row;
        //    }

        //    Marshal.ReleaseComObject(oWs);
        //    return lastRow;
        //}

        //public static void UnFilter(string Path, string SheetName)
        //{
        //    Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);
        //    oWs.AutoFilterMode = false;
        //}

    }
}
