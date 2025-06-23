using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace IPA_Excel_Extension
{
    public static partial class ExcelExtension
    {
        public static Excel.Application getExcelApplication()
        {
            Excel.Application oWebApp = new Excel.Application();
            try
            {
                oWebApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            }
            catch (Exception e)
            {
                if (oWebApp == null)
                {
                    throw new Exception("Unable to Find Excel Application. Launch Excel first.", e)
                    {
                        Source = "ExcelExtension.getExcelApplication"
                    };
                }
            }
            oWebApp.Visible = true;
            return oWebApp;
        }

        public static Excel.Workbook getWorkbookObject(string Path)
        {
            Excel.Application oWebApp = getExcelApplication();
            Excel.Workbook oWB = null;
            bool WorkBookFound = false;
            foreach (Excel.Workbook item in oWebApp.Workbooks)
            {
                if (item.FullName.Contains(Path))
                {
                    oWB = item;
                    WorkBookFound = true;
                    break;
                }
            }
            if (!WorkBookFound)
            {
                oWB = OpenExcelDocument(Path);
            }
            Marshal.ReleaseComObject(oWebApp);
            return oWB;
        }

        public static void MoveColumnBeforeColumn(string Path, string SheetName, string ColumnToMove, int ColumnHeaderRow, string ColumnToMovebefore)
        {
            //Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWs = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path, ReadOnly: true);
            oWs = workbook.Sheets.Item[SheetName] as Worksheet;

            Excel.Range ColumnToMoveRange = ExcelExtension.Find(Path, SheetName, oWs.Range[oWs.Cells[ColumnHeaderRow, 1], oWs.Cells[ColumnHeaderRow, oWs.Columns.Count]].Address, "A" + ColumnHeaderRow, ColumnToMove, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext);
            Excel.Range ColumnToMoveBeforeRange = ExcelExtension.Find(Path, SheetName, oWs.Range[oWs.Cells[ColumnHeaderRow, 1], oWs.Cells[ColumnHeaderRow, oWs.Columns.Count]].Address, "A" + ColumnHeaderRow, ColumnToMovebefore, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext);
            if (ColumnToMoveRange == null)
            {
                Marshal.ReleaseComObject(ColumnToMoveRange);
                Marshal.ReleaseComObject(oWs);
                throw new Exception("The column '" + ColumnToMove + "' is not available in the provided source.");
            }
            if (ColumnToMoveBeforeRange == null)
            {
                Marshal.ReleaseComObject(ColumnToMoveBeforeRange);
                Marshal.ReleaseComObject(oWs);
                throw new Exception("The column '" + ColumnToMovebefore + "' is not available in the provided source.");
            }
            ColumnToMoveRange.EntireColumn.Cut();
            ColumnToMoveBeforeRange.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            Marshal.ReleaseComObject(ColumnToMoveRange);
            Marshal.ReleaseComObject(ColumnToMoveBeforeRange);
            Marshal.ReleaseComObject(oWs);
        }

        public static long FindLastColumn(string Path, string SheetName, int ColumnHeaderRow)
        {
            long lastColumnNo = 0;
            //Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);
            //Marshal.ReleaseComObject(oWs);
            //lastColumnNo = oWs.Cells[ColumnHeaderRow, oWs.Columns.Count].End(Excel.XlDirection.xlToLeft).Column;
            //return lastColumnNo;

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Application();
                workbook = excelApp.Workbooks.Open(Path, ReadOnly: true);
                worksheet = workbook.Sheets.Item[SheetName] as Worksheet;

                lastColumnNo = worksheet.Cells[ColumnHeaderRow, worksheet.Columns.Count].End(Excel.XlDirection.xlToLeft).Column;

                //Excel.Range lastCell = worksheet.Cells.Find(
                //    "*",
                //    System.Reflection.Missing.Value,
                //    XlFindLookIn.xlFormulas,
                //    XlLookAt.xlPart,
                //    XlSearchOrder.xlByColumns,
                //    XlSearchDirection.xlPrevious,
                //    false,
                //    false,
                //    false);

                //int lastCol = lastCell?.Column ?? 0;
                return lastColumnNo;
            }
            finally
            {

                workbook?.Close(false);
                excelApp?.Quit();

                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

        }

        public static void CopyPasteColumns(string srcPath, string srcSheetName, string dstPath, string dstSheetName, string[] ColumnNames, int srcColumnHeaderRow, int dstColumnHeaderRow)
        {
            //Excel.Worksheet oWsrc = ExcelExtension.getExcelSheet(srcPath, srcSheetName);
            //Excel.Worksheet oWdst = ExcelExtension.getExcelSheet(dstPath, dstSheetName);

            Application srcexcelApp = null;
            Workbook srcworkbook = null;
            Worksheet oWsrc = null;

            srcexcelApp = new Application();
            srcworkbook = srcexcelApp.Workbooks.Open(srcPath);
            oWsrc = srcworkbook.Sheets.Item[srcSheetName] as Worksheet;


            Application dstexcelApp = null;
            Workbook dstworkbook = null;
            Worksheet oWdst = null;

            dstexcelApp = new Application();
            dstworkbook = dstexcelApp.Workbooks.Open(dstSheetName);
            oWdst = dstworkbook.Sheets.Item[srcSheetName] as Worksheet;


            long lastSrcRow = 0;
            long lastDstRow = 0;




            foreach (string ColName in ColumnNames)
            {
                //Source getDestination
                string srcColumnAddress = FindColumnAddress(srcPath, srcSheetName, ColName, srcColumnHeaderRow);
                long srcLastRow = FindLastRow(srcPath, srcSheetName, ColName, srcColumnHeaderRow);

                //#region getDestination
                Excel.Range dstColumnrange = ExcelExtension.Find(dstPath, dstSheetName, oWdst.Range[oWdst.Cells[dstColumnHeaderRow, 1],
            oWdst.Cells[dstColumnHeaderRow, oWdst.Columns.Count]].Address, "A" + dstColumnHeaderRow, ColName, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext);
                long dstLastCol = FindLastColumn(dstPath, dstSheetName, dstColumnHeaderRow);

                if (dstColumnrange == null)
                {
                    dstLastCol = dstLastCol + 1;
                    dstColumnrange = oWdst.Cells[dstColumnHeaderRow, dstLastCol];
                }

                if ((Excel.Range)oWdst.Cells[dstColumnrange.Row, dstColumnrange.Column + 1].Value2 == null ? ((Excel.Range)oWdst.Cells[dstColumnrange.Row, dstColumnrange.Column + 1]).Value2 : ColName) ;

                lastDstRow = FindLastRow(dstPath, dstSheetName, ColName, dstColumnHeaderRow);
                string dstRegion = FindColumnAddress(dstPath, dstSheetName, ColName, dstColumnHeaderRow);
                //#endregion

                Excel.Range SrcRange = oWsrc.Range[srcColumnAddress + srcColumnHeaderRow + ":" + srcColumnAddress + srcLastRow];
                Excel.Range DstRange = oWdst.Range[dstRegion + dstColumnHeaderRow + ":" + dstRegion + (lastDstRow - srcColumnHeaderRow + dstColumnHeaderRow)].Offset[0, 1].Resize[SrcRange.Rows.Count, 1].Copy();
                ((Excel.Range)oWdst.Cells[dstColumnrange.Row, dstColumnrange.Column + 1]).PasteSpecial(Excel.XlPasteType.xlPasteAll,
                Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                if (dstColumnrange != null) Marshal.ReleaseComObject(dstColumnrange);
            }

            Marshal.ReleaseComObject(oWsrc);
            Marshal.ReleaseComObject(oWdst);
        }

        public static void InsertColumnsBefore(string Path, string SheetName, int ColumnHeaderRow, string[] NewColumnNames, string ExistingColumnName)
        {
            //Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWs = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWs = workbook.Sheets.Item[SheetName] as Worksheet;


            long columnCount = 0;
            long lastColumn = oWs.Cells[ColumnHeaderRow, oWs.Columns.Count].End(Excel.XlDirection.xlToLeft).Column;
            if (string.IsNullOrEmpty(ExistingColumnName))
            {
                columnCount = oWs.Cells[ColumnHeaderRow, oWs.Columns.Count].End(Excel.XlDirection.xlToLeft).Column;
                int i = 1;
                foreach (var name in NewColumnNames)
                {
                    oWs.Cells[ColumnHeaderRow, columnCount + i] = name;
                    i++;
                }
            }
            else
            {
                Excel.Range ColumnRange = ExcelExtension.Find(Path, SheetName, oWs.Range[oWs.Cells[ColumnHeaderRow, 1], oWs.Cells[ColumnHeaderRow, oWs.Columns.Count]].Address, "A" + ColumnHeaderRow, ExistingColumnName, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext);
                if (ColumnRange != null)
                {
                    foreach (var name in NewColumnNames)
                    {
                        ColumnRange.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        ColumnRange.Offset[0, -1].Value2 = name;
                    }
                    Marshal.ReleaseComObject(ColumnRange);
                }
            }
            Marshal.ReleaseComObject(oWs);
        }

        public static void FilterRange(string Path, string SheetName, Dictionary<string, string[]> ColumnNamesWithValues, int ColumnHeaderRow, bool DisableCurrentFilter)
        {
            //Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWs = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWs = workbook.Sheets.Item[SheetName] as Worksheet;


            if (DisableCurrentFilter)
            {
                oWs.AutoFilterMode = false;
            }
            foreach (string ColumnName in ColumnNamesWithValues.Keys)
            {
                string[] FilterValue = ColumnNamesWithValues[ColumnName];
                Excel.Range Columnrange = ExcelExtension.Find(Path, SheetName, oWs.Range[oWs.Cells[ColumnHeaderRow, 1], oWs.Cells[ColumnHeaderRow, oWs.Columns.Count]].Address, "A" + ColumnHeaderRow, ColumnName, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext);
                Excel.Range VisibleRange = oWs.UsedRange;
                VisibleRange.AutoFilter(Columnrange.Column, FilterValue, Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                Marshal.ReleaseComObject(FilterValue);
                Marshal.ReleaseComObject(Columnrange);
            }
            Marshal.ReleaseComObject(oWs);
        }

        public static long FindLastRow(string Path, string SheetName, string ColumnName, int ColumnHeaderRow)
        {
            //Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWs = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path, ReadOnly: true);
            oWs = workbook.Sheets[SheetName] as Worksheet;

            string columnAddress;
            long lastRow = 0;
            if (string.IsNullOrEmpty(ColumnName)) { columnAddress = "A"; }
            else { columnAddress = ExcelExtension.FindColumnAddress(Path, SheetName, ColumnName, ColumnHeaderRow); }
            if (!string.IsNullOrEmpty(columnAddress))
            {
                lastRow = oWs.Cells[oWs.Rows.Count, columnAddress].End(Excel.XlDirection.xlUp).Row;
            }
            Marshal.ReleaseComObject(oWs);
            return lastRow;



            //Application excelApp = null;
            //Workbook workbook = null;
            //Worksheet worksheet = null;

            //try
            //{
            //    excelApp = new Application();
            //    workbook = excelApp.Workbooks.Open(Path, ReadOnly: true);
            //    worksheet = workbook.Sheets[SheetName] as Worksheet;

            //    // Find the last used row
            //    Excel.Range lastCell = worksheet.Cells.Find(
            //        "*",
            //        System.Reflection.Missing.Value,
            //        XlFindLookIn.xlFormulas,
            //        XlLookAt.xlPart,
            //        XlSearchOrder.xlByRows,
            //        XlSearchDirection.xlPrevious,
            //        false,
            //        false,
            //        false);

            //    int lastRow = lastCell?.Row ?? 0;
            //    return lastRow;
            //}
            //finally
            //{
            //    workbook?.Close(false);
            //    excelApp?.Quit();

            //    if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            //    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            //    if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            //}
        }

        public static void UnFilter(string Path, string SheetName)
        {

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWs = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path, ReadOnly: true);
            oWs = workbook.Sheets[SheetName] as Worksheet;

            //Excel.Worksheet oWs = ExcelExtension.getExcelSheet(Path, SheetName);
            oWs.AutoFilterMode = false;
        }

        public static void VLookUpRange(string SourcePath, string DestinationPath, string SourceSheetName, string DestinationSheetName, string SourceKeyColumnName, string DestinationKeyColumnName, string SourceColumnName, string DestinationColumnName, int SourceHeaderRowNumber = 1, int DestinationHeaderRowNumber = 1)
        {
            //Excel.Worksheet oWsSource = ExcelExtension.getExcelSheet(SourcePath, SourceSheetName);
            //Excel.Worksheet oWsDestination = ExcelExtension.getExcelSheet(DestinationPath, DestinationSheetName);

            Application SrcexcelApp = null;
            Workbook Srcworkbook = null;
            Worksheet oWsSource = null;

            SrcexcelApp = new Application();
            Srcworkbook = SrcexcelApp.Workbooks.Open(SourcePath);
            oWsSource = Srcworkbook.Sheets[SourceSheetName] as Worksheet;

            Application DstexcelApp = null;
            Workbook Dstworkbook = null;
            Worksheet oWsDestination = null;

            DstexcelApp = new Application();
            Dstworkbook = DstexcelApp.Workbooks.Open(DestinationPath);
            oWsDestination = Dstworkbook.Sheets[DestinationSheetName] as Worksheet;


            long DestinationLastRow = oWsDestination.UsedRange.Rows.Count;
            string DestKeyColumn = FindColumnAddress(DestinationPath, DestinationSheetName, DestinationKeyColumnName, DestinationHeaderRowNumber);
            string SourceKeyColumn = FindColumnAddress(SourcePath, SourceSheetName, SourceKeyColumnName, SourceHeaderRowNumber);
            string SourceColumn = FindColumnAddress(SourcePath, SourceSheetName, SourceColumnName, SourceHeaderRowNumber);
            string DestColumn = FindColumnAddress(DestinationPath, DestinationSheetName, DestinationColumnName, DestinationHeaderRowNumber);
            long columnIndex = oWsSource.Range[SourceColumn + "1"].Column - oWsSource.Range[SourceKeyColumn + "1"].Column + 1;

            string lookup = "=VLOOKUP(" +
                DestKeyColumn + (DestinationHeaderRowNumber + 1) +
                ", '" + SourcePath.Split('\\')[SourcePath.Split('\\').Count() - 1] +
                "'!" + SourceSheetName + "'!" +
                SourceKeyColumn + "1:" + SourceColumn + "1" +
                columnIndex +
                ",FALSE)";

            oWsDestination.Activate();
            long column = oWsDestination.Columns[DestColumn].Column;
            oWsDestination.Cells[DestinationHeaderRowNumber + 1, column].Formula = lookup;
            oWsDestination.Range[DestColumn + (DestinationHeaderRowNumber + 1)].Copy();
            oWsDestination.Range[DestColumn + (DestinationHeaderRowNumber + 1) + ":" + DestColumn + DestinationLastRow].SpecialCells(Excel.XlCellType.xlCellTypeVisible).PasteSpecial(Excel.XlPasteType.xlPasteAll);
            Marshal.ReleaseComObject(oWsSource);
            Marshal.ReleaseComObject(oWsDestination);
        }

        public static string FindColumnAddress(string Path, string SheetName, string ColumnName, int ColumnHeaderRow)
        {
            //Excel.Worksheet oWS = ExcelExtension.getExcelSheet(Path, SheetName);

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            Excel.Range findRange = ExcelExtension.Find(Path, SheetName, oWS.Range[oWS.Cells[ColumnHeaderRow, 1], oWS.Cells[ColumnHeaderRow, oWS.Columns.Count]].Address,
                                    "A" + ColumnHeaderRow.ToString(), ColumnName, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                                    Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext);

            string range = string.Empty;
            if (findRange != null)
            {
                range = findRange.Address;
                range = range.Split('$')[1];
                range = range.Split(':')[0];
            }

            //string range = ExcelExtension.find(Path, SheetName, oWS.Range[oWS.Cells[ColumnHeaderRow, 1], oWS.Cells[ColumnHeaderRow, oWS.Columns.Count]].Address, 
            //    "A" + ColumnHeaderRow, ColumnName, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, 
            //    Excel.XlSearchDirection.xlNext).Address;

            //range = range.Split('$')[1];
            //range = range.Split(':')[0];

            if (findRange != null) Marshal.ReleaseComObject(findRange);
            Marshal.ReleaseComObject(oWS);

            return range;
        }

        public static void FindAndReplace(string Path, string SheetName, string FindInRange, string FindAfterRange, string FindValue, string ValueToReplace, Excel.XlLookAt LookAt, Excel.XlSearchOrder SearchOrder)
        {
            //Excel.Worksheet oWS = ExcelExtension.getExcelSheet(Path, SheetName);


            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            oWS.Range[FindInRange].Replace(FindValue, ValueToReplace, LookAt, SearchOrder, false);

            Marshal.ReleaseComObject(oWS);
        }

        public static void TextToColumn(string Path, string SheetName, string RangeToSplit, string destinationRange, Excel.XlTextParsingType DataType,
                                        Excel.XlTextQualifier TextQualifier, string Delimiter)
        {
            //Excel.Worksheet oWS = ExcelExtension.getExcelSheet(Path, SheetName);


            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            oWS.Range[RangeToSplit].TextToColumns(destinationRange, DataType, TextQualifier, true,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Delimiter, Type.Missing, Type.Missing);
        }


        public static void CopyPasteRange(string SourcePath, string SourceSheetName, string SourceRange,
                                  string DestinationPath, string DestinationSheetName, string DestinationRange,
                                  Excel.XlPasteType PasteType)
        {
            //Excel.Worksheet SourceoWs = ExcelExtension.getExcelSheet(SourcePath, SourceSheetName);
            //Excel.Worksheet DestinationoWs = ExcelExtension.getExcelSheet(DestinationPath, DestinationSheetName);

            Application SrcexcelApp = null;
            Workbook Srcworkbook = null;
            Worksheet SourceoWs = null;

            SrcexcelApp = new Application();
            Srcworkbook = SrcexcelApp.Workbooks.Open(SourcePath);
            SourceoWs = Srcworkbook.Sheets[SourceSheetName] as Worksheet;


            Application DstexcelApp = null;
            Workbook Dstworkbook = null;
            Worksheet DestinationoWs = null;

            DstexcelApp = new Application();
            Dstworkbook = DstexcelApp.Workbooks.Open(DestinationPath);
            DestinationoWs = Dstworkbook.Sheets[DestinationSheetName] as Worksheet;


            SourceoWs.Range[SourceRange].Copy();
            DestinationoWs.Range[DestinationRange].PasteSpecial(PasteType);
        }
        public static Excel.Range Find(string Path, string SheetName, string FindInRange, string FindAfterRange,
                                       string FindValue, Excel.XlFindLookIn LookIn, Excel.XlLookAt LookAt,
                                       Excel.XlSearchOrder SearchOrder, Excel.XlSearchDirection SearchDirection)
        {
            //Excel.Worksheet oWS = ExcelExtension.getExcelSheet(Path, SheetName);

            Application SrcexcelApp = null;
            Workbook Srcworkbook = null;
            Worksheet oWS = null;

            SrcexcelApp = new Application();
            Srcworkbook = SrcexcelApp.Workbooks.Open(Path);
            oWS = Srcworkbook.Sheets[SheetName] as Worksheet;


            Excel.Range range = oWS.Range[FindInRange].Find(FindValue, oWS.Range[FindAfterRange],
                                                             LookIn, LookAt, SearchOrder, SearchDirection, false, false);

            Marshal.ReleaseComObject(oWS);
            Marshal.ReleaseComObject(Srcworkbook);
            Marshal.ReleaseComObject(SrcexcelApp);

            return range;
        }

        public static void FormatColumnStyle(string Path, string SheetName, string ColumnName, string Style, int ColumnheaderRow)
        {
            // string Address = FindColumnAddress(Path, SheetName, ColumnName, ColumnheaderRow);
            string Address = string.Empty;

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            Excel.Range findRange = ExcelExtension.Find(Path, SheetName, oWS.Range[oWS.Cells[ColumnheaderRow, 1], oWS.Cells[ColumnheaderRow, oWS.Columns.Count]].Address,
                                    "A" + ColumnheaderRow.ToString(), ColumnName, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                                    Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext);

            string range = string.Empty;
            if (findRange != null)
            {
                range = findRange.Address;
                range = range.Split('$')[1];
                range = range.Split(':')[0];
            }
            Address = Address + ":" + Address;
            FormatRangeStyle(Path, SheetName, Address, Style);
        }


        public static void NumberFormatColumnValues(string Path, string SheetName, string ColumnName, string Format, int ColumnheaderRow)
        {
            string Address = FindColumnAddress(Path, SheetName, ColumnName, ColumnheaderRow);
            Address = Address + ":" + Address;
            NumberFormatRangeValues(Path, SheetName, Address, Format);
        }

        public static void FormatRangeStyle(string Path, string SheetName, string Range, string Style)
        {
            //Excel.Worksheet oWS = null;
            //oWS = getExcelSheet(Path, SheetName);

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            oWS.Range[Range].Style = Style;
            Marshal.ReleaseComObject(oWS);
        }
        public static void SaveAs(string currPath, string NewPath, Excel.XlFileFormat format)
        {
            //Excel.Workbook oWB = null;
            //oWB = getWorkBookObject(currPath);


            Application excelApp = null;
            Workbook oWB = null;


            excelApp = new Application();
            oWB = excelApp.Workbooks.Open(currPath);
            oWB.SaveAs(NewPath, format);
            Marshal.ReleaseComObject(oWB);
        }

        public static void DeleteSheet(string Path, string SheetName)
        {
            //Excel.Worksheet oWS = null;
            //oWS = getExcelSheet(Path, SheetName);

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            oWS.Delete();

            Marshal.ReleaseComObject(oWS);
        }

        public static void NumberFormatRangeValues(string Path, string SheetName, string Range, string Format)
        {
            //Excel.Worksheet oWS = null;
            //oWS = getExcelSheet(Path, SheetName);

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            oWS.Range[Range].NumberFormat = Format;
            Marshal.ReleaseComObject(oWS);
        }


        public static Excel.Workbook OpenExcelDocument(string Path)
        {
            Excel.Application oExApp = getExcelApplication();
            Excel.Workbook oWB = null;
            oWB = oExApp.Workbooks.Open(Path, false);
            Marshal.ReleaseComObject(oExApp);
            return oWB;
        }

        public static Excel.Worksheet getExcelSheet(string Path, string SheetName)
        {
            //Excel.Workbook oWB = null;
            //Excel.Worksheet oWS = null;

            //oWB = getWorkBookObject(Path);
            //oWS = oWB.Sheets[SheetName];

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);
            oWS = workbook.Sheets[SheetName] as Worksheet;

            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(oWS);
            return oWS;
        }

        public static Excel.Workbook getWorkBookObject(string Path)
        {
            //Excel.Application oExApp = getExcelApplication();
            //Excel.Workbook oWB = null;

            Application excelApp = null;
            Workbook workbook = null;
            Worksheet oWS = null;

            excelApp = new Application();
            workbook = excelApp.Workbooks.Open(Path);          

            bool WorkBookFound = false;

            foreach (Excel.Workbook item in excelApp.Workbooks)
            {
                if (item.FullName.Contains(Path))
                {
                    workbook = item;
                    WorkBookFound = true;
                    break;
                }
            }

            //if (!WorkBookFound)
            //{
            //    workbook = OpenExcelDocument(Path);
            //}

            Marshal.ReleaseComObject(excelApp);
            return workbook;
        }


        public static void TextToColumns(string excelFilePath, string sheetName, string rangeAddress, Excel.XlTextParsingType parsingType, Excel.XlTextQualifier textQualifier, object[] fieldInfo, string delimiter = "")
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false; // Set to true if you want to see Excel open

                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = (Excel.Worksheet)workbook.Sheets[sheetName];

                Excel.Range range = worksheet.Range[rangeAddress];

                // Perform Text to Columns
                range.TextToColumns(
                    Destination: range.Cells[1, 1], // The top-left cell of the destination range
                    DataType: parsingType,
                    TextQualifier: textQualifier,
                    ConsecutiveDelimiter: false, // Set to true if multiple delimiters should be treated as one
                    Tab: delimiter == "\t",      // Is tab the delimiter?
                    Semicolon: delimiter == ";", // Is semicolon the delimiter?
                    Comma: delimiter == ",",     // Is comma the delimiter?
                    Space: delimiter == " ",     // Is space the delimiter?
                    Other: !string.IsNullOrEmpty(delimiter) && delimiter != "\t" && delimiter != ";" && delimiter != "," && delimiter != " ",
                    OtherChar: delimiter,
                    FieldInfo: fieldInfo, // This is an array of arrays (e.g., new object[,] { {1, Excel.XlColumnDataType.xlGeneralFormat}, {2, Excel.XlColumnDataType.xlTextFormat} })
                    TrailingMinusNumbers: true // For numbers with trailing minus signs
                );

                workbook.Save();
                Console.WriteLine($"Text to Columns operation completed successfully on sheet '{sheetName}' in '{excelFilePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up Excel processes
                if (workbook != null)
                {
                    workbook.Close(false); // Close without saving if an error occurred
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                }
                // Optional: Garbage collect to ensure COM objects are released
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


    }
}
