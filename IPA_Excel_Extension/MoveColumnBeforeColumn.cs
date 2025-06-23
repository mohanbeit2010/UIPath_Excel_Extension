using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class MoveColumnBeforeColumn : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel Workbook Path With Extension. eg., test.xlsx")]
        public InArgument<string> In_Str_ExcelWorkbookPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel WorkSheet Name As String. eg: Sheet1")]
        public InArgument<string> In_Str_SheetName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the ColumnName that need to be moved in the ExcelWorksheet. eg, 'Column1'")]
        public InArgument<string> In_Str_ColumnToMove { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the ColumnHeaderRow Number to which we find the LastRow. eg, '1'")]
        public InArgument<int> In_Int_ColumnHeaderRow { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the ColumnName where column will be moved before it in the ExcelWorksheet. eg., 'Column1'")]
        public InArgument<string> In_Str_ColumnToMoveBefore { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.MoveColumnBeforeColumn(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_ColumnToMove.Get(context),
                In_Int_ColumnHeaderRow.Get(context),
                In_Str_ColumnToMoveBefore.Get(context)
            );
        }
    }
}