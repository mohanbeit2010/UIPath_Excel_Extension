using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class NumberFormatColumnValues : CodeActivity
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
        [Description("Specify the ColumnName To Be Formatted. eg., 'B:B'")]
        public InArgument<string> In_Str_ColumnName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Specify Format As String eg., '#,##0.00'")]
        public InArgument<string> In_Str_NumberFormat { get; set; }

        [Category("Input")]
        [Description("Specify the row in which the column header is present in sheet as int.")]
        public InArgument<int> In_Int_Destination_headerRow { get; set; } = 1;

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.NumberFormatColumnValues(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_ColumnName.Get(context),
                In_Str_NumberFormat.Get(context),
                In_Int_Destination_headerRow.Get(context)
            );
        }
    }
}