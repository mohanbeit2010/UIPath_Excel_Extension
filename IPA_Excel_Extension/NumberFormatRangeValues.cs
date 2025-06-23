using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class NumberFormatRangeValues : CodeActivity
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
        [Description("Specify the Range To Be Formatted. eg., 'B:B'")]
        public InArgument<string> In_Str_Range { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Specify Format As String eg., '#,##0.00'")]
        public InArgument<string> In_Str_NumberFormat { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.NumberFormatRangeValues(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_Range.Get(context),
                In_Str_NumberFormat.Get(context)
            );
        }
    }
}