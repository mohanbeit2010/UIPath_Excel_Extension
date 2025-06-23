using System.Activities;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace IPA_Excel_Extension
{
    public class SaveAs : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel Workbook Path With Extension. eg., test.xlsx")]
        public InArgument<string> In_Str_ExcelWorkbookPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel WorkSheet Name As String. eg: Sheet1")]
        public InArgument<string> In_Str_NewWorkbookPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Specify the Format")]
        public InArgument<Excel.XlFileFormat> In_Format { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.SaveAs(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_NewWorkbookPath.Get(context),
                In_Format.Get(context)
            );
        }
    }
}