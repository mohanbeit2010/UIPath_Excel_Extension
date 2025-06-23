using System.Activities;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace IPA_Excel_Extension
{
    public class FindAndReplace : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel Workbook Path with Extension. eg., test.xlsx")]
        public InArgument<string> In_Str_ExcelWorkbookPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel WorkSheet Name As String. eg: Sheet1")]
        public InArgument<string> In_Str_SheetName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Specify the Range to search in")]
        public InArgument<string> In_Str_Range { get; set; }

        [Category("Input")]
        [Description("Specify the Range after which to search")]
        public InArgument<string> In_Str_SearchAfterRange { get; set; }

        [Category("Input")]
        [Description("Value to search")]
        public InArgument<string> In_Str_ValueToFind { get; set; }

        [Category("Input")]
        [Description("Value to Replace")]
        public InArgument<string> In_Str_ValueToReplace { get; set; }

        [Category("Find options")]
        [Description("XlFindLookAt")]
        public InArgument<Excel.XlLookAt> In_XlLookAt { get; set; }

        [Category("Find options")]
        [Description("XlSearchOrder")]
        public InArgument<Excel.XlSearchOrder> In_SearchOrder { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.FindAndReplace(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_Range.Get(context),
                In_Str_SearchAfterRange.Get(context),
                In_Str_ValueToFind.Get(context),
                In_Str_ValueToReplace.Get(context),
                In_XlLookAt.Get(context),
                In_SearchOrder.Get(context)
            );
        }
    }
}