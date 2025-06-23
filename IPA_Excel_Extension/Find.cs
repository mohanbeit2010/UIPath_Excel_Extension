using System.Activities;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace IPA_Excel_Extension
{
    public class Find : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel Workbook Path with Extension. eg., test.xlsx")]
        public InArgument<string> In_Str_ExcelWorkbookPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the excel WorkSheet Name As String. eg: Sheet1")]
        public InArgument<string> In_Str_SheetName { get; set; }

        [Category("Input")]
        [Description("Specify the Range to search in")]
        public InArgument<string> In_Str_Range { get; set; }

        [Category("Input")]
        [Description("Specify the Range after which to search")]
        public InArgument<string> In_Str_SearchAfterRange { get; set; }

        [Category("Input")]
        [Description("The value to search")]
        public InArgument<string> In_Str_ValueToFind { get; set; }

        [Category("Input")]
        [Description("Find options")]
        public InArgument<Excel.XlFindLookIn> In_FindLookIn { get; set; }

        [Category("Input")]
        [Description("LookAt options")]
        public InArgument<Excel.XlLookAt> In_LookAt { get; set; }

        [Category("Input")]
        [Description("SearchOrder")]
        public InArgument<Excel.XlSearchOrder> In_SearchOrder { get; set; }

        [Category("Input")]
        [Description("SearchDirection")]
        public InArgument<Excel.XlSearchDirection> In_SearchDirection { get; set; }

        [Category("Output")]
        [Description("Address for the cell")]
        public OutArgument<string> Out_Str_Address { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Out_Str_Address.Set(context, ExcelExtension.Find(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_Range.Get(context),
                In_Str_SearchAfterRange.Get(context),
                In_Str_ValueToFind.Get(context),
                In_FindLookIn.Get(context),
                In_LookAt.Get(context),
                In_SearchOrder.Get(context),
                In_SearchDirection.Get(context)
            ).Address);
        }
    }
}