using System.Activities;
using System.ComponentModel;




namespace IPA_Excel_Extension
{
    
    public class DeleteSheet : CodeActivity
    {
        
        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel Workbook Path with Extension. eg., test.xlsx")]
        public InArgument<string> In_Str_ExcelWorkbookPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the Excel WorkSheet Name As String. eg: Sheet1")]
        public InArgument<string> In_Str_SheetName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.DeleteSheet(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context)
                );
        }
    }
}