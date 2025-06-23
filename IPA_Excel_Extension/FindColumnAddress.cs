using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class FindColumnAddress : CodeActivity
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
        [Description("Specify the Column Name")]
        public InArgument<string> In_Str_ColName { get; set; }

        [Category("Input")]
        [Description("Specify the row in which the column header is present in sheet as int.")]
        public InArgument<int> In_int_Destination_headerRow { get; set; } = 1;

        [Category("Output")]
        [Description("Address of the column")]
        public OutArgument<string> Out_str_ColumnAddress { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Out_str_ColumnAddress.Set(context, ExcelExtension.FindColumnAddress(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_ColName.Get(context),
                In_int_Destination_headerRow.Get(context)
            ));
        }
    }
}