using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class FindLastColumn : CodeActivity
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
        [Description("Pass the ColumnHeaderRow Number to which we find the LastColumn. eg, '1'")]
        public InArgument<int> In_Int_ColumnHeaderRow { get; set; }

        [Category("Output")]
        [Description("Returns the LastColumn for the given Sheet. eg, '150'")]
        public OutArgument<long> Out_Lng_LastColumn { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Out_Lng_LastColumn.Set(context, ExcelExtension.FindLastColumn(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Int_ColumnHeaderRow.Get(context)
            ));
        }
    }
}