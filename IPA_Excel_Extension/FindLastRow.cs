using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class FindLastRow : CodeActivity
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
        [Description("Pass the ColumnName to which we find the LastRow eg., \"Name\"")]
        public InArgument<string> In_Str_ColumnName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the ColumnHeaderRow Number to which we find the LastRow. eg, '1'")]
        public InArgument<int> In_Int_ColumnHeaderRow { get; set; }

        [Category("Output")]
        [Description("Returns the LastRow for the given Column in the Sheet. eg, '150'")]
        public OutArgument<long> Out_Lng_LastRow { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            Out_Lng_LastRow.Set(context, ExcelExtension.FindLastRow(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_ColumnName.Get(context),
                In_Int_ColumnHeaderRow.Get(context)
            ));
        }
    }
}