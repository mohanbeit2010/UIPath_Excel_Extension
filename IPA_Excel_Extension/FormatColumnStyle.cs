using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class FormatColumnStyle : CodeActivity
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
        [Description("Specify the ColumnName To Be Styled. eg., {\"B\",\"B\"}")]
        public InArgument<string> In_Str_ColumnName { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Specify Style As String (eg, 'Comma')")]
        public InArgument<string> In_Str_Style { get; set; }

        [Category("Input")]
        [Description("Specify the row in which the column header is present in sheet as int.")]
        public InArgument<int> In_Int_Destination_headerRow { get; set; } = 1;

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.FormatColumnStyle(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Str_ColumnName.Get(context),
                In_Str_Style.Get(context),
                In_Int_Destination_headerRow.Get(context)
            );
        }
    }
}