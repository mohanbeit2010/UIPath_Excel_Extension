using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class InsertColumnsBefore : CodeActivity
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
        [Description("Pass the ColumnHeaderRow Number to which we find the LastRow. eg, '1'")]
        public InArgument<int> In_Int_ColumnHeaderRow { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the ColumnNames that need to be inserted in the ExcelWorksheet. eg., {\"Column1\",\"Column2\"}")]
        public InArgument<string[]> In_StrAry_NewColumnNames { get; set; }

        [Category("Input")]
        [Description("Pass the ColumnName to which new column to be inserted before it.\nif it is null then new column will be appended after Last Column. eg., \"Name\"")]
        public InArgument<string> In_Str_ExistingColumnName { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.InsertColumnsBefore(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Int_ColumnHeaderRow.Get(context),
                In_StrAry_NewColumnNames.Get(context),
                In_Str_ExistingColumnName.Get(context)
            );
        }
    }
}