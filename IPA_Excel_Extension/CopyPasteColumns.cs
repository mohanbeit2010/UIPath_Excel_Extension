using System.ComponentModel;
using System.Activities;


namespace IPA_Excel_Extension
{
    public class CopyPasteColumns : CodeActivity
    {
        [Category("Input")]
        [Description("Pass the Excel Source Workbook Path with Extension, eg: test.xlsx")]
        [RequiredArgument]
        public InArgument<string> In_SrcWorkbookPath { get; set; }

        [Category("Input")]
        [Description("Pass the excel Source WorkSheetName As String, eg: Sheet1")]
        [RequiredArgument]
        public InArgument<string> In_SrcSheetName { get; set; }

        [Category("Input")]
        [Description("Pass the Excel Destination Workbook Path with Extension, eg: test.xlsx")]
        [RequiredArgument]
        public InArgument<string> In_DstWorkbookPath { get; set; }

        [Category("Input")]
        [Description("Pass the excel Destination WorkSheetName As String, eg: Sheet1")]
        [RequiredArgument]
        public InArgument<string> In_DstSheetName { get; set; }

        [Category("Input")]
        [Description("Pass the ColumnNames that need to be copied in the ExcelWorksheet, eg: {\"Column1\",\"Column2\"}")]
        [RequiredArgument]
        public InArgument<string[]> In_strArray_ColumnNames { get; set; }

        [Category("Input")]
        [Description("Pass the Source ColumnHeaderRow Number in which we find the LastRow, eg: \"1\"")]
        [RequiredArgument]
        public InArgument<int> In_Int_SrcColumnHeaderRow { get; set; }

        [Category("Input")]
        [Description("Pass the Destination ColumnHeaderRow Number in which we find the LastRow, eg: \"1\"")]
        [RequiredArgument]
        public InArgument<int> In_Int_DstColumnHeaderRow { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.CopyPasteColumns(In_SrcWorkbookPath.Get(context),
                                             In_SrcSheetName.Get(context),
                                             In_DstWorkbookPath.Get(context),
                                             In_DstSheetName.Get(context),
                                             In_strArray_ColumnNames.Get(context),
                                             In_Int_SrcColumnHeaderRow.Get(context),
                                             In_Int_DstColumnHeaderRow.Get(context));
        }
    }
}