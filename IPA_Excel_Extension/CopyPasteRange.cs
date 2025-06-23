using System.ComponentModel;
using System.Activities;
using Microsoft.Office.Interop.Excel;

namespace IPA_Excel_Extension
{
    public class CopyPasteRange : CodeActivity
    {
        [Category("Source")]
        [Description("Pass the Excel Source Workbook Path with Extension, eg: test.xlsx")]
        [RequiredArgument]
        public InArgument<string> In_SrcWorkbookPath { get; set; }

        [Category("Source")]
        [Description("Pass the excel Source WorkSheetName As String, eg: Sheet1")]
        [RequiredArgument]
        public InArgument<string> In_SrcSheetName { get; set; }

        [Category("Source")]
        [Description("Pass the Source Range, eg: A1:B10")]
        [RequiredArgument]
        public InArgument<string> In_SrcRange { get; set; }

        [Category("Destination")]
        [Description("Pass the Excel Destination Workbook Path with Extension, eg: test.xlsx")]
        [RequiredArgument]
        public InArgument<string> In_DstWorkbookPath { get; set; }

        [Category("Destination")]
        [Description("Pass the excel Destination WorkSheetName As String, eg: Sheet1")]
        [RequiredArgument]
        public InArgument<string> In_DstSheetName { get; set; }

        [Category("Destination")]
        [Description("Pass the Destination Range, eg: C1")]
        [RequiredArgument]
        public InArgument<string> In_DstRange { get; set; }

        [Category("Property")]
        [Description("Pass the Excel Paste Type, eg: All")]
        [RequiredArgument]
        public InArgument<XlPasteType> In_PasteType { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.CopyPasteRange(In_SrcWorkbookPath.Get(context),
                                            In_SrcSheetName.Get(context),
                                            In_SrcRange.Get(context),
                                            In_DstWorkbookPath.Get(context),
                                            In_DstSheetName.Get(context),
                                            In_DstRange.Get(context),
                                            In_PasteType.Get(context));
        }
    }
}