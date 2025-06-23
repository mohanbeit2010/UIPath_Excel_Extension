using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class vlookup : CodeActivity
    {
        [Category("Source")]
        [RequiredArgument]
        [Description("Pass the Source Excel Workbook Path With Extension. eg., test.xlsx")]
        public InArgument<string> In_Str_SourceExcelWorkbookPath { get; set; }

        [Category("Source")]
        [RequiredArgument]
        [Description("Pass the Source Excel WorkSheet Name As String. eg: Sheet1")]
        public InArgument<string> In_Str_SourceSheetName { get; set; }

        [Category("Source")]
        [RequiredArgument]
        [Description("Specify the Source Key Column Name eg., 'Associate ID' as string")]
        public InArgument<string> In_Str_Source_KeyColumnName { get; set; }

        [Category("Source")]
        [RequiredArgument]
        [Description("Specify the Column Name eg., 'Associate Name' as string from which the values must be looked up.")]
        public InArgument<string> In_Str_Source_LookupColumnName { get; set; }

        [Category("Source")]
        [Description("Specify the row in which the column header is present in Source sheet as int.")]
        public InArgument<int> In_Int_Source_HeaderRow { get; set; } = 1;

        [Category("Destination")]
        [RequiredArgument]
        [Description("Pass the Destination Excel Workbook Path With Extension. eg., test.xlsx")]
        public InArgument<string> In_Str_DestinationExcelWorkbookPath { get; set; }
        public InArgument<string> In_Str_DestinationSheetName { get; set; }

        [Category("Destination")]
        [RequiredArgument]
        [Description("Specify the Destination Key Column Name eg., 'Associate ID' as string")]
        public InArgument<string> In_Str_DestinationKeyColumnName { get; set; }

        [Category("Destination")]
        [RequiredArgument]
        [Description("Specify the Column Name eg., 'Associate Name' as string from which the values must be looked up.")]
        public InArgument<string> In_Str_DestinationLookupColumnName { get; set; }

        [Category("Destination")]
        [Description("Specify the row in which the column header is present in destination sheet as int.")]
        public InArgument<int> In_Int_Destination_headerRow { get; set; } = 1;

        protected override void Execute(CodeActivityContext context)
        {
            //ExcelExtension.VLookupRange(
            //    In_Str_SourceExcelWorkbookPath.Get(context),
            //    In_Str_DestinationExcelWorkbookPath.Get(context),
            //    In_Str_SourceSheetName.Get(context),
            //    In_Str_DestinationSheetName.Get(context),
            //    In_Str_Source_KeyColumnName.Get(context),
            //    In_Str_DestinationKeyColumnName.Get(context),
            //    In_Str_Source_LookupColumnName.Get(context),
            //    In_Str_DestinationLookupColumnName.Get(context),
            //    In_Int_Source_HeaderRow.Get(context),
            //    In_Int_Destination_headerRow.Get(context)
            //);
        }
    }
}
