using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class FilterRange : CodeActivity
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
        [RequiredArgument]
        [Description("Pass the ColumnName & ColumnValues that need to be filtered, eg: {\"EmployeeName\",{\"Kim\",\"John\",\"Mike\"}}")]
        public InArgument<Dictionary<string, string[]>> In_Dict_ColumnNamesWithValues { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the ColumnHeaderRow Number in which we find the LastRow, eg: \"1\"")]
        public InArgument<int> In_Int_ColumnHeaderRow { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the bool value to Remove Autofilter in sheet eg: \"true\"")]
        public InArgument<bool> In_Bool_DisableExistingFilter { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            ExcelExtension.FilterRange(
                In_Str_ExcelWorkbookPath.Get(context),
                In_Str_SheetName.Get(context),
                In_Dict_ColumnNamesWithValues.Get(context),
                In_Int_ColumnHeaderRow.Get(context),
                In_Bool_DisableExistingFilter.Get(context)
                );
        }
    }
}