using System.Activities;
using System.ComponentModel;

namespace IPA_Excel_Extension
{
    public class DisplayAlertsONOFF : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Pass the boolean")]
        public InArgument<bool> ONOFF { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
          //  ExcelExtension.DisplayAlertsONOFF(ONOFF.Get(context));
        }
    }
}