using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    public class ExcelConditionalFormattingTop : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTopBottomGroup
    {
        internal ExcelConditionalFormattingTop(
         ExcelAddress address,
         int priority,
         ExcelWorksheet worksheet)
         : base(eExcelConditionalFormattingRuleType.Top, address, priority, worksheet)
        {
            Bottom = false;
            Percent = false;
            Rank = 10;  // First 10 values
        }

        internal ExcelConditionalFormattingTop(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(address, eExcelConditionalFormattingRuleType.Top, ws, xr)
        {
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);

            Rank = ushort.Parse(xr.GetAttribute("rank"));

            Bottom = !string.IsNullOrEmpty(xr.GetAttribute("bottom"));
            Percent = !string.IsNullOrEmpty(xr.GetAttribute("percent"));
        }

    }
}
