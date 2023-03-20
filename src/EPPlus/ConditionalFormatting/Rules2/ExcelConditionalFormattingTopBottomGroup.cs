using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    public class ExcelConditionalFormattingTopBottomGroup : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTopBottomGroup
    {
        internal ExcelConditionalFormattingTopBottomGroup(
         ExcelAddress address,
         eExcelConditionalFormattingRuleType type,
         int priority,
         ExcelWorksheet worksheet)
         : base(type, address, priority, worksheet)
        {
            Rank = 10;  // First 10 values
        }

        internal ExcelConditionalFormattingTopBottomGroup(
          ExcelAddress address, eExcelConditionalFormattingRuleType type, ExcelWorksheet ws, XmlReader xr)
          : base(address, type, ws, xr)
        {
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);

            Rank = UInt16.Parse(xr.GetAttribute("rank"));

        }

    }
}
