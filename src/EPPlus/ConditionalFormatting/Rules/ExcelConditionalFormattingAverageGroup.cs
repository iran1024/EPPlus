using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingAverageGroup : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingAverageGroup
    {
        internal ExcelConditionalFormattingAverageGroup(
         ExcelAddress address,
         eExcelConditionalFormattingRuleType type,
         int priority,
         ExcelWorksheet worksheet)
         : base(type, address, priority, worksheet)
        {
        }

        internal ExcelConditionalFormattingAverageGroup(
          ExcelAddress address, eExcelConditionalFormattingRuleType type, ExcelWorksheet ws, XmlReader xr)
          : base(address, type, ws, xr)
        {
        }
    }
}
