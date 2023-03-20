using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    public class ExcelConditionalFormattingDuplicateValues : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingDuplicateValues
    {
        public ExcelConditionalFormattingDuplicateValues(
         ExcelAddress address,
         int priority,
         ExcelWorksheet worksheet)
         : base(eExcelConditionalFormattingRuleType.DuplicateValues, address, priority, worksheet)
        {

        }

        public ExcelConditionalFormattingDuplicateValues(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(address, eExcelConditionalFormattingRuleType.DuplicateValues, ws, xr)
        {
        }
    }
}
