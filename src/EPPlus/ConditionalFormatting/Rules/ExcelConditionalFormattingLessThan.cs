using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingLessThan : ExcelConditionalFormattingRule, IExcelConditionalFormattingLessThan
    {
        public ExcelConditionalFormattingLessThan(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.LessThan, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.LessThan;
        }

        public ExcelConditionalFormattingLessThan(
          ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
          : base(address, eExcelConditionalFormattingRuleType.LessThan, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.LessThan;
        }
    }
}
