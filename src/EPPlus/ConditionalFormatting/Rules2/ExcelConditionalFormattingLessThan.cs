using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    internal class ExcelConditionalFormattingLessThan : ExcelConditionalFormattingRule, IExcelConditionalFormattingGreaterThan
    {
        public ExcelConditionalFormattingLessThan(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.LessThan, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.LessThan;
        }
    }
}
