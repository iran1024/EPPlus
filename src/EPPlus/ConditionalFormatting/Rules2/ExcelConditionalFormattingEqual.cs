using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    internal class ExcelConditionalFormattingEqual : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingEqual
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingEqual(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.Equal, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.Equal;
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingEqual(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(address, eExcelConditionalFormattingRuleType.Equal, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.Equal;
        }

        #endregion Constructors
    }
}
