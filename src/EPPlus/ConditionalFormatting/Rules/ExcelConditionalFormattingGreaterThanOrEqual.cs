using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingGreaterThanOrEqual : ExcelConditionalFormattingRule, IExcelConditionalFormattingGreaterThanOrEqual
    {

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingGreaterThanOrEqual(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(eExcelConditionalFormattingRuleType.GreaterThanOrEqual, address, priority, worksheet)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThanOrEqual;
            Formula = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingGreaterThanOrEqual(ExcelAddress address, ExcelWorksheet ws, XmlReader xr) 
            : base(eExcelConditionalFormattingRuleType.GreaterThanOrEqual, address, ws, xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.GreaterThanOrEqual;
        }

        #endregion Constructors
    }
}
