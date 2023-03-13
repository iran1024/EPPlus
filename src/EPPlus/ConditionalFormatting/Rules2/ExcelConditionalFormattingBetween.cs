using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    internal class ExcelConditionalFormattingBetween : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingBetween
    {
        /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="priority"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingBetween(
          ExcelAddress address,
          int priority,
          ExcelWorksheet worksheet)
          : base(
                eExcelConditionalFormattingRuleType.Between,
                address,
                priority,
                worksheet
                )
        {
            Operator = eExcelConditionalFormattingOperatorType.Between;
            Formula = string.Empty;
            Formula2 = string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        /// <param name="xr"></param>
        internal ExcelConditionalFormattingBetween(
          ExcelAddress address,
          ExcelWorksheet worksheet,
          XmlReader xr)
          : base(
                address,
                eExcelConditionalFormattingRuleType.Between,
                worksheet,
                xr)
        {
            Operator = eExcelConditionalFormattingOperatorType.Between;
        }
        #endregion Constructors

        /****************************************************************************************/
    }
}
