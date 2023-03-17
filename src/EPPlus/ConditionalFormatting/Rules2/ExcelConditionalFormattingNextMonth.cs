using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
    /// <summary>
    /// ExcelConditionalFormattingLast7Days
    /// </summary>
    public class ExcelConditionalFormattingNextMonth: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingNextMonth(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(address, eExcelConditionalFormattingRuleType.NextMonth, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.NextMonth;
            Formula = string.Format(
              "AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)), YEAR({0})=YEAR(EDATE(TODAY(),0+1)))",
              Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingNextMonth(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(address, eExcelConditionalFormattingRuleType.NextMonth, ws, xr)
        {
        }
        #endregion
    }
}