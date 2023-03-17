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
    public class ExcelConditionalFormattingLast7Days: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingLast7Days(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(address, eExcelConditionalFormattingRuleType.Last7Days, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.Last7Days;
            Formula = string.Format(
            "AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())",
            Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingLast7Days(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(address, eExcelConditionalFormattingRuleType.Last7Days, ws, xr)
        {
        }
        #endregion
    }
}