using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// ExcelConditionalFormattingLast7Days
    /// </summary>
    public class ExcelConditionalFormattingToday: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingToday(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(address, eExcelConditionalFormattingRuleType.Today, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.Today;
            Formula = string.Format(
              "FLOOR({0},1)=TODAY()",
              Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingToday(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(address, eExcelConditionalFormattingRuleType.Today, ws, xr)
        {
        }
        #endregion
    }
}