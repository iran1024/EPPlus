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
    public class ExcelConditionalFormattingYesterday: ExcelConditionalFormattingTimePeriodGroup
    {
        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingYesterday(
            ExcelAddress address,
            int priority,
            ExcelWorksheet worksheet)
        : base(address, eExcelConditionalFormattingRuleType.Yesterday, priority, worksheet)
        {
            TimePeriod = eExcelConditionalFormattingTimePeriodType.Yesterday;
            Formula = string.Format(
              "FLOOR({0},1)=TODAY()-1",
              Address.Start.Address);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <param name="ws"></param>
        /// <param name="xr"></param>
        public ExcelConditionalFormattingYesterday(
            ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            : base(address, eExcelConditionalFormattingRuleType.Yesterday, ws, xr)
        {
        }
        #endregion
    }
}