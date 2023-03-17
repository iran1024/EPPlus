/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting.Rules2
{
  /// <summary>
  /// ExcelConditionalFormattingTimePeriodGroup
  /// </summary>
  public class ExcelConditionalFormattingTimePeriodGroup: ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTimePeriodGroup
  {
    /****************************************************************************************/

        #region Constructors
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="priority"></param>
        /// <param name="address"></param>
        /// <param name="worksheet"></param>
        internal ExcelConditionalFormattingTimePeriodGroup(ExcelAddress address,
            eExcelConditionalFormattingRuleType type,
            int priority,
            ExcelWorksheet worksheet)
            : base(type, address, priority, worksheet)
        {
        }

        internal ExcelConditionalFormattingTimePeriodGroup(ExcelAddress address, 
            eExcelConditionalFormattingRuleType type, ExcelWorksheet ws, XmlReader xr)
            : base(address, type, ws, xr)
        {
            string timePeriodString = xr.GetAttribute("timePeriod").ToString().CapitalizeFirstLetter();

            TimePeriod = timePeriodString.ConvertToEnum<eExcelConditionalFormattingTimePeriodType>();

            //var typeTest = eExcelConditionalFormattingTimePeriodType.Today;

            //TimePeriod = typeTest.ParseEnum<eExcelConditionalFormattingTimePeriodType>(timePeriodString);

            ////Enum.ParseEnum()
            ////eExcelConditionalFormattingTimePeriodType.Parse()

            //var eType =
            //   (eExcelConditionalFormattingTimePeriodType)
            //    Enum.Parse(typeof(eExcelConditionalFormattingTimePeriodType),
            //    timePeriodString);

            //TimePeriod = eType;
        }


    #endregion Constructors

        /****************************************************************************************/
  }
}