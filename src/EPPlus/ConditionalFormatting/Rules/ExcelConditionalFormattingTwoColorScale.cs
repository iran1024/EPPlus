using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.ConditionalFormatting
{
    public class ExcelConditionalFormattingTwoColorScale : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingTwoColorScale
    {
        internal ExcelConditionalFormattingTwoColorScale(
        ExcelAddress address,
        int priority,
        ExcelWorksheet ws) 
        : base(
            eExcelConditionalFormattingRuleType.TwoColorScale, 
            address, 
            priority, 
            ws)
        {
            LowValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Min, 
                ExcelConditionalFormattingConstants.Colors.CfvoLowValue, 
                priority);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
                eExcelConditionalFormattingValueObjectType.Min,
                ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
                priority);
        }

        internal ExcelConditionalFormattingTwoColorScale(ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
            :base(eExcelConditionalFormattingRuleType.TwoColorScale, address, ws, xr)
        {
            xr.Read();
            LowValue = new ExcelConditionalFormattingColorScaleValue(
                xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>() ?? 0,
                ExcelConditionalFormattingConstants.Colors.CfvoLowValue,
                Priority);

            xr.Read();
            HighValue = new ExcelConditionalFormattingColorScaleValue(
                xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>() ?? 0,
                ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
                Priority);
        }

        public ExcelConditionalFormattingColorScaleValue LowValue
        {
            get;
            set;
        }

        /// <summary>
        /// High Value for Two Color Scale Object Value
        /// </summary>
        public ExcelConditionalFormattingColorScaleValue HighValue
        {
            get;
            set;
        }
    }
}
