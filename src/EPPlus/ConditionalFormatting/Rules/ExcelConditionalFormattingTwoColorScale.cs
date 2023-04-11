using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
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
                eExcelConditionalFormattingValueObjectType.Max,
                ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
                priority);
        }

        internal ExcelConditionalFormattingTwoColorScale(
            ExcelConditionalFormattingRule rule, 
            eExcelConditionalFormattingValueObjectType? low, 
            eExcelConditionalFormattingValueObjectType? high,
            string lowVal,
            string highVal,
            XmlReader xr) : base (rule)
        {
            SetValues(low, high, lowVal, highVal);
            ReadColors(xr);
        }

        internal void SetValues(
            eExcelConditionalFormattingValueObjectType? low, 
            eExcelConditionalFormattingValueObjectType? high,
            string lowVal = "",
            string highVal = "",
            string middleVal = "",
            eExcelConditionalFormattingValueObjectType? middle = null)
        {
            LowValue = new ExcelConditionalFormattingColorScaleValue(
            low,
            ExcelConditionalFormattingConstants.Colors.CfvoLowValue,
            Priority);

            HighValue = new ExcelConditionalFormattingColorScaleValue(
            high,
            ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
            Priority);

            if(lowVal != "")
            {
                LowValue.Value = double.Parse(lowVal, CultureInfo.InvariantCulture);
            }

            if(highVal!= null) 
            {
                HighValue.Value = double.Parse(highVal, CultureInfo.InvariantCulture);
            }
        }

        internal virtual void ReadColors(XmlReader xr)
        {
            Type = eExcelConditionalFormattingRuleType.TwoColorScale;
            LowValue.Color = ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb"));

            xr.Read();

            HighValue.Color = ExcelConditionalFormattingHelper.ConvertFromColorCode(xr.GetAttribute("rgb"));

            xr.Read();
            xr.Read();
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
