using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal static class ColourScaleReadHandler
    {
        //We have no way of knowing what type of colorScale it is until we've read its first 3 nodes and xr is forward only.
        //This way we can determine type after reading the initial data.
        internal static ExcelConditionalFormattingRule CreateScales(ExcelAddress address, XmlReader xr, ExcelWorksheet ws)
        {
            //Read base rules
            ExcelConditionalFormattingRule rule = new ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType.ThreeColorScale, address, ws, xr);

            xr.Read();
            var lowType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();

            xr.Read();
            var middleOrHigh = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();

            xr.Read();

            if (xr.LocalName == "color")
            {
                var twoColor = new ExcelConditionalFormattingTwoColorScale(rule, lowType, middleOrHigh, xr);
                //twoColor.SetValues(lowType, middleOrHigh);
                //twoColor.ReadColors(xr);

                twoColor.Type = eExcelConditionalFormattingRuleType.TwoColorScale;

                return twoColor;
            }

            var highType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();

            var threeColor = new ExcelConditionalFormattingThreeColorScale(rule, lowType, middleOrHigh, highType, xr);

            //threeColor.SetValues(lowType, middleOrHigh, highType);
            //threeColor.ReadColors(xr);

            return threeColor;

            //LowValue = new ExcelConditionalFormattingColorScaleValue(
            //    xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>() ?? 0,
            //    ExcelConditionalFormattingConstants.Colors.CfvoLowValue,
            //    Priority);

            //xr.Read();
            //HighValue = new ExcelConditionalFormattingColorScaleValue(
            //    xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>() ?? 0,
            //    ExcelConditionalFormattingConstants.Colors.CfvoHighValue,
            //    Priority);

        }
    }
}
