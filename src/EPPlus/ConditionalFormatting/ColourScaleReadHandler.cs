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
            string lowVal = xr.GetAttribute("val");

            xr.Read();
            var middleOrHigh = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();
            string middleOrHighVal = xr.GetAttribute("val");

            xr.Read();

            if (xr.LocalName == "color")
            {
                var twoColor = new ExcelConditionalFormattingTwoColorScale(
                    rule, lowType, middleOrHigh, lowVal, middleOrHighVal, xr);

                twoColor.Type = eExcelConditionalFormattingRuleType.TwoColorScale;

                return twoColor;
            }

            var highType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();
            string highVal = xr.GetAttribute("val");

            xr.Read();

            var threeColor = new ExcelConditionalFormattingThreeColorScale(
                rule, lowType, middleOrHigh, highType, lowVal, middleOrHighVal, highVal, xr);

            return threeColor;
        }
    }
}
