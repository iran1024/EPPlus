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
    public class ExcelConditionalFormattingThreeColorScale : ExcelConditionalFormattingTwoColorScale,
    IExcelConditionalFormattingThreeColorScale
    {

        internal ExcelConditionalFormattingThreeColorScale(ExcelAddress address, int priority, ExcelWorksheet ws)
            : base(address, priority, ws)
        {
            MiddleValue = new ExcelConditionalFormattingColorScaleValue(
            eExcelConditionalFormattingValueObjectType.Percentile,
            ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue,
            priority);

            MiddleValue.Value = 50;
        }

        internal ExcelConditionalFormattingThreeColorScale(
        ExcelConditionalFormattingRule rule,
        eExcelConditionalFormattingValueObjectType? low,
        eExcelConditionalFormattingValueObjectType? middle,
        eExcelConditionalFormattingValueObjectType? high,
        XmlReader xr) : base(rule, low, high, xr)
        {
            MiddleValue = new ExcelConditionalFormattingColorScaleValue(
            middle,
            ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue,
            Priority);
        }

        //internal override void SetValues(eExcelConditionalFormattingValueObjectType? low, eExcelConditionalFormattingValueObjectType? high, eExcelConditionalFormattingValueObjectType? middle)
        //{
        //    base.SetValues(low, high);

        //    MiddleValue = new ExcelConditionalFormattingColorScaleValue(
        //    middle,
        //    ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue,
        //    Priority);
        //}


        internal override void ReadColors(XmlReader xr)
        {
            //we don't call base as the order of nodes are different. Second node is middle.

            Type = eExcelConditionalFormattingRuleType.ThreeColorScale;

            var lowCol = int.Parse(xr.GetAttribute("rgb"), NumberStyles.HexNumber);
            LowValue.Color = Color.FromArgb(lowCol);

            xr.Read();

            var middleCol = int.Parse(xr.GetAttribute("rgb"), NumberStyles.HexNumber);
            MiddleValue.Color = Color.FromArgb(middleCol);

            xr.Read();

            var highCol = int.Parse(xr.GetAttribute("rgb"), NumberStyles.HexNumber);
            MiddleValue.Color = Color.FromArgb(highCol);

            xr.Read();
            xr.Read();
        }

        public ExcelConditionalFormattingColorScaleValue LowValue
        {
            get;
            set;
        }


        public ExcelConditionalFormattingColorScaleValue MiddleValue
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
