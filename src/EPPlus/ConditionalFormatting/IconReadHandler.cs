using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal static class IconReadHandler
    {
        //We have no way of knowing what type of IconRead it is until we've read its first node and xr is forward only.
        //This way we can determine type after reading the initial data.
        internal static ExcelConditionalFormattingRule ReadIcons(ExcelAddress address, XmlReader xr, ExcelWorksheet ws)
        {
            //Read base rules
            ExcelConditionalFormattingRule rule = new ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType.ThreeIconSet, address, ws, xr);

            //xr.Read();
            var set = xr.GetAttribute("iconSet");

            //The first char of all iconSet types start with number of their type.
            if (set[0] == '3')
            {
                return new ExcelConditionalFormattingThreeIconSet(address, ws, xr, rule);
            }
            else if(set[0] == '4')
            {

            }
            else if(set[0] == '5')
            {

            }

            throw new NotImplementedException();
        }
    }
}
