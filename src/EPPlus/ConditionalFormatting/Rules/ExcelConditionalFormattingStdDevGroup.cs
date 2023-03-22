using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingStdDevGroup
      : ExcelConditionalFormattingRule,
      IExcelConditionalFormattingStdDevGroup
    {
        internal ExcelConditionalFormattingStdDevGroup(
         ExcelAddress address,
         eExcelConditionalFormattingRuleType type,
         int priority,
         ExcelWorksheet worksheet)
         : base(type, address, priority, worksheet)
        {
            StdDev = 1;
        }

        internal ExcelConditionalFormattingStdDevGroup(
          ExcelAddress address, eExcelConditionalFormattingRuleType type, ExcelWorksheet ws, XmlReader xr)
          : base(address, type, ws, xr)
        {
        }

        internal override void ReadClassSpecificXmlNodes(XmlReader xr)
        {
            base.ReadClassSpecificXmlNodes(xr);
            if(string.IsNullOrEmpty(xr.GetAttribute("stdDev")))
            {
                throw new InvalidOperationException($"Could not read stdDev value of ConditionalFormatting {this} of type: {Type} at adress {Address}. " +
                                                    $"XML corrupted or reading faulty");
            }
            StdDev = UInt16.Parse(xr.GetAttribute("stdDev"));
        }
    }
}
