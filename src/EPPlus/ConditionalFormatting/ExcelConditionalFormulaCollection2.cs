using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingCollection2 : IEnumerable<Rules2.ExcelConditionalFormattingRule>
    {
        List<Rules2.ExcelConditionalFormattingRule> _rules;
        ExcelWorksheet _ws;
        int LastPriority = 1;

        internal ExcelConditionalFormattingCollection2(ExcelWorksheet ws)
        {
            _ws = ws;
            _rules = new List<Rules2.ExcelConditionalFormattingRule>();
        }

        internal ExcelConditionalFormattingCollection2(XmlReader xr)
        {
            while (xr.LocalName == "conditionalFormatting")
            {
                string address = xr.GetAttribute("sqref");

                if (xr.NodeType == XmlNodeType.Element)
                {
                    xr.Read();
                    var cf = Rules2.ExcelConditionalFormattingRuleFactory.Create(address, xr);

                    _rules.Add(cf);
                    //_validations.Add(validation);
                    //_validationsRD.Add(validation.Address._fromRow, validation.Address._fromCol,
                    //                   validation.Address._toRow, validation.Address._toCol, validation);
                }

                xr.Read();
                xr.Read();
            }

            ////string address = xr.GetAttribute("sqref");

            //if (xr.LocalName != "cfRule")
            //{
            //    xr.Read();
            //    xr.Read(); //Read beyond the end element
            //    ////Read beyond the end element
            //    //break;
            //}

            //if (xr.NodeType == XmlNodeType.Element)
            //{
            //    var cf = Rules2.ExcelConditionalFormattingRuleFactory.Create(address, xr);
            //    //_validations.Add(validation);
            //    //_validationsRD.Add(validation.Address._fromRow, validation.Address._fromCol,
            //    //                   validation.Address._toRow, validation.Address._toCol, validation);
            //}

            //while (xr.Read())
            //{
            //    if (xr.LocalName != "cfRule")
            //    {
            //        xr.Read(); //Read beyond the end element
            //        break;
            //    }

            //    if (xr.NodeType == XmlNodeType.Element)
            //    {
            //        var cf = Rules2.ExcelConditionalFormattingRuleFactory.Create(address, xr);
            //        //_validations.Add(validation);
            //        //_validationsRD.Add(validation.Address._fromRow, validation.Address._fromCol,
            //        //                   validation.Address._toRow, validation.Address._toCol, validation);
            //    }
            //}
        }



        IEnumerator<Rules2.ExcelConditionalFormattingRule> IEnumerable<Rules2.ExcelConditionalFormattingRule>.GetEnumerator()
        {
            return _rules.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _rules.GetEnumerator();
        }

        /// <summary>
        /// Add rule (internal)
        /// </summary>
        /// <param name="type"></param>
        /// <param name="address"></param>
        /// <returns></returns>F
        internal IExcelConditionalFormattingRule AddRule(
          eExcelConditionalFormattingRuleType type,
          ExcelAddress address)
        {
            Require.Argument(address).IsNotNull("address");

            // address = ValidateAddress(address);

            // Create the Rule according to the correct type, address and priority
            Rules2.ExcelConditionalFormattingRule cfRule = Rules2.ExcelConditionalFormattingRuleFactory.Create(
              type,
              address,
              LastPriority++,
              _ws);

            // Add the newly created rule to the list
            _rules.Add(cfRule);

            // Return the newly created rule
            return cfRule;
        }

        /// <summary>
        /// Add GreaterThan Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThan AddGreaterThan(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingGreaterThan)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThan,
              address);
        }
    }
}
