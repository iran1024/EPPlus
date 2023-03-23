using OfficeOpenXml.ConditionalFormatting.Contracts;

using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class ExcelConditionalFormattingCollection : IEnumerable<ExcelConditionalFormattingRule>
    {
        List<ExcelConditionalFormattingRule> _rules = new List<ExcelConditionalFormattingRule>();
        ExcelWorksheet _ws;
        int LastPriority = 1;

        internal ExcelConditionalFormattingCollection(ExcelWorksheet ws)
        {
            _ws = ws;
            _rules = new List<ExcelConditionalFormattingRule>();
        }

        internal ExcelConditionalFormattingCollection(XmlReader xr, ExcelWorksheet ws)
        {
            _ws = ws;

            while (xr.ReadUntil(1, "conditionalFormatting", "sheetData", "dataValidations", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks", "extLst", "pageMargins"))
            {
                if(xr.LocalName == "conditionalFormatting")
                {
                    string address = xr.GetAttribute("sqref");

                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        xr.Read();

                        var cf = ExcelConditionalFormattingRuleFactory.Create(new ExcelAddress(address), _ws, xr);

                        _rules.Add(cf);
                    }

                    xr.Read();
                    xr.Read();
                }
            }
        }

        internal void CopyRule(ExcelConditionalFormattingRule rule, ExcelAddress address = null)
        {
            var ruleCopy = rule.Clone();
            if(address != null)
            {
                ruleCopy.Address = address;
            }
            _rules.Add(ruleCopy);
        }

        internal void DeepCopyWithNewAddress(ExcelAddress address, ExcelConditionalFormattingRule orginalRule)
        {
            _rules.Add(orginalRule.Clone());
        }

        //internal void AddFromXml(ExcelAddress address, bool pivot, string ruleXml)
        //{
        //    var cfRuleNode = (XmlElement)CreateNode(ExcelConditionalFormattingConstants.Paths.ConditionalFormatting, false, true);
        //    cfRuleNode.SetAttribute("sqref", address.AddressSpaceSeparated);
        //    cfRuleNode.InnerXml = ruleXml;
        //    var rule = AddNewCf(address, cfRuleNode.FirstChild);
        //    rule.PivotTable = pivot;
        //}


        IEnumerator<ExcelConditionalFormattingRule> IEnumerable<ExcelConditionalFormattingRule>.GetEnumerator()
        {
            return _rules.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _rules.GetEnumerator();
        }

        /// <summary>
        /// Index operator, returns by 0-based index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelConditionalFormattingRule this[int index]
        {
            get { return _rules[index]; }
            set { _rules[index] = value; }
        }

        /// <summary>
        /// Number of validations
        /// </summary>
        public int Count
        {
            get { return _rules.Count; }
        }

        /// <summary>
        /// Removes all 'cfRule' from the collection and from the XML.
        /// <remarks>
        /// This is the same as removing all the 'conditionalFormatting' nodes.
        /// </remarks>
        /// </summary>
        public void RemoveAll()
        {
            // Clear the <cfRule> item list
            _rules.Clear();
        }

        /// <summary>
        /// Remove a Conditional Formatting Rule by its object
        /// </summary>
        /// <param name="item"></param>
        public void Remove(
          IExcelConditionalFormattingRule item)
        {
            Require.Argument(item).IsNotNull("item");

            try
            {
                _rules.Remove((ExcelConditionalFormattingRule)item);
            }
            catch
            {
                throw new Exception($"Cannot remove {item} as it is not part of this collection.");
            }
        }

        /// <summary>
        /// Remove a Conditional Formatting Rule by its 0-based index
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(
          int index)
        {
            Require.Argument(index).IsInRange(0, this.Count - 1, "index");

            Remove(this[index]);
        }

        /// <summary>
        /// Remove a Conditional Formatting Rule by its priority
        /// </summary>
        /// <param name="priority"></param>
        public void RemoveByPriority(
          int priority)
        {
            try
            {
                Remove(RulesByPriority(priority));
            }
            catch
            {
            }
        }

        /// <summary>
        /// Get a rule by its priority
        /// </summary>
        /// <param name="priority"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingRule RulesByPriority(
          int priority)
        {
            return _rules.Find(x => x.Priority == priority);
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
            ExcelConditionalFormattingRule cfRule = ExcelConditionalFormattingRuleFactory.Create(
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

        public IExcelConditionalFormattingLessThan AddLessThan(
            ExcelAddress address)
        {
            return (IExcelConditionalFormattingLessThan)AddRule(
              eExcelConditionalFormattingRuleType.LessThan,
              address);
        }

        public IExcelConditionalFormattingBetween AddBetween(
            ExcelAddress address)
        {
            return (IExcelConditionalFormattingBetween)AddRule(
              eExcelConditionalFormattingRuleType.Between,
              address);
        }

        public IExcelConditionalFormattingEqual AddEqual(ExcelAddress address)
        {
            return (IExcelConditionalFormattingEqual)AddRule(
              eExcelConditionalFormattingRuleType.Equal,
              address);
        }

        public IExcelConditionalFormattingContainsText AddTextContains(ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsText)AddRule(
              eExcelConditionalFormattingRuleType.ContainsText,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddYesterday(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Yesterday,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddToday(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Today,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddTomorrow(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Tomorrow,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddLast7Days(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.Last7Days,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddLastWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.LastWeek,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddThisWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.ThisWeek,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddNextWeek(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.NextWeek,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddLastMonth(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.LastMonth,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddThisMonth(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.ThisMonth,
              address);
        }

        public IExcelConditionalFormattingTimePeriodGroup AddNextMonth(ExcelAddress address)
        {
            return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
              eExcelConditionalFormattingRuleType.NextMonth,
              address);
        }

        /// <summary>
        /// Add DuplicateValues Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingDuplicateValues AddDuplicateValues(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingDuplicateValues)AddRule(
              eExcelConditionalFormattingRuleType.DuplicateValues,
              address);
        }

        /// <summary>
        /// Add Bottom Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottom(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.Bottom,
              address);
        }

        /// <summary>
        /// Add BottomPercent Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottomPercent(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.BottomPercent,
              address);
        }

        /// <summary>
        /// Add Top Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTop(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.Top,
              address);
        }

        /// <summary>
        /// Add TopPercent Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTopPercent(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTopBottomGroup)AddRule(
              eExcelConditionalFormattingRuleType.TopPercent,
              address);
        }

        /// <summary>
        /// Add AboveAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveAverage,
              address);  
        }

        /// <summary>
        /// Add AboveOrEqualAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveOrEqualAverage,
              address);
        }

        /// <summary>
        /// Add BelowAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowAverage,
              address);
        }

        /// <summary>
        /// Add BelowOrEqualAverage Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingAverageGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowOrEqualAverage,
              address);
        }

        /// <summary>
        /// Add AboveStdDev Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddAboveStdDev(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingStdDevGroup)AddRule(
              eExcelConditionalFormattingRuleType.AboveStdDev,
              address);
        }

        /// <summary>
        /// Add BelowStdDev Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddBelowStdDev(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingStdDevGroup)AddRule(
              eExcelConditionalFormattingRuleType.BelowStdDev,
              address);
        }

        //Beyond this line are temporary placeholders of unimplemented methods.
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


        /// <summary>
        /// Add BeginsWith Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingBeginsWith AddBeginsWith(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingBeginsWith)AddRule(
              eExcelConditionalFormattingRuleType.BeginsWith,
              address);
        }

        /// <summary>
        /// Add ContainsBlanks Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsBlanks AddContainsBlanks(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsBlanks)AddRule(
              eExcelConditionalFormattingRuleType.ContainsBlanks,
              address);
        }

        /// <summary>
        /// Add ContainsErrors Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsErrors AddContainsErrors(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsErrors)AddRule(
              eExcelConditionalFormattingRuleType.ContainsErrors,
              address);
        }

        /// <summary>
        /// Add ContainsText Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingContainsText AddContainsText(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingContainsText)AddRule(
              eExcelConditionalFormattingRuleType.ContainsText,
              address);
        }

        /// <summary>
        /// Add EndsWith Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingEndsWith AddEndsWith(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingEndsWith)AddRule(
              eExcelConditionalFormattingRuleType.EndsWith,
              address);
        }

        /// <summary>
        /// Add Expression Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingExpression AddExpression(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingExpression)AddRule(
              eExcelConditionalFormattingRuleType.Expression,
              address);
        }

        /// <summary>
        /// Add GreaterThanOrEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingGreaterThanOrEqual)AddRule(
              eExcelConditionalFormattingRuleType.GreaterThanOrEqual,
              address);
        }

        /// <summary>
        /// Add LessThanOrEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingLessThanOrEqual)AddRule(
              eExcelConditionalFormattingRuleType.LessThanOrEqual,
              address);
        }

        /// <summary>
        /// Add NotBetween Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotBetween AddNotBetween(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotBetween)AddRule(
              eExcelConditionalFormattingRuleType.NotBetween,
              address);
        }

        /// <summary>
        /// Add NotContainsBlanks Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotContainsBlanks)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsBlanks,
              address);
        }

        /// <summary>
        /// Add NotContainsErrors Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotContainsErrors)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsErrors,
              address);
        }

        /// <summary>
        /// Add NotContainsText Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotContainsText AddNotContainsText(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotContainsText)AddRule(
              eExcelConditionalFormattingRuleType.NotContainsText,
              address);
        }

        /// <summary>
        /// Add NotEqual Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingNotEqual AddNotEqual(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingNotEqual)AddRule(
              eExcelConditionalFormattingRuleType.NotEqual,
              address);
        }

        /// <summary>
        /// Add Unique Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingUniqueValues AddUniqueValues(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingUniqueValues)AddRule(
              eExcelConditionalFormattingRuleType.UniqueValues,
              address);
        }

        /// <summary>
        /// Add ThreeColorScale Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingThreeColorScale AddThreeColorScale(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingThreeColorScale)AddRule(
              eExcelConditionalFormattingRuleType.ThreeColorScale,
              address);
        }

        /// <summary>
        /// Add TwoColorScale Rule
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingTwoColorScale AddTwoColorScale(
          ExcelAddress address)
        {
            return (IExcelConditionalFormattingTwoColorScale)AddRule(
              eExcelConditionalFormattingRuleType.TwoColorScale,
              address);
        }

        /// <summary>
        /// Add ThreeIconSet Rule
        /// </summary>
        /// <param name="Address">The address</param>
        /// <param name="IconSet">Type of iconset</param>
        /// <returns></returns>
        public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(ExcelAddress Address, eExcelconditionalFormatting3IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)AddRule(
                eExcelConditionalFormattingRuleType.ThreeIconSet,
                Address);
            icon.IconSet = IconSet;
            return icon;
        }
        /// <summary>
        /// Adds a FourIconSet rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(ExcelAddress Address, eExcelconditionalFormatting4IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)AddRule(
                eExcelConditionalFormattingRuleType.FourIconSet,
                Address);
            icon.IconSet = IconSet;
            return icon;
        }
        /// <summary>
        /// Adds a FiveIconSet rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(ExcelAddress Address, eExcelconditionalFormatting5IconsSetType IconSet)
        {
            var icon = (IExcelConditionalFormattingFiveIconSet)AddRule(
                eExcelConditionalFormattingRuleType.FiveIconSet,
                Address);
            icon.IconSet = IconSet;
            return icon;
        }
        /// <summary>
        /// Adds a databar rule
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingDataBarGroup AddDatabar(ExcelAddress Address, Color color)
        {
            var dataBar = (IExcelConditionalFormattingDataBarGroup)AddRule(
                eExcelConditionalFormattingRuleType.DataBar,
                Address);
            dataBar.Color = color;
            return dataBar;
        }
    }
}
