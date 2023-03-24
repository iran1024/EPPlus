/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;

namespace EPPlusTest.ConditionalFormatting
{
    /// <summary>
    /// Test the Conditional Formatting feature
    /// </summary>
    [TestClass]
    public class ConditionalFormattingTests : TestBase
    {
        private static ExcelPackage _pck;
        [ClassInitialize()]
        public static void Init(TestContext testContext)
        {
            _pck = OpenPackage("ConditionalFormatting.xlsx", true);
        }
        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void MyClassCleanup()
        {
            SaveAndCleanup(_pck);
        }

        /// <summary>
        /// 
        /// </summary>
        [TestMethod]
        public void TwoColorScale()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColorScale");
            var cf = ws.ConditionalFormatting.AddTwoColorScale(ws.Cells["A1:A5"]);
            cf.PivotTable = true;
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);
            ws.SetValue(4, 1, 4);
            ws.SetValue(5, 1, 5);
        }
        [TestMethod]
        public void Pivot()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pivot");
            var cf = ws.ConditionalFormatting.AddThreeColorScale(ws.Cells["A1:A5"]);
            cf.PivotTable = false;
        }

        /// <summary>
        /// 
        /// </summary>
        [TestMethod]
        public void TwoBackColor()
        {
            var ws = _pck.Workbook.Worksheets.Add("TwoBackColor");
            IExcelConditionalFormattingEqual condition1 = ws.ConditionalFormatting.AddEqual(ws.Cells["A1"]);
            condition1.StopIfTrue = true;
            condition1.Priority = 1;
            condition1.Formula = "TRUE";
            condition1.Style.Fill.BackgroundColor.Color = Color.Green;
            IExcelConditionalFormattingEqual condition2 = ws.ConditionalFormatting.AddEqual(ws.Cells["A2"]);
            condition2.StopIfTrue = true;
            condition2.Priority = 2;
            condition2.Formula = "FALSE";
            condition2.Style.Fill.BackgroundColor.Color = Color.Red;
        }
        [TestMethod]
        public void Databar()
        {
            var ws = _pck.Workbook.Worksheets.Add("Databar");
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);
            ws.SetValue(4, 1, 4);
            ws.SetValue(5, 1, 5);
        }
        [TestMethod]
        public void DatabarChangingAddressAddsConditionalFormatNodeInSchemaOrder()
        {
            var ws = _pck.Workbook.Worksheets.Add("DatabarAddressing");
            // Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
            ws.HeaderFooter.AlignWithMargins = true;
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
            Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
            cf.Address = new ExcelAddress("C3");
            Assert.AreEqual("sheetData", cf.Node.ParentNode.PreviousSibling.LocalName);
            Assert.AreEqual("headerFooter", cf.Node.ParentNode.NextSibling.LocalName);
        }
        [TestMethod]
        public void IconSet()
        {
            var ws = _pck.Workbook.Worksheets.Add("IconSet");
            var cf = ws.ConditionalFormatting.AddThreeIconSet(ws.Cells["A1:A3"], eExcelconditionalFormatting3IconsSetType.Symbols);
            ws.SetValue(1, 1, 1);
            ws.SetValue(2, 1, 2);
            ws.SetValue(3, 1, 3);

            var cf4 = ws.ConditionalFormatting.AddFourIconSet(ws.Cells["B1:B4"], eExcelconditionalFormatting4IconsSetType.Rating);
            cf4.Icon1.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon1.Formula = "0";
            cf4.Icon2.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon2.Formula = "1/3";
            cf4.Icon3.Type = eExcelConditionalFormattingValueObjectType.Formula;
            cf4.Icon3.Formula = "2/3";
            ws.SetValue(1, 2, 1);
            ws.SetValue(2, 2, 2);
            ws.SetValue(3, 2, 3);
            ws.SetValue(4, 2, 4);

            var cf5 = ws.ConditionalFormatting.AddFiveIconSet(ws.Cells["C1:C5"], eExcelconditionalFormatting5IconsSetType.Quarters);
            cf5.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon1.Value = 1;
            cf5.Icon2.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon2.Value = 2;
            cf5.Icon3.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon3.Value = 3;
            cf5.Icon4.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon4.Value = 4;
            cf5.Icon5.Type = eExcelConditionalFormattingValueObjectType.Num;
            cf5.Icon5.Value = 5;
            cf5.ShowValue = false;
            cf5.Reverse = true;

            ws.SetValue(1, 3, 1);
            ws.SetValue(2, 3, 2);
            ws.SetValue(3, 3, 3);
            ws.SetValue(4, 3, 4);
            ws.SetValue(5, 3, 5);
        }
        [TestMethod]
        public void WriteReadEqual()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Equal");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddEqual();
                cf.Formula = "1";
                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.Equal;
                    Assert.AreEqual("1", cf.Formula);
                }
            }
        }

        [TestMethod]
        public void WriteReadThreeIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.TrafficLights2);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting3IconsSetType.TrafficLights2, cf.IconSet);
                }
            }
        }
        [TestMethod]
        public void WriteReadFourIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FourIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddFourIconSet(eExcelconditionalFormatting4IconsSetType.ArrowsGray);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.FourIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting4IconsSetType.ArrowsGray, cf.IconSet);
                }
            }
        }
        [TestMethod]
        public void WriteReadFiveIcon()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("FiveIcon");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddFiveIconSet(eExcelconditionalFormatting5IconsSetType.Arrows);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.FiveIconSet;
                    Assert.AreEqual(eExcelconditionalFormatting5IconsSetType.Arrows, cf.IconSet);
                }
            }
        }


        [TestMethod]
        public void WriteReadDataBar()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("DataBar");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddDatabar(Color.Red);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.DataBar;
                    Assert.AreEqual(Color.Red.ToArgb(), cf.Color.ToArgb());
                }
            }
        }

        [TestMethod]
        public void WriteReadTwoColorScale()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("TwoColorScale");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddTwoColorScale();
                cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cf.LowValue.Value = 2;
                cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                cf.HighValue.Value = 50;
                cf.PivotTable = true;

                Assert.AreEqual(2, cf.LowValue.Value);
                Assert.AreEqual(50, cf.HighValue.Value);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.TwoColorScale;
                    Assert.AreEqual(2, cf.LowValue.Value);
                    Assert.AreEqual(50, cf.HighValue.Value);
                }

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void WriteReadThreeColorScale()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ThreeColorScale");
                var cf = ws.Cells["A1"].ConditionalFormatting.AddThreeColorScale();
                cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cf.LowValue.Value = 2;
                cf.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                cf.MiddleValue.Value = 25;
                cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
                cf.HighValue.Value = 50;
                cf.PivotTable = true;

                Assert.AreEqual(2, cf.LowValue.Value);
                Assert.AreEqual(50, cf.HighValue.Value);

                p.Save();
                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p2.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.ThreeColorScale;
                    Assert.AreEqual(2, cf.LowValue.Value);
                    Assert.AreEqual(50, cf.HighValue.Value);
                }

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void VerifyReadStyling()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddBetween(ws.Cells["A1:A3"]);
                cf.Formula = "1";
                cf.Formula2 = "2";

                string expectedFormat = "#,##0";
                cf.Style.Font.Bold = true;
                cf.Style.Font.Italic = true;
                cf.Style.Font.Color.SetColor(Color.Red);
                cf.Style.NumberFormat.Format = expectedFormat;

                p.Save();

                using (var p2 = new ExcelPackage(p.Stream))
                {
                    ws = p.Workbook.Worksheets[0];
                    cf = ws.ConditionalFormatting[0].As.Between;
                    Assert.IsTrue(cf.Style.Font.Bold.Value);
                    Assert.IsTrue(cf.Style.Font.Italic.Value);
                    Assert.AreEqual(Color.Red.ToArgb(), cf.Style.Font.Color.Color.Value.ToArgb());
                    Assert.AreEqual(expectedFormat, cf.Style.NumberFormat.Format);
                }
            }
        }
        [TestMethod]
        public void VerifyExpression()
        {
            using (var p = OpenPackage("cf.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cf = ws.ConditionalFormatting.AddExpression(new ExcelAddress("$1:$1048576"));
                cf.Formula = "IsError(A1)";
                cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cf.Style.Fill.BackgroundColor.SetColor(Color.Red);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void TestInsertRowsIntoVeryLongRangeWithConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the whole of column A except row 1
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A2:A1048576";
                var cf = wks.ConditionalFormatting.AddExpression(new ExcelAddress(cfAddress));
                cf.Formula = "=($A$1=TRUE)";

                // Check that the conditional formatting address was set correctly
                Assert.AreEqual(cfAddress, cf.Address.Address);

                // Insert some rows into the worksheet
                wks.InsertRow(5, 3);

                // Check that the conditional formatting rule still applies to the same range (since there's nowhere to extend it to)
                Assert.AreEqual(cfAddress, cf.Address.Address);
            }
        }
        [TestMethod]
        public void TestInsertRowsAboveVeryLongRangeWithConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the whole of column A except rows 1-10
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A11:A1048576";
                var cf = wks.ConditionalFormatting.AddExpression(new ExcelAddress(cfAddress));
                cf.Formula = "=($A$1=TRUE)";

                // Check that the conditional formatting address was set correctly
                Assert.AreEqual(cfAddress, cf.Address.Address);

                // Insert 3 rows into the worksheet above the conditional formatting
                wks.InsertRow(5, 3);

                // Check that the conditional formatting rule starts lower down, but ends in the same place
                Assert.AreEqual("A14:A1048576", cf.Address.Address);
            }
        }

        [TestMethod]
        public void TestInsertRowsToPushConditionalFormattingOffSheet()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the last two rows of column A
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A1048575:A1048576";
                var cf = wks.ConditionalFormatting.AddExpression(new ExcelAddress(cfAddress));
                cf.Formula = "=($A$1=TRUE)";

                // Check that the conditional formatting address was set correctly
                Assert.AreEqual(1, wks.ConditionalFormatting.Count);
                Assert.AreEqual(cfAddress, cf.Address.Address);

                // Insert enough rows into the worksheet above the conditional formatting rule to push it off the sheet 
                wks.InsertRow(5, 10);

                // Check that the conditional formatting rule no longer exists
                Assert.AreEqual(0, wks.ConditionalFormatting.Count);
            }
        }

        [TestMethod]
        public void TestNewConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                // Add a sheet with conditional formatting on the last two rows of column A
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                var cfAddress = "A1:A10";

                for (int i = 1; i < 11; i++)
                {
                    wks.Cells[i, 1].Value = i;
                }

                var cf = wks.ConditionalFormatting.AddGreaterThan(new ExcelAddress(cfAddress));
                cf.Formula = "5.5";
                cf.Style.Fill.BackgroundColor.SetColor(Color.Red);
                cf.Style.Font.Color.SetColor(Color.White);

                pck.SaveAs("C:\\epplusTest\\Workbooks\\conditionalTest.xlsx");
            }
        }

        static string[] numbers = new string[] 
        { "zero", 
          "one", 
          "two",
          "three", 
          "four",
          "five",
          "six", 
          "seven", 
          "eight", 
          "nine", 
          "ten",
          "eleven" 
        };

        [TestMethod]
        public void TestReadingConditionalFormatting()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                string date = "2023-03-";

                string lastMonth = "2023-02-";
                string thisMonth = "2023-03-";
                string nextMonth = "2023-04-";

                for (int i = 1; i < 11; i++)
                {
                    wks.Cells[i, 5].Value = i;
                    wks.Cells[i, 6].Value = i;
                    wks.Cells[i, 8].Value = i % 2;
                    wks.Cells[i, 10].Value = numbers[i];

                    wks.Cells[i, 12].Value = date + $"{i+10}";
                    wks.Cells[i + 7, 12].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 13].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 13].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 14].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 14].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 15].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 15].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 16].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 16].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 16].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 16].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 17].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 17].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i, 18].Value = date + $"{i + 10}";
                    wks.Cells[i + 7, 18].Value = date + $"{i + 10 + 7}";

                    wks.Cells[i,19].Value = lastMonth + $"{i + 10}";
                    wks.Cells[i + 7, 19].Value = thisMonth + $"{i + 10}";
                    wks.Cells[i + 14, 19].Value = nextMonth + $"{i + 10}";

                    wks.Cells[i, 20].Value = lastMonth + $"{i + 10}";
                    wks.Cells[i + 7, 20].Value = thisMonth + $"{i + 10}";
                    wks.Cells[i + 14, 20].Value = nextMonth + $"{i + 10}";

                    wks.Cells[i, 21].Value = lastMonth + $"{i + 10}";
                    wks.Cells[i + 7, 21].Value = thisMonth + $"{i + 10}";
                    wks.Cells[i + 14, 21].Value = nextMonth + $"{i + 10}";

                    int counter = 0;
                    wks.Cells[i, 23].Value = i % 2 == 1 ? i : counter++ % 2 ;

                    wks.Cells[i, 25].Value = i;
                    wks.Cells[i + 10, 25].Value = i + 10;

                    wks.Cells[i, 26].Value = i;
                    wks.Cells[i + 10, 26].Value = i + 10;

                    wks.Cells[i, 27].Value = i;
                    wks.Cells[i + 10, 27].Value = i + 10;

                    wks.Cells[i, 28].Value = i;
                    wks.Cells[i + 10, 28].Value = i + 10;

                }

                for(int i = 0; i < 4; i++)
                {
                    wks.Cells[1, 30 + i].Value = 3;
                    wks.Cells[2, 30 + i].Value = 2;
                    wks.Cells[3, 30 + i].Value = 4;
                }

                for(int i = 0;i < 2; i++)
                {
                    wks.Cells[1, 35 + i].Value = -19;
                    wks.Cells[2, 35 + i].Value = -10;
                    wks.Cells[3, 35 + i].Value = -1;
                    wks.Cells[4, 35 + i].Value = 0;
                    wks.Cells[5, 35 + i].Value = 1;
                    wks.Cells[6, 35 + i].Value = 9;
                    wks.Cells[7, 35 + i].Value = 17;
                }

                var betweenFormatting = wks.ConditionalFormatting.AddBetween(new ExcelAddress(1, 5, 10, 5));
                betweenFormatting.Formula = "3";
                betweenFormatting.Formula2 = "8";

                betweenFormatting.Style.Fill.BackgroundColor.Color = Color.Red;
                betweenFormatting.Style.Font.Color.Color = Color.Orange;

                var lessFormatting = wks.ConditionalFormatting.AddLessThan(new ExcelAddress(1, 6, 10, 6));
                lessFormatting.Formula = "7";

                lessFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
                lessFormatting.Style.Font.Color.Color = Color.Violet;

                var equalFormatting = wks.ConditionalFormatting.AddEqual(new ExcelAddress(1, 8, 10, 8));
                equalFormatting.Formula = "1";

                equalFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
                equalFormatting.Style.Font.Color.Color = Color.Violet;

                var containsFormatting = wks.ConditionalFormatting.AddTextContains(new ExcelAddress(1, 10, 10, 10));
                containsFormatting.ContainText = "o";

                containsFormatting.Style.Fill.BackgroundColor.Color = Color.Blue;
                containsFormatting.Style.Font.Color.Color = Color.Yellow;

                var dateFormatting = wks.ConditionalFormatting.AddLast7Days(new ExcelAddress(1, 12, 10, 12));

                dateFormatting.Style.Fill.BackgroundColor.Color = Color.Red;
                dateFormatting.Style.Font.Color.Color = Color.Yellow;

                var yesterdayFormatting = wks.ConditionalFormatting.AddYesterday(new ExcelAddress(1, 13, 10, 13));

                //TODO: Fix Priority. It doesn't seem to apply correctly.

                yesterdayFormatting.Style.Fill.BackgroundColor.Color = Color.Gray;
                yesterdayFormatting.Style.Font.Color.Color = Color.Red;
                yesterdayFormatting.Priority = 1;

                var todayFormatting = wks.ConditionalFormatting.AddToday(new ExcelAddress(1, 14, 10, 14));

                todayFormatting.Style.Fill.BackgroundColor.Color = Color.Yellow;
                todayFormatting.Style.Font.Color.Color = Color.Green;
                yesterdayFormatting.Priority = 2;

                var tomorrow = wks.ConditionalFormatting.AddTomorrow(new ExcelAddress(1, 15, 10, 15));

                tomorrow.Style.Fill.BackgroundColor.Color = Color.Black;
                tomorrow.Style.Font.Color.Color = Color.Violet;

                var lastWeek = wks.ConditionalFormatting.AddLastWeek(new ExcelAddress(1, 16, 20, 16));

                lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
                lastWeek.Style.Font.Color.Color = Color.Violet;

                var thisWeek = wks.ConditionalFormatting.AddThisWeek(new ExcelAddress(1, 17, 20, 17));

                thisWeek.Style.Fill.BackgroundColor.Color = Color.Black;
                thisWeek.Style.Font.Color.Color = Color.Violet;

                var nextWeek = wks.ConditionalFormatting.AddNextWeek(new ExcelAddress(1, 18, 20, 18));

                nextWeek.Style.Fill.BackgroundColor.Color = Color.Black;
                nextWeek.Style.Font.Color.Color = Color.Violet;

                var lastMonthCF = wks.ConditionalFormatting.AddLastMonth(new ExcelAddress(1, 19, 27, 19));

                lastMonthCF.Style.Fill.BackgroundColor.Color = Color.Black;
                lastMonthCF.Style.Font.Color.Color = Color.Violet;

                var thisMonthCF = wks.ConditionalFormatting.AddThisMonth(new ExcelAddress(1, 20, 27, 20));

                thisMonthCF.Style.Fill.BackgroundColor.Color = Color.Black;
                thisMonthCF.Style.Font.Color.Color = Color.Violet;

                var nextMonthCF = wks.ConditionalFormatting.AddNextMonth(new ExcelAddress(1, 21, 27, 21));

                nextMonthCF.Style.Fill.BackgroundColor.Color = Color.Black;
                nextMonthCF.Style.Font.Color.Color = Color.Violet;

                var duplicateValues = wks.ConditionalFormatting.AddDuplicateValues(new ExcelAddress(1, 23, 10, 23));

                duplicateValues.Style.Fill.BackgroundColor.Color = Color.Blue;
                duplicateValues.Style.Font.Color.Color = Color.Yellow;


                var top11 = wks.ConditionalFormatting.AddTop(new ExcelAddress(1, 25, 20, 25));

                top11.Rank = 11;
                top11.Style.Fill.BackgroundColor.Color = Color.Black;
                top11.Style.Font.Color.Color = Color.Violet;

                var bot12 = wks.ConditionalFormatting.AddBottom(new ExcelAddress(1, 26, 20, 26));

                bot12.Rank = 12;
                bot12.Style.Fill.BackgroundColor.Color = Color.Black;
                bot12.Style.Font.Color.Color = Color.Violet;

                var top13Percent = wks.ConditionalFormatting.AddTopPercent(new ExcelAddress(1, 27, 20, 27));

                top13Percent.Rank = 13;
                top13Percent.Style.Fill.BackgroundColor.Color = Color.Black;
                top13Percent.Style.Font.Color.Color = Color.Violet;

                var bot14Percent = wks.ConditionalFormatting.AddBottomPercent(new ExcelAddress(1, 28, 20, 28));

                bot14Percent.Rank = 14;
                bot14Percent.Style.Fill.BackgroundColor.Color = Color.Black;
                bot14Percent.Style.Font.Color.Color = Color.Violet;

                var aboveAverage = wks.ConditionalFormatting.AddAboveAverage(new ExcelAddress(1, 30, 20, 30));

                aboveAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                aboveAverage.Style.Font.Color.Color = Color.Violet;

                var aboveOrEqualAverage = wks.ConditionalFormatting.AddAboveOrEqualAverage(new ExcelAddress(1, 31, 20, 31));

                aboveOrEqualAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                aboveOrEqualAverage.Style.Font.Color.Color = Color.Violet;

                var belowAverage = wks.ConditionalFormatting.AddBelowAverage(new ExcelAddress(1, 32, 20, 32));

                belowAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                belowAverage.Style.Font.Color.Color = Color.Violet;

                var belowEqualAverage = wks.ConditionalFormatting.AddBelowOrEqualAverage(new ExcelAddress(1, 33, 20, 33));

                belowEqualAverage.Style.Fill.BackgroundColor.Color = Color.Black;
                belowEqualAverage.Style.Font.Color.Color = Color.Violet;

                var aboveStdDev = wks.ConditionalFormatting.AddAboveStdDev(new ExcelAddress(1, 35, 10, 35));

                aboveStdDev.Style.Fill.BackgroundColor.Color = Color.Black;
                aboveStdDev.Style.Font.Color.Color = Color.Violet;

                aboveStdDev.StdDev = 2;

                var belowStdDev = wks.ConditionalFormatting.AddBelowStdDev(new ExcelAddress(1, 36, 10, 36));

                belowStdDev.Style.Fill.BackgroundColor.Color = Color.Black;
                belowStdDev.Style.Font.Color.Color = Color.Violet;

                belowStdDev.StdDev = 3;

                //var databar = wks.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 38, 10, 38), Color.Crimson);

                pck.SaveAs("C:/epplusTest/Workbooks/conditionalTestEppCopy.xlsx");

                var newPck = new ExcelPackage("C:/epplusTest/Workbooks/conditionalTestEppCopy.xlsx");

                var formattings = newPck.Workbook.Worksheets[0].ConditionalFormatting;

                Assert.AreEqual(formattings.ToList()[0].Formula, "3");
                Assert.AreEqual(formattings.ToList()[0].Formula2, "8");
                Assert.AreEqual(formattings.ToList()[1].Formula, "7");
                Assert.AreEqual(formattings.ToList()[2].Formula, "1");
                Assert.AreEqual(formattings.ToList()[3].Text, "o");

                Assert.AreEqual(formattings.ToList()[4].TimePeriod, eExcelConditionalFormattingTimePeriodType.Last7Days);
                Assert.AreEqual(formattings.ToList()[5].TimePeriod, eExcelConditionalFormattingTimePeriodType.Yesterday);
                Assert.AreEqual(formattings.ToList()[6].TimePeriod, eExcelConditionalFormattingTimePeriodType.Today);
                Assert.AreEqual(formattings.ToList()[7].TimePeriod, eExcelConditionalFormattingTimePeriodType.Tomorrow);
                Assert.AreEqual(formattings.ToList()[8].TimePeriod, eExcelConditionalFormattingTimePeriodType.LastWeek);
                Assert.AreEqual(formattings.ToList()[9].TimePeriod, eExcelConditionalFormattingTimePeriodType.ThisWeek);
                Assert.AreEqual(formattings.ToList()[10].TimePeriod, eExcelConditionalFormattingTimePeriodType.NextWeek);
                Assert.AreEqual(formattings.ToList()[11].TimePeriod, eExcelConditionalFormattingTimePeriodType.LastMonth);
                Assert.AreEqual(formattings.ToList()[12].TimePeriod, eExcelConditionalFormattingTimePeriodType.ThisMonth);
                Assert.AreEqual(formattings.ToList()[13].TimePeriod, eExcelConditionalFormattingTimePeriodType.NextMonth);

                Assert.AreEqual(formattings.ToList()[14].Type, eExcelConditionalFormattingRuleType.DuplicateValues);

                Assert.AreEqual(formattings.ToList()[15].Rank, 11);
                Assert.AreEqual(formattings.ToList()[15].Bottom, false);
                Assert.AreEqual(formattings.ToList()[15].Percent, false);

                Assert.AreEqual(formattings.ToList()[16].Rank, 12);
                Assert.AreEqual(formattings.ToList()[16].Bottom, true);
                Assert.AreEqual(formattings.ToList()[16].Percent, false);

                Assert.AreEqual(formattings.ToList()[17].Bottom, false);
                Assert.AreEqual(formattings.ToList()[17].Percent, true);
                Assert.AreEqual(formattings.ToList()[17].Rank, 13);

                Assert.AreEqual(formattings.ToList()[18].Bottom, true);
                Assert.AreEqual(formattings.ToList()[18].Percent, true);
                Assert.AreEqual(formattings.ToList()[18].Rank, 14);

                Assert.AreEqual(formattings.ToList()[19].AboveAverage, true);
                Assert.AreEqual(formattings.ToList()[19].EqualAverage, false);

                Assert.AreEqual(formattings.ToList()[20].AboveAverage, true);
                Assert.AreEqual(formattings.ToList()[20].EqualAverage, true);

                Assert.AreEqual(formattings.ToList()[21].AboveAverage, false);
                Assert.AreEqual(formattings.ToList()[21].EqualAverage, false);

                Assert.AreEqual(formattings.ToList()[22].AboveAverage, false);
                Assert.AreEqual(formattings.ToList()[22].EqualAverage, true);

                Assert.AreEqual(formattings.ToList()[23].Type, eExcelConditionalFormattingRuleType.AboveStdDev);
                Assert.AreEqual(formattings.ToList()[23].StdDev, 2);

                Assert.AreEqual(formattings.ToList()[24].Type, eExcelConditionalFormattingRuleType.BelowStdDev);
                Assert.AreEqual(formattings.ToList()[24].StdDev, 3);
            }
        }

        [TestMethod]
        public void CFShouldNotThrowIfStyleNotSet()
        {
            //Currently throws bc dxfID. Either give default style or make a better throw.
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                var dateFormatting = wks.ConditionalFormatting.AddLast7Days(new ExcelAddress(1, 12, 10, 12));

                MemoryStream stream = new MemoryStream();
                pck.SaveAs(stream);
                var newPck = new ExcelPackage(stream);
                var formattings = newPck.Workbook.Worksheets[0].ConditionalFormatting;
            }
        }

        [TestMethod]
        public void GreaterThanCanReadWrite()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("GreaterThan");

                for (int i = 1; i < 11; i++)
                {
                    ws.Cells[i, 1].Value = i;
                }

                var greaterThanFormatting = ws.ConditionalFormatting.AddGreaterThan(new ExcelAddress(1, 1, 10, 1));
                greaterThanFormatting.Formula = "3";

                greaterThanFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
                greaterThanFormatting.Style.Font.Color.Color = Color.Violet;

                pck.Save();

                var readPck = new ExcelPackage(pck.Stream);

                foreach (var format in readPck.Workbook.Worksheets[0].ConditionalFormatting)
                {
                    Assert.AreEqual(format.Formula, "3");
                    Assert.AreEqual(Color.Black.ToArgb(),  format.Style.Fill.BackgroundColor.Color.Value.ToArgb());
                    Assert.AreEqual(Color.Violet.ToArgb(), format.Style.Font.Color.Color.Value.ToArgb());
                }
            }
        }
        [TestMethod]
        public void CFWholeSheetRangeDeleteRowShouldNotRemoveCF()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                var cf = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("$1:$1048576"));
                cf.Formula = "Pizza";
                cf.Style.Font.Color.SetColor(Color.Red);

                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
                sheet.DeleteRow(3);
                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
            }
        }


        [TestMethod]
        public void CFColumnsRangeDeleteRowShouldNotRemoveCF()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                var cf = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("$A:$P"));
                cf.Formula = "Pizza";
                cf.Style.Font.Color.SetColor(Color.Red);

                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
                sheet.DeleteRow(3);
                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
            }
        }
        [TestMethod]
        public void CFWholeSheetRange2DeleteRowShouldNotRemoveCF()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("Test");
                var cf = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("A1:XFD1048576"));
                cf.Formula = "Pizza";
                cf.Style.Font.Color.SetColor(Color.Red);

                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
                sheet.DeleteRow(3);
                Assert.AreEqual(1, sheet.ConditionalFormatting.Count);
            }
        }

    }
}