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
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting;
using System.Xml.Linq;

namespace EPPlusTest.ConditionalFormatting
{
    /// <summary>
    /// Test the Conditional Formatting feature
    /// </summary>
    [TestClass]
    public class ConditionalFormattingTests : TestBase
    {
        private static ExcelPackage _pck;

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

        [ClassInitialize()]
        public static void Init(TestContext testContext)
        {
            _pck = OpenPackage("ConditionalFormatting.xlsx", true);
            var wks = _pck.Workbook.Worksheets.Add("Overview");

            //Filling the overview sheet
            var year = $"{DateTime.Now.Year}";
            string date = $"{DateTime.Now.Year}-{DateTime.Now.Month}-";

            string lastMonth = $"{year}-{DateTime.Now.AddMonths(-1).Month}-";
            string thisMonth = $"{year}-{DateTime.Now.Month}-";
            string nextMonth = $"{year}-{DateTime.Now.AddMonths(+1).Month}-";

            for (int i = 1; i < 11; i++)
            {
                wks.Cells[i, 1].Value = i;
                wks.Cells[i, 2].Value = i;
                wks.Cells[i, 4].Value = i % 2;
                wks.Cells[i, 6].Value = numbers[i];

                wks.Cells[i, 8].Value = date + $"{i + 10}";
                wks.Cells[i + 7, 8].Value = date + $"{i + 10 + 7}";

                wks.Cells[i, 9].Value = date + $"{i + 10}";
                wks.Cells[i + 7, 9].Value = date + $"{i + 10 + 7}";

                wks.Cells[i, 10].Value = date + $"{i + 10}";
                wks.Cells[i + 7, 10].Value = date + $"{i + 10 + 7}";

                wks.Cells[i, 11].Value = date + $"{i + 10}";
                wks.Cells[i + 7, 11].Value = date + $"{i + 10 + 7}";

                wks.Cells[i, 12].Value = date + $"{i + 10}";
                wks.Cells[i + 7, 12].Value = date + $"{i + 10 + 7}";

                wks.Cells[i, 13].Value = date + $"{i + 10}";
                wks.Cells[i + 7, 13].Value = date + $"{i + 10 + 7}";

                wks.Cells[i, 14].Value = date + $"{i + 10}";
                wks.Cells[i + 7, 14].Value = date + $"{i + 10 + 7}";

                wks.Cells[i, 15].Value = lastMonth + $"{i + 10}";
                wks.Cells[i + 7, 15].Value = thisMonth + $"{i + 10}";
                wks.Cells[i + 14, 15].Value = nextMonth + $"{i + 10}";

                wks.Cells[i, 16].Value = lastMonth + $"{i + 10}";
                wks.Cells[i + 7, 16].Value = thisMonth + $"{i + 10}";
                wks.Cells[i + 14, 16].Value = nextMonth + $"{i + 10}";

                wks.Cells[i, 17].Value = lastMonth + $"{i + 10}";
                wks.Cells[i + 7, 17].Value = thisMonth + $"{i + 10}";
                wks.Cells[i + 14, 17].Value = nextMonth + $"{i + 10}";

                int counter = 0;
                wks.Cells[i, 19].Value = i % 2 == 1 ? i : counter++ % 2;

                wks.Cells[i, 21].Value = i;
                wks.Cells[i + 10, 21].Value = i + 10;

                wks.Cells[i, 22].Value = i;
                wks.Cells[i + 10, 22].Value = i + 10;

                wks.Cells[i, 23].Value = i;
                wks.Cells[i + 10, 23].Value = i + 10;

                wks.Cells[i, 24].Value = i;
                wks.Cells[i + 10, 24].Value = i + 10;

                wks.Cells[i, 34].Value = i;

                wks.Cells[i, 36].Value = i;
                wks.Cells[i, 37].Value = i;

                wks.Cells[i, 39].Value = i;
                wks.Cells[i, 40].Value = i;
                wks.Cells[i, 41].Value = i;
                wks.Cells[i, 42].Value = i;
                wks.Cells[i, 43].Value = i;
            }

            for (int i = 0; i < 4; i++)
            {
                wks.Cells[1, 26 + i].Value = 3;
                wks.Cells[2, 26 + i].Value = 2;
                wks.Cells[3, 26 + i].Value = 4;
            }

            for (int i = 0; i < 2; i++)
            {
                wks.Cells[1, 31 + i].Value = -500;
                wks.Cells[2, 31 + i].Value = -10;
                wks.Cells[3, 31 + i].Value = -1;
                wks.Cells[4, 31 + i].Value = 0;
                wks.Cells[5, 31 + i].Value = 1;
                wks.Cells[6, 31 + i].Value = 9;
                wks.Cells[7, 31 + i].Value = 17;
                wks.Cells[8, 31 + i].Value = 25;
                wks.Cells[9, 31 + i].Value = 200;
            }

            wks.Cells["A1:AZ50"].AutoFitColumns();
            wks.Cells["H1:Q30"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            wks.Cells["H1:Q30"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            wks.Cells["H1:Q30"].Style.Border.Right.Color.SetColor(Color.Black);
            wks.Cells["H1:Q30"].Style.Border.Left.Color.SetColor(Color.Black);
        }
        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void CleanUp()
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
        public void DatabarChangingAddressCorrectly()
        {
            var ws = _pck.Workbook.Worksheets.Add("DatabarAddressing");
            // Ensure there is at least one element that always exists below ConditionalFormatting nodes.   
            ws.HeaderFooter.AlignWithMargins = true;
            var cf = ws.ConditionalFormatting.AddDatabar(ws.Cells["A1:A5"], Color.BlueViolet);
            cf.Address = new ExcelAddress("C3");

            Assert.AreEqual(cf.Address, "C3");
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

                //SaveAndCleanup(p);

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
            using (var pck = OpenPackage("conditionalTest.xlsx", true))
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

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void CustomIconsWriteRead()
        {

            using (var pck = OpenPackage("FlagTest.xlsx", true))
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                for (int i = 1; i < 21; i++)
                {
                    wks.Cells[i, 1].Value = i;
                }

                wks.Cells[1, 1, 20, 1].Copy(wks.Cells[1, 2, 20, 2]);
                wks.Cells[1, 1, 20, 1].Copy(wks.Cells[1, 3, 20, 3]);
                wks.Cells[1, 1, 20, 1].Copy(wks.Cells[1, 4, 20, 4]);

                var threeIcon = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1:A20"), eExcelconditionalFormatting3IconsSetType.Triangles);

                threeIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.RedFlag;
                threeIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.NoIcon;
                threeIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow;

                var fourIcon = wks.ConditionalFormatting.AddFourIconSet(new ExcelAddress("B1:B20"), eExcelconditionalFormatting4IconsSetType.Rating);

                fourIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                fourIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                fourIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircleWithBorder;
                fourIcon.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;

                var fiveIcon = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("C1:C20"), eExcelconditionalFormatting5IconsSetType.Boxes);

                fiveIcon.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                fiveIcon.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                fiveIcon.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircleWithBorder;
                fiveIcon.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircle;
                fiveIcon.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.RedCircle;

                var specialCase = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress("D1:D20"), eExcelconditionalFormatting5IconsSetType.Boxes);

                specialCase.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars;
                specialCase.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithOneFilledBar;
                specialCase.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithTwoFilledBars;
                specialCase.Icon4.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars;
                specialCase.Icon5.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithFourFilledBars;

                SaveAndCleanup(pck);

                ExcelPackage package2 = OpenPackage("FlagTest.xlsx");
                var threeIconRead = (ExcelConditionalFormattingThreeIconSet)package2.Workbook.Worksheets[0].ConditionalFormatting[0];

                Assert.AreEqual(threeIconRead.Icon1.CustomIcon, eExcelconditionalFormattingCustomIcon.RedFlag);
                Assert.AreEqual(threeIconRead.Icon2.CustomIcon, eExcelconditionalFormattingCustomIcon.NoIcon);
                Assert.AreEqual(threeIconRead.Icon3.CustomIcon, eExcelconditionalFormattingCustomIcon.GrayDownInclineArrow);

                SaveAndCleanup(package2);
            }
        }

        [TestMethod]
        public void EnsureCustomIconsReturnCorrectStrings()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("FormattingTest");

                var validation = (ExcelConditionalFormattingThreeIconSet)wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress("A1"), eExcelconditionalFormatting3IconsSetType.Triangles);

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowSideArrow;

                Assert.AreEqual("3Arrows", validation.Icon1.GetCustomIconStringValue());

                validation.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.GrayUpArrow;

                Assert.AreEqual("3ArrowsGray", validation.Icon2.GetCustomIconStringValue());

                validation.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowFlag;
                Assert.AreEqual("3Flags", validation.Icon3.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCircle;

                Assert.AreEqual("3TrafficLights1", validation.Icon1.GetCustomIconStringValue());

                validation.Icon2.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowTrafficLight;
                Assert.AreEqual("3TrafficLights2", validation.Icon2.GetCustomIconStringValue());

                validation.Icon3.CustomIcon = eExcelconditionalFormattingCustomIcon.RedDiamond;
                Assert.AreEqual("3Signs", validation.Icon3.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowExclamationSymbol;
                Assert.AreEqual("3Symbols", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.GreenCheck;
                Assert.AreEqual("3Symbols2", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.HalfGoldStar;
                Assert.AreEqual("3Stars", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDash;
                Assert.AreEqual("3Triangles", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.YellowDownInclineArrow;
                Assert.AreEqual("4Arrows", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.PinkCircle;
                Assert.AreEqual("4RedToBlack", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithThreeFilledBars;
                Assert.AreEqual("4Rating", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.BlackCircleWithBorder;
                Assert.AreEqual("4TrafficLights", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.SignalMeterWithNoFilledBars;
                Assert.AreEqual("5Rating", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.CircleWithThreeWhiteQuarters;
                Assert.AreEqual("5Quarters", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.OneFilledBox;
                Assert.AreEqual("5Boxes", validation.Icon1.GetCustomIconStringValue());

                validation.Icon1.CustomIcon = eExcelconditionalFormattingCustomIcon.NoIcon;
                Assert.AreEqual("NoIcons", validation.Icon1.GetCustomIconStringValue());
            }
        }


        [TestMethod]
        public void BeginsWith_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.BeginsWith;

            BaseReadWriteTest("A1:A5", "BeginsWith", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddBeginsWith(address);
                });
        }

        [TestMethod]
        public void EndsWith_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.EndsWith;

            BaseReadWriteTest("A1:A5", "EndsWith", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddEndsWith(address);
                });
        }


        [TestMethod]
        public void Expression_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.Expression;

            BaseReadWriteTest("A1:A5", "Expression", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddExpression(address);
                });
        }


        [TestMethod]
        public void GreaterThanOrEqual_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.GreaterThanOrEqual;

            BaseReadWriteTest("A1:A5", "GreaterThanOrEqual", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddGreaterThanOrEqual(address);
                });
        }

        [TestMethod]
        public void LessThanOrEqual_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.LessThanOrEqual;

            BaseReadWriteTest("A1:A5", "LessThanOrEqual", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddLessThanOrEqual(address);
                });
        }

        [TestMethod]
        public void ContainsBlanks_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.ContainsBlanks;

            BaseReadWriteTest("A1:A5", "ContainsBlanks", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddContainsBlanks(address);
                });
        }

        [TestMethod]
        public void NotBetween_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.NotBetween;

            BaseReadWriteTest("A1:A5", "NotBetween", type,
                (sheet, address) =>
                {
                    var cf = sheet.ConditionalFormatting.AddNotBetween(address);
                    cf.Formula = "1";
                    cf.Formula2 = "5";

                    return (ExcelConditionalFormattingRule)cf;
                });
        }

        [TestMethod]
        public void ContainsErrors_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.ContainsErrors;

            BaseReadWriteTest("A1:A5", "ContainsErrors", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddContainsErrors(address);
                });
        }


        [TestMethod]
        public void Equal_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.Equal;

            BaseReadWriteTest("A1:A5", "Equal", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddEqual(address);
                });
        }

        [TestMethod]
        public void NotEqual_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.NotEqual;

            BaseReadWriteTest("A1:A5", "NotEqual", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddNotEqual(address);
                });
        }

        [TestMethod]
        public void UniqueValues()
        {
            var type = eExcelConditionalFormattingRuleType.UniqueValues;

            BaseReadWriteTest("A1:A5", "UniqueValues", type,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddUniqueValues(address);
                });
        }

        [TestMethod]
        public void ContainsText_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.ContainsText;

            BaseReadWriteTest("A1:A5", "ContainsText", type,
                (sheet, address) =>
                {
                    var cf = sheet.ConditionalFormatting.AddContainsText(address);
                    cf.ContainText = "a";
                    return (ExcelConditionalFormattingRule)cf;
                });
        }

        [TestMethod]
        public void NotContainsText_ReadWrite()
        {
            var type = eExcelConditionalFormattingRuleType.NotContainsText;

            BaseReadWriteTest("A1:A5", "NotContainsText", type,
                (sheet, address) =>
                {
                    var cf = sheet.ConditionalFormatting.AddNotContainsText(address);
                    cf.ContainText = "a";
                    return (ExcelConditionalFormattingRule)cf;
                });
        }

        [TestMethod]
        public void NotContainsBlanks_WriteRead()
        {
            BaseReadWriteTest("A1:A5", "NotContainsBlanks", eExcelConditionalFormattingRuleType.NotContainsBlanks,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddNotContainsBlanks(address);
                });
        }

        [TestMethod]
        public void NotContainsErrors_WriteRead()
        {
            BaseReadWriteTest("A1:A5", "NotContainsErrors", eExcelConditionalFormattingRuleType.NotContainsErrors,
                (sheet, address) =>
                {
                    return (ExcelConditionalFormattingRule)sheet.ConditionalFormatting.AddNotContainsErrors(address);
                });
        }

        private void BaseReadWriteTest(string address, string wsName, eExcelConditionalFormattingRuleType type,
            Func<ExcelWorksheet, ExcelAddress, ExcelConditionalFormattingRule> MakeCF)
        {
            ExcelWorksheet origWS, ws;
            var package = GenerateWorkSheets(wsName, out origWS, out ws);

            var cf1 = MakeCF(origWS, new ExcelAddress(address));
            var cf2 = MakeCF(ws, new ExcelAddress(address));

            ApplyColorStyle(cf1, cf2);
            TestReadWrite(package, cf2, type);
        }

        private ExcelPackage GenerateWorkSheets(string name, out ExcelWorksheet sheet1, out ExcelWorksheet sheet2)
        {
            var package = new ExcelPackage();
            sheet1 = _pck.Workbook.Worksheets.Add(name);
            sheet2 = package.Workbook.Worksheets.Add(name);

            return package;
        }

        private void ApplyColorStyle(ExcelConditionalFormattingRule origCF, ExcelConditionalFormattingRule cf)
        {
            origCF.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            origCF.Style.Fill.BackgroundColor.Color = Color.Aquamarine;

            cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cf.Style.Fill.BackgroundColor.Color = Color.Aquamarine;
        }

        private void TestReadWrite(ExcelPackage package, ExcelConditionalFormattingRule cf, eExcelConditionalFormattingRuleType type)
        {
            var stream = new MemoryStream();
            package.SaveAs(stream);

            ExcelPackage package2 = new ExcelPackage(stream);

            var cf2 = package.Workbook.Worksheets[0].ConditionalFormatting[0];

            Assert.AreEqual(cf.Formula, cf2.Formula);
            Assert.AreEqual(cf.Formula2, cf2.Formula2);

            Assert.AreEqual(cf.Text, cf2.Text);
            Assert.AreEqual(cf2.Type, type);

            var stream2 = new MemoryStream();
            package2.SaveAs("C:\\epplusTest\\Workbooks\\cf.xlsx");
        }

        ExcelConditionalFormattingCollection SavePackageReadCollection(ExcelPackage pck)
        {
            var stream = new MemoryStream();
            pck.SaveAs(stream);

            var cf = new ExcelPackage(stream);
            return cf.Workbook.Worksheets[0].ConditionalFormatting;
        }

        [TestMethod]
        public void ReadWriteBetween()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("BetweenWorksheet");

            var betweenFormatting = sheet.ConditionalFormatting.AddBetween(new ExcelAddress(1, 1, 10, 1));
            betweenFormatting.Formula = "3";
            betweenFormatting.Formula2 = "8";

            betweenFormatting.Style.Fill.BackgroundColor.Color = Color.Red;
            betweenFormatting.Style.Font.Color.Color = Color.Orange;

            var cf = SavePackageReadCollection(pck)[0];

            Assert.AreEqual(cf.Formula, "3");
            Assert.AreEqual(cf.Formula2, "8");

            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule((ExcelConditionalFormattingRule)betweenFormatting);
        }

        [TestMethod]
        public void ReadWriteLess()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("LessThan");

            var lessFormatting = sheet.ConditionalFormatting.AddLessThan(new ExcelAddress(1, 2, 10, 2));
            lessFormatting.Formula = "7";

            lessFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
            lessFormatting.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];

            Assert.AreEqual(cf.Formula, "7");
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule((ExcelConditionalFormattingRule)lessFormatting);
        }

        [TestMethod]
        public void ReadWriteEqual()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("LessThan");

            var equalFormatting = sheet.ConditionalFormatting.AddEqual(new ExcelAddress(1, 4, 10, 4));
            equalFormatting.Formula = "1";

            equalFormatting.Style.Fill.BackgroundColor.Color = Color.Black;
            equalFormatting.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];

            Assert.AreEqual(cf.Formula, "1");

            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule((ExcelConditionalFormattingRule)equalFormatting);
        }

        [TestMethod]
        public void ReadWriteNotEqual() { }

        [TestMethod]
        public void ReadWriteTextContains()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("TextContains");

            var textContains = sheet.ConditionalFormatting.AddTextContains(new ExcelAddress(1, 6, 10, 6));
            textContains.ContainText = "o";

            textContains.Style.Fill.BackgroundColor.Color = Color.Blue;
            textContains.Style.Font.Color.Color = Color.Yellow;

            var cf = SavePackageReadCollection(pck)[0];

            Assert.AreEqual(((IExcelConditionalFormattingContainsText)cf).ContainText, "o");

            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);
        }

        [TestMethod]
        public void ReadWriteLast7Days()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Last7Days");

            var sevenDays = sheet.ConditionalFormatting.AddLast7Days(new ExcelAddress(1, 8, 10, 8));

            sevenDays.Style.Fill.BackgroundColor.Color = Color.Red;
            sevenDays.Style.Font.Color.Color = Color.Yellow;

            var cf = SavePackageReadCollection(pck)[0];

            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.Last7Days);
        }

        [TestMethod] public void ReadWriteYesterday()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Yesterday");

            var yesterdayFormatting = sheet.ConditionalFormatting.AddYesterday(new ExcelAddress(1, 9, 10, 9));

            yesterdayFormatting.Style.Fill.BackgroundColor.Color = Color.Gray;
            yesterdayFormatting.Style.Font.Color.Color = Color.Red;

            var cf = SavePackageReadCollection(pck)[0];

            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.Yesterday);
        }

        [TestMethod]
        public void ReadWriteToday()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Today");

            var todayFormatting = sheet.ConditionalFormatting.AddToday(new ExcelAddress(1, 10, 10, 10));

            todayFormatting.Style.Fill.BackgroundColor.Color = Color.Yellow;
            todayFormatting.Style.Font.Color.Color = Color.Green;
            todayFormatting.Priority = 2;

            var cf = SavePackageReadCollection(pck)[0];

            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.Today);
        }

        [TestMethod]
        public void ReadWriteTommorow()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Tomorrow");

            var tomorrow = sheet.ConditionalFormatting.AddTomorrow(new ExcelAddress(1, 11, 10, 11));

            tomorrow.Style.Fill.BackgroundColor.Color = Color.Green;
            tomorrow.Style.Font.Color.Color = Color.Orange;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.Tomorrow);
        }

        [TestMethod]
        public void ReadWriteLastWeek()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("LastWeek");

            var lastWeek = sheet.ConditionalFormatting.AddLastWeek(new ExcelAddress(1, 12, 20, 12));

            lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
            lastWeek.Style.Font.Color.Color = Color.Violet;


            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.LastWeek);
        }

        [TestMethod]
        public void ReadWriteThisWeek()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("ThisWeek");

            var lastWeek = sheet.ConditionalFormatting.AddThisWeek(new ExcelAddress(1, 13, 20, 13));

            lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
            lastWeek.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.ThisWeek);
        }

        [TestMethod]
        public void ReadWriteNextWeek()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("NextWeek");

            var lastWeek = sheet.ConditionalFormatting.AddNextWeek(new ExcelAddress(1, 14, 20, 14));

            lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
            lastWeek.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.NextWeek);
        }

        [TestMethod]
        public void ReadWriteLastMonth()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("LastMonth");

            var lastWeek = sheet.ConditionalFormatting.AddLastMonth(new ExcelAddress(1, 15, 30, 15));

            lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
            lastWeek.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.LastMonth);
        }

        [TestMethod]
        public void ReadWriteThisMonth()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("ThisMonth");

            var lastWeek = sheet.ConditionalFormatting.AddThisMonth(new ExcelAddress(1, 16, 30, 16));

            lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
            lastWeek.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.ThisMonth);
        }

        [TestMethod]
        public void ReadWriteNextMonth()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("NextMonth");

            var lastWeek = sheet.ConditionalFormatting.AddNextMonth(new ExcelAddress(1, 17, 30, 17));

            lastWeek.Style.Fill.BackgroundColor.Color = Color.Black;
            lastWeek.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.TimePeriod, eExcelConditionalFormattingTimePeriodType.NextMonth);
        }

        [TestMethod]
        public void ReadWriteDuplicate()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Duplicate");

            var duplicateValues = sheet.ConditionalFormatting.AddDuplicateValues(new ExcelAddress(1, 19, 10, 19));

            duplicateValues.Style.Fill.BackgroundColor.Color = Color.Blue;
            duplicateValues.Style.Font.Color.Color = Color.Yellow;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.DuplicateValues);
        }

        [TestMethod]
        public void ReadWriteTop()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Top");

            var top11 = sheet.ConditionalFormatting.AddTop(new ExcelAddress(1, 21, 20, 21));

            top11.Rank = 11;
            top11.Style.Fill.BackgroundColor.Color = Color.Black;
            top11.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Rank, 11);
            Assert.AreEqual(cf.Bottom, false);
            Assert.AreEqual(cf.Percent, false);
        }

        [TestMethod]
        public void ReadWriteBottom()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Bottom");

            var bot12 = sheet.ConditionalFormatting.AddBottom(new ExcelAddress(1, 22, 20, 22));

            bot12.Rank = 12;
            bot12.Style.Fill.BackgroundColor.Color = Color.Black;
            bot12.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Rank, 12);
            Assert.AreEqual(cf.Bottom, true);
            Assert.AreEqual(cf.Percent, false);
        }

        [TestMethod]
        public void ReadWriteTopPercent()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("TopPercent");

            var top13Percent = sheet.ConditionalFormatting.AddTopPercent(new ExcelAddress(1, 23, 20, 23));

            top13Percent.Rank = 13;
            top13Percent.Style.Fill.BackgroundColor.Color = Color.Black;
            top13Percent.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Rank, 13);
            Assert.AreEqual(cf.Bottom, false);
            Assert.AreEqual(cf.Percent, true);
        }

        [TestMethod]
        public void ReadWriteBottomPercent()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("BottomPercent");

            var bot14Percent = sheet.ConditionalFormatting.AddBottomPercent(new ExcelAddress(1, 24, 20, 24));

            bot14Percent.Rank = 14;
            bot14Percent.Style.Fill.BackgroundColor.Color = Color.Black;
            bot14Percent.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Rank, 14);
            Assert.AreEqual(cf.Bottom, true);
            Assert.AreEqual(cf.Percent, true);
        }


        //var belowAverage = wks.ConditionalFormatting.AddBelowAverage(new ExcelAddress(1, 32, 20, 32));

        //belowAverage.Style.Fill.BackgroundColor.Color = Color.Black;
        //        belowAverage.Style.Font.Color.Color = Color.Violet;

        //        var belowEqualAverage = wks.ConditionalFormatting.AddBelowOrEqualAverage(new ExcelAddress(1, 33, 20, 33));

        //belowEqualAverage.Style.Fill.BackgroundColor.Color = Color.Black;
        //        belowEqualAverage.Style.Font.Color.Color = Color.Violet;

        //        var aboveStdDev = wks.ConditionalFormatting.AddAboveStdDev(new ExcelAddress(1, 35, 10, 35));

        //aboveStdDev.Style.Fill.BackgroundColor.Color = Color.Black;
        //        aboveStdDev.Style.Font.Color.Color = Color.Violet;

        //        aboveStdDev.StdDev = 1;

        [TestMethod]
        public void ReadWriteAboveAverage()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("AboveAverage");

            var aboveAverage = sheet.ConditionalFormatting.AddAboveAverage(new ExcelAddress(1, 26, 10, 26));

            aboveAverage.Style.Fill.BackgroundColor.Color = Color.Black;
            aboveAverage.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.AboveAverage, true);
            Assert.AreEqual(cf.EqualAverage, false);
        }

        [TestMethod]
        public void ReadWriteAboveOrEqualAverage()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("AboveAverage");

            var aboveAverage = sheet.ConditionalFormatting.AddAboveOrEqualAverage(new ExcelAddress(1, 27, 10, 27));

            aboveAverage.Style.Fill.BackgroundColor.Color = Color.Black;
            aboveAverage.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.AboveAverage, true);
            Assert.AreEqual(cf.EqualAverage, true);
        }

        [TestMethod]
        public void ReadWriteBelowAverage()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("BelowAverage");

            var belowAverage = sheet.ConditionalFormatting.AddBelowAverage(new ExcelAddress(1, 28, 10, 28));

            belowAverage.Style.Fill.BackgroundColor.Color = Color.Black;
            belowAverage.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.AboveAverage, false);
            Assert.AreEqual(cf.EqualAverage, false);
        }


        [TestMethod]
        public void ReadWriteBelowOrEqualAverage()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("BelowOrEqualAverage");

            var belowEqualAverage = sheet.ConditionalFormatting.AddBelowOrEqualAverage(new ExcelAddress(1, 29, 10, 29));

            belowEqualAverage.Style.Fill.BackgroundColor.Color = Color.Black;
            belowEqualAverage.Style.Font.Color.Color = Color.Violet;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.AboveAverage, false);
            Assert.AreEqual(cf.EqualAverage, true);
        }

        [TestMethod]
        public void ReadWriteStandardDeviationAboveAverage()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("stdAboveAverage");

            var std1 = sheet.ConditionalFormatting.AddAboveStdDev(new ExcelAddress(1, 31, 10, 31));

            std1.Style.Fill.BackgroundColor.Color = Color.Black;
            std1.Style.Font.Color.Color = Color.Violet;

            std1.StdDev = 1;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.AboveStdDev);
            Assert.AreEqual(cf.StdDev, 1);
        }

        [TestMethod]
        public void ReadWriteStandardDeviationBelowAverage()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("stdBelowAverage");

            var std1 = sheet.ConditionalFormatting.AddBelowStdDev(new ExcelAddress(1, 32, 10, 32));

            std1.Style.Fill.BackgroundColor.Color = Color.Black;
            std1.Style.Font.Color.Color = Color.Violet;

            std1.StdDev = 2;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.BelowStdDev);
            Assert.AreEqual(cf.StdDev, 2);
        }

        [TestMethod]
        public void ReadWriteDataBar()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("dataBar");

            var databar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 34, 10, 34), Color.Aqua);

            databar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
            databar.LowValue.Value = 0;
            databar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
            databar.HighValue.Value = 50;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.DataBar);
            Assert.AreEqual(cf.As.DataBar.LowValue.Value, 0);
            Assert.AreEqual(cf.As.DataBar.HighValue.Value, 50);
        }

        [TestMethod]
        public void ReadWriteTwoColorScale()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("twoColorScale");

            var twoColor = sheet.ConditionalFormatting.AddTwoColorScale(new ExcelAddress(1, 36, 10, 36));

            twoColor.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
            twoColor.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;

            twoColor.LowValue.Value = 5;
            twoColor.HighValue.Value = 80;

            twoColor.LowValue.Color = Color.Gold;
            twoColor.HighValue.Color = Color.Silver;

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.TwoColorScale);
            Assert.AreEqual(cf.As.TwoColorScale.LowValue.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.TwoColorScale.HighValue.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.TwoColorScale.LowValue.Value, 5);
            Assert.AreEqual(cf.As.TwoColorScale.HighValue.Value, 80);
            Assert.AreEqual(cf.As.TwoColorScale.LowValue.Color.ToColorString(), Color.Gold.ToColorString());
            Assert.AreEqual(cf.As.TwoColorScale.HighValue.Color.ToColorString(), Color.Silver.ToColorString());
        }

        [TestMethod]
        public void ReadWriteThreeColorScale()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("threeColorScale");

            var threeColor = sheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress(1, 37, 10, 37));

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.ThreeColorScale);
            Assert.AreEqual(cf.As.ThreeColorScale.LowValue.Type, eExcelConditionalFormattingValueObjectType.Min);
            Assert.AreEqual(cf.As.ThreeColorScale.MiddleValue.Type, eExcelConditionalFormattingValueObjectType.Percentile);
            Assert.AreEqual(cf.As.ThreeColorScale.HighValue.Type, eExcelConditionalFormattingValueObjectType.Max);
            Assert.AreEqual(cf.As.ThreeColorScale.MiddleValue.Value, 50);
        }

        [TestMethod]
        public void ReadWriteThreeIcon()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("threeIcon");

            var threeColor = sheet.ConditionalFormatting.AddThreeIconSet(new ExcelAddress(1, 39, 10, 39),
                                                                         eExcelconditionalFormatting3IconsSetType.Symbols2);

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.ThreeIconSet);
            Assert.AreEqual(cf.As.ThreeIconSet.IconSet, eExcelconditionalFormatting3IconsSetType.Symbols2);
            Assert.AreEqual(cf.As.ThreeIconSet.Icon1.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.ThreeIconSet.Icon2.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.ThreeIconSet.Icon3.Type, eExcelConditionalFormattingValueObjectType.Percent);

            Assert.AreEqual(cf.As.ThreeIconSet.Icon1.Value, 0);
            Assert.AreEqual(cf.As.ThreeIconSet.Icon2.Value, Math.Round(100D / 3, 0));
            Assert.AreEqual(cf.As.ThreeIconSet.Icon3.Value, Math.Round(100D * (2D / 3), 0));
        }

        [TestMethod]
        public void ReadWriteFourIcon()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("fourIcon");

            var fourIcons = sheet.ConditionalFormatting.AddFourIconSet(new ExcelAddress(1, 40, 10, 40),
                                                                       eExcelconditionalFormatting4IconsSetType.RedToBlack);
            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.FourIconSet);
            Assert.AreEqual(cf.As.FourIconSet.IconSet, eExcelconditionalFormatting4IconsSetType.RedToBlack);
            Assert.AreEqual(cf.As.FourIconSet.Icon1.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.FourIconSet.Icon2.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.FourIconSet.Icon3.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.FourIconSet.Icon4.Type, eExcelConditionalFormattingValueObjectType.Percent);

            Assert.AreEqual(cf.As.FourIconSet.Icon1.Value, 0);
            Assert.AreEqual(cf.As.FourIconSet.Icon2.Value, Math.Round(100D / 4, 0));
            Assert.AreEqual(cf.As.FourIconSet.Icon3.Value, Math.Round(100D * (2D / 4), 0));
            Assert.AreEqual(cf.As.FourIconSet.Icon4.Value, 75);
        }


        [TestMethod]
        public void ReadWriteFiveIcon()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("fiveIcon");

            var fiveIcons = sheet.ConditionalFormatting.AddFiveIconSet(new ExcelAddress(1, 41, 10, 41), eExcelconditionalFormatting5IconsSetType.Rating);

            var cf = SavePackageReadCollection(pck)[0];
            var ws = _pck.Workbook.Worksheets.GetByName("Overview");
            ws.ConditionalFormatting.CopyRule(cf);

            Assert.AreEqual(cf.Type, eExcelConditionalFormattingRuleType.FiveIconSet);
            Assert.AreEqual(cf.As.FiveIconSet.IconSet, eExcelconditionalFormatting5IconsSetType.Rating);
            Assert.AreEqual(cf.As.FiveIconSet.Icon1.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.FiveIconSet.Icon2.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.FiveIconSet.Icon3.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.FiveIconSet.Icon4.Type, eExcelConditionalFormattingValueObjectType.Percent);
            Assert.AreEqual(cf.As.FiveIconSet.Icon5.Type, eExcelConditionalFormattingValueObjectType.Percent);

            Assert.AreEqual(cf.As.FiveIconSet.Icon1.Value, 0);
            Assert.AreEqual(cf.As.FiveIconSet.Icon2.Value, 20);
            Assert.AreEqual(cf.As.FiveIconSet.Icon3.Value, 40);
            Assert.AreEqual(cf.As.FiveIconSet.Icon4.Value, 60);
            Assert.AreEqual(cf.As.FiveIconSet.Icon5.Value, 80);
        }

        //var threeIcons = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress(1, 43, 10, 43), eExcelconditionalFormatting3IconsSetType.Symbols2);

        //var fourIcons = wks.ConditionalFormatting.AddFourIconSet(new ExcelAddress(1, 44, 10, 44), eExcelconditionalFormatting4IconsSetType.RedToBlack);

        //var fiveIcons = wks.ConditionalFormatting.AddFiveIconSet(new ExcelAddress(1, 45, 10, 45), eExcelconditionalFormatting5IconsSetType.Rating);

        //var threeGreatherThan = wks.ConditionalFormatting.AddThreeIconSet(new ExcelAddress(1, 48, 10, 48), eExcelconditionalFormatting3IconsSetType.TrafficLights2);

        //threeGreatherThan.Icon2.GreaterThanOrEqualTo = false;
        //        threeGreatherThan.Icon3.GreaterThanOrEqualTo = false;

        [TestMethod]
        public void PriorityChangeTest()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Today");

            var yesterdayFormatting = sheet.ConditionalFormatting.AddToday(new ExcelAddress(1, 11, 10, 11));

            yesterdayFormatting.Style.Fill.BackgroundColor.Color = Color.Gray;
            yesterdayFormatting.Style.Font.Color.Color = Color.Red;
            yesterdayFormatting.Priority = 2;

            var yesterdayFormatting2 = sheet.ConditionalFormatting.AddToday(new ExcelAddress(1, 11, 10, 11));

            yesterdayFormatting2.Style.Fill.BackgroundColor.Color = Color.Black;
            yesterdayFormatting2.Style.Font.Color.Color = Color.Violet;
            yesterdayFormatting2.Priority = 1;

            string date = $"{DateTime.Now.Year}-{DateTime.Now.Month}-{DateTime.Now.Day}";

            sheet.Cells[1, 11, 10, 11].Value = date;

            var stream = new MemoryStream();
            pck.SaveAs(stream);

            var newPck = new ExcelPackage(stream);

            Assert.AreEqual(Color.FromArgb(255, Color.Gray.R, Color.Gray.G, Color.Gray.B), newPck.Workbook.Worksheets[0].
                            ConditionalFormatting[0].Style.Fill.
                            BackgroundColor.Color);
            newPck.SaveAs("C:\\epplusTest\\Testoutput\\cfPriorityChangeTest2.xlsx");
        }


        [TestMethod]
        public void PriorityChangeTest2()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("Today");

            var yesterdayFormatting = sheet.ConditionalFormatting.AddToday(new ExcelAddress(1, 11, 10, 11));

            yesterdayFormatting.Style.Fill.BackgroundColor.Color = Color.Gray;
            yesterdayFormatting.Style.Font.Color.Color = Color.Red;
            yesterdayFormatting.Priority = 2;

            var yesterdayFormatting2 = sheet.ConditionalFormatting.AddToday(new ExcelAddress(1, 11, 10, 11));

            yesterdayFormatting2.Style.Fill.BackgroundColor.Color = Color.Black;
            yesterdayFormatting2.Style.Font.Color.Color = Color.Violet;
            yesterdayFormatting2.Priority = 1;

            string date = $"{DateTime.Now.Year}-{DateTime.Now.Month}-{DateTime.Now.Day}";

            sheet.Cells[1, 11, 10, 11].Value = date;

            var stream = new MemoryStream();
            pck.SaveAs("C:\\epplusTest\\Testoutput\\cfPriorityChangeTest1.xlsx");

            var newPck = new ExcelPackage("C:\\epplusTest\\Testoutput\\cfPriorityChangeTest1.xlsx");

            Assert.AreEqual(Color.FromArgb(255, Color.Gray.R, Color.Gray.G, Color.Gray.B), newPck.Workbook.Worksheets[0].
                            ConditionalFormatting[0].Style.Fill.
                            BackgroundColor.Color);
            newPck.SaveAs("C:\\epplusTest\\Testoutput\\cfPriorityChangeTest2.xlsx");
        }

        [TestMethod]
        public void MultipleValidationsOneRange()
        {
            var pck = new ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("MultipleOneRange");

        }

        //TODO: We should most likely throw a clearer exception.
        [TestMethod]
        public void CFThrowsIfDatabarValueNotSetOnSave()
        {
            ExcelPackage pck = new ExcelPackage(new MemoryStream());

            var sheet = pck.Workbook.Worksheets.Add("DatabarValueTest");

            var databar = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Green);

            databar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
            databar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;

            var stream = new MemoryStream();
            pck.SaveAs(stream);
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

                for (int i = 0; i < readPck.Workbook.Worksheets[0].ConditionalFormatting.Count; i++)
                {
                    var format = readPck.Workbook.Worksheets[0].ConditionalFormatting[i];

                    Assert.AreEqual(format.Formula, "3");
                    Assert.AreEqual(Color.Black.ToArgb(), format.Style.Fill.BackgroundColor.Color.Value.ToArgb());
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

        [TestMethod]
        public void PriorityTestOverlap()
        {
            using (var pck = OpenPackage("CFPriorityTest.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("priorityTest");

                sheet.Cells["A1:A7"].Value = "A Player's handbook";
                sheet.Cells["A1:A7"].AutoFitColumns();
                sheet.Cells["B1"].Value = "A1:A5 should be green, A6 yellow, A7 red";
                sheet.Cells["B1"].AutoFitColumns();

                var cfHighestPriority = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1:A5"));

                cfHighestPriority.ContainText = "A";
                cfHighestPriority.Style.Fill.BackgroundColor.Color = Color.Green;

                var cfMiddlePriority = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1:A6"));

                cfMiddlePriority.ContainText = "A";
                cfMiddlePriority.Style.Fill.BackgroundColor.Color = Color.Yellow;

                var cfLowestPriority = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1:A7"));
                cfLowestPriority.Style.Fill.BackgroundColor.Color = Color.Red;
                cfLowestPriority.ContainText = "A";

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void PriorityTestExtLst()
        {
            using (var pck = OpenPackage("CFPriorityTestExtLst.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("priorityTest");

                sheet.Cells["A1:A7"].Formula = "=Row()";
                sheet.Cells["A1:A7"].AutoFitColumns();
                sheet.Cells["B1"].Value = "A1:A5 should be green, A6 yellow, A7 red";
                sheet.Cells["B1"].AutoFitColumns();

                var cfHighestPriorityExt = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A5"), Color.Green);

                var cfMiddlePriorityExt = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A6"), Color.Yellow);

                var cfLowestPriorityExt = sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("A1:A7"), Color.Red);

                sheet.Cells["B1"].Value = "A1:A5 should be green, A6 yellow, A7 red";
                sheet.Cells["B1"].AutoFitColumns();

                var cfHighestPriority = sheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A5"));

                cfHighestPriority.Formula = "0";
                cfHighestPriority.Style.Fill.BackgroundColor.Color = Color.Orange;

                var cfMiddlePriority = sheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A6"));

                cfMiddlePriority.Formula = "0";
                cfMiddlePriority.Style.Fill.BackgroundColor.Color = Color.Silver;

                var cfLowestPriority = sheet.ConditionalFormatting.AddGreaterThan(new ExcelAddress("A1:A7"));
                cfLowestPriority.Style.Fill.BackgroundColor.Color = Color.Yellow;
                cfLowestPriority.Formula = "0";

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void PriorityTestChangedOrder()
        {
            using (var pck = OpenPackage("CFPriorityTestChangedOrder.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("priorityTest");

                sheet.Cells["A1:A7"].Value = "A Player's handbook";
                sheet.Cells["A1:A7"].AutoFitColumns();
                sheet.Cells["B1"].Value = "A1:A5 should be green, A6 yellow, A7 red";
                sheet.Cells["B1"].AutoFitColumns();

                var cfHighestPriority = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1:A5"));

                cfHighestPriority.ContainText = "A";
                cfHighestPriority.Style.Fill.BackgroundColor.Color = Color.Green;

                var cfMiddlePriority = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1:A5"));

                cfMiddlePriority.ContainText = "A";
                cfMiddlePriority.Style.Fill.BackgroundColor.Color = Color.Yellow;

                var cfLowestPriority = sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1:A5"));
                cfLowestPriority.Style.Fill.BackgroundColor.Color = Color.Red;
                cfLowestPriority.ContainText = "A";

                cfLowestPriority.Priority = 1;
                cfMiddlePriority.Priority = 3;

                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void ExtLstFormulaValidations()
        {

            using (var pck = OpenPackage("ExtLstFormulas.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                refSheet.Cells["B5"].Value = 5;

                sheet.Cells["B1:B5"].Value = 5;
                sheet.Cells["B3"].Value = 2;

                //var validation = sheet.Cells["C5"].DataValidation.AddIntegerDataValidation();
                //validation.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.equal;
                //validation.Formula.ExcelFormula = "formulasReference!$B$5";

                //sheet.ConditionalFormatting.AddBeginsWith(new ExcelAddress("A1"));
                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";
                equal.Style.Fill.BackgroundColor.Color = Color.Blue;
                equal.Style.Font.Italic = true;

                //sheet.ConditionalFormatting.AddDatabar(new ExcelAddress("C1:C205"), Color.Beige);

                SaveAndCleanup(pck);
            }
        }


        [TestMethod]
        public void ExtLstWithDxf() 
        {
            using (var pck = OpenPackage("ExtLstFormulasDxf.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("formulas");
                var refSheet = pck.Workbook.Worksheets.Add("formulasReference");

                var equal = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("B1:B5"));
                equal.Formula = "formulasReference!$B$5";
                equal.Style.Fill.BackgroundColor.Color = Color.Blue;
                equal.Style.Font.Italic = true;
                equal.Style.Font.Bold = true;

                var equal2 = sheet.ConditionalFormatting.AddEqual(new ExcelAddress("C1:C5"));
                equal2.Formula = "formulasReference!$B$1";
                equal2.Style.Fill.Style = OfficeOpenXml.Style.eDxfFillStyle.GradientFill;
                var c1 = equal2.Style.Fill.Gradient.Colors.Add(0);
                var c2 = equal2.Style.Fill.Gradient.Colors.Add(100);

                equal2.Style.Fill.Gradient.Degree = 90;

                c1.Color.SetColor(Color.LightGreen);
                c2.Color.SetColor(Color.MediumPurple);

                //equal.Style.Fill.Gradient.Colors[0] = new ExcelDxfGradientFill(ExcelStyles.st;

                SaveAndCleanup(pck);
            }
        }
    }
}