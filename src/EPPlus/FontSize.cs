/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using OfficeOpenXml.Packaging.Ionic.Zip;
using System.Reflection;
using System.IO;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Defines font size in pixels for different font families and sized used when determining auto widths for columns.
    /// This is used as .NET and Excel does not measure font widths in pixels in a similar way.
    /// </summary>
    public class FontSizeInfo
    {
        /// <summary>
        /// Construtor
        /// </summary>
        /// <param name="height">Height in pixels</param>
        /// <param name="width">Width in pixels</param>
        public FontSizeInfo(float height, float width)
        {
            Width = width;
            Height = height;
        }
        /// <summary>
        /// Height in pixels
        /// </summary>
        public float Height { get; set; }
        /// <summary>
        /// Width in pixels
        /// </summary>
        public float Width { get; set; }
    }
    /// <summary>
    /// A collection of fonts and there size in pixels used when determining auto widths for columns.
    /// This is used as .NET and Excel does not measure font widths in pixels in a similar way.
    /// </summary>
    public static class FontSize
    {
        /// <summary>
        /// Dictionary containing Font Width and heights in pixels.
        /// You can add your own fonts and sizes here.
        /// </summary>
        public static Dictionary<string, Dictionary<float, FontSizeInfo>> FontHeights = new Dictionary<string, Dictionary<float, FontSizeInfo>>(StringComparer.OrdinalIgnoreCase)
        {
            {"Arial",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 6)},{9,new FontSizeInfo(16, 7)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(19, 8)},{12,new FontSizeInfo(20, 9)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(24, 11)},{15,new FontSizeInfo(25, 11)},{16,new FontSizeInfo(27, 12)},{17,new FontSizeInfo(29, 14)},{18,new FontSizeInfo(31, 14)},{20,new FontSizeInfo(34, 16)},{22,new FontSizeInfo(36, 17)},{24,new FontSizeInfo(40, 19)},{26,new FontSizeInfo(44, 20)},{28,new FontSizeInfo(46, 22)},{30,new FontSizeInfo(50, 23)},{32,new FontSizeInfo(54, 25)},{34,new FontSizeInfo(56, 26)},{36,new FontSizeInfo(59, 28)},{38,new FontSizeInfo(63, 29)},{40,new FontSizeInfo(66, 31)},{44,new FontSizeInfo(73, 35)},{48,new FontSizeInfo(79, 38)},{52,new FontSizeInfo(85, 40)},{56,new FontSizeInfo(92, 44)},{60,new FontSizeInfo(100, 46)},{64,new FontSizeInfo(107, 50)},{68,new FontSizeInfo(114, 54)},{72,new FontSizeInfo(120, 56)},{96,new FontSizeInfo(159, 75)},{128,new FontSizeInfo(213, 101)},{256,new FontSizeInfo(424, 202)},}},
            {"Arial Black",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(17, 7)},{9,new FontSizeInfo(19, 8)},{10,new FontSizeInfo(20, 9)},{11,new FontSizeInfo(25, 10)},{12,new FontSizeInfo(26, 11)},{13,new FontSizeInfo(27, 11)},{14,new FontSizeInfo(30, 14)},{15,new FontSizeInfo(31, 14)},{16,new FontSizeInfo(33, 15)},{17,new FontSizeInfo(35, 16)},{18,new FontSizeInfo(36, 17)},{20,new FontSizeInfo(42, 19)},{22,new FontSizeInfo(45, 20)},{24,new FontSizeInfo(49, 22)},{26,new FontSizeInfo(55, 24)},{28,new FontSizeInfo(57, 26)},{30,new FontSizeInfo(61, 28)},{32,new FontSizeInfo(65, 31)},{34,new FontSizeInfo(70, 32)},{36,new FontSizeInfo(74, 34)},{38,new FontSizeInfo(78, 36)},{40,new FontSizeInfo(80, 37)},{44,new FontSizeInfo(90, 41)},{48,new FontSizeInfo(97, 45)},{52,new FontSizeInfo(105, 49)},{56,new FontSizeInfo(115, 53)},{60,new FontSizeInfo(122, 56)},{64,new FontSizeInfo(130, 60)},{68,new FontSizeInfo(138, 65)},{72,new FontSizeInfo(147, 68)},{96,new FontSizeInfo(195, 90)},{128,new FontSizeInfo(259, 121)},{256,new FontSizeInfo(516, 241)},}},
            {"Arial Narrow",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(17, 5)},{9,new FontSizeInfo(18, 5)},{10,new FontSizeInfo(17, 6)},{11,new FontSizeInfo(22, 7)},{12,new FontSizeInfo(21, 7)},{13,new FontSizeInfo(23, 8)},{14,new FontSizeInfo(24, 9)},{15,new FontSizeInfo(26, 9)},{16,new FontSizeInfo(27, 10)},{17,new FontSizeInfo(30, 10)},{18,new FontSizeInfo(31, 11)},{20,new FontSizeInfo(34, 12)},{22,new FontSizeInfo(36, 14)},{24,new FontSizeInfo(40, 16)},{26,new FontSizeInfo(45, 17)},{28,new FontSizeInfo(47, 18)},{30,new FontSizeInfo(50, 19)},{32,new FontSizeInfo(54, 21)},{34,new FontSizeInfo(56, 22)},{36,new FontSizeInfo(61, 23)},{38,new FontSizeInfo(63, 24)},{40,new FontSizeInfo(66, 25)},{44,new FontSizeInfo(74, 28)},{48,new FontSizeInfo(80, 31)},{52,new FontSizeInfo(86, 33)},{56,new FontSizeInfo(92, 36)},{60,new FontSizeInfo(99, 38)},{64,new FontSizeInfo(105, 41)},{68,new FontSizeInfo(112, 44)},{72,new FontSizeInfo(118, 46)},{96,new FontSizeInfo(157, 61)},{128,new FontSizeInfo(210, 83)},{256,new FontSizeInfo(418, 165)},}},
            {"Bookman Old Style",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(17, 7)},{9,new FontSizeInfo(17, 7)},{10,new FontSizeInfo(20, 8)},{11,new FontSizeInfo(20, 9)},{12,new FontSizeInfo(21, 10)},{13,new FontSizeInfo(22, 11)},{14,new FontSizeInfo(24, 12)},{15,new FontSizeInfo(26, 12)},{16,new FontSizeInfo(27, 14)},{17,new FontSizeInfo(29, 15)},{18,new FontSizeInfo(31, 16)},{20,new FontSizeInfo(34, 18)},{22,new FontSizeInfo(37, 19)},{24,new FontSizeInfo(42, 21)},{26,new FontSizeInfo(44, 23)},{28,new FontSizeInfo(47, 24)},{30,new FontSizeInfo(50, 26)},{32,new FontSizeInfo(54, 28)},{34,new FontSizeInfo(57, 29)},{36,new FontSizeInfo(61, 32)},{38,new FontSizeInfo(64, 34)},{40,new FontSizeInfo(68, 35)},{44,new FontSizeInfo(76, 39)},{48,new FontSizeInfo(82, 42)},{52,new FontSizeInfo(88, 45)},{56,new FontSizeInfo(96, 50)},{60,new FontSizeInfo(101, 53)},{64,new FontSizeInfo(108, 56)},{68,new FontSizeInfo(115, 59)},{72,new FontSizeInfo(123, 63)},{96,new FontSizeInfo(163, 84)},{128,new FontSizeInfo(218, 112)},{256,new FontSizeInfo(437, 224)},}},
            {"Calibri",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 6)},{9,new FontSizeInfo(16, 6)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(20, 7)},{12,new FontSizeInfo(21, 8)},{13,new FontSizeInfo(23, 9)},{14,new FontSizeInfo(25, 10)},{15,new FontSizeInfo(26, 10)},{16,new FontSizeInfo(28, 11)},{17,new FontSizeInfo(30, 12)},{18,new FontSizeInfo(31, 12)},{20,new FontSizeInfo(35, 15)},{22,new FontSizeInfo(38, 16)},{24,new FontSizeInfo(42, 17)},{26,new FontSizeInfo(45, 19)},{28,new FontSizeInfo(48, 20)},{30,new FontSizeInfo(52, 21)},{32,new FontSizeInfo(56, 23)},{34,new FontSizeInfo(58, 24)},{36,new FontSizeInfo(62, 25)},{38,new FontSizeInfo(66, 27)},{40,new FontSizeInfo(68, 28)},{44,new FontSizeInfo(76, 32)},{48,new FontSizeInfo(82, 34)},{52,new FontSizeInfo(89, 37)},{56,new FontSizeInfo(96, 40)},{60,new FontSizeInfo(102, 43)},{64,new FontSizeInfo(109, 45)},{68,new FontSizeInfo(116, 49)},{72,new FontSizeInfo(123, 52)},{96,new FontSizeInfo(163, 69)},{128,new FontSizeInfo(218, 92)},{256,new FontSizeInfo(434, 184)},}},
            {"Calibri Light",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 6)},{9,new FontSizeInfo(16, 6)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(20, 8)},{12,new FontSizeInfo(21, 8)},{13,new FontSizeInfo(23, 9)},{14,new FontSizeInfo(25, 10)},{15,new FontSizeInfo(26, 10)},{16,new FontSizeInfo(28, 11)},{17,new FontSizeInfo(30, 12)},{18,new FontSizeInfo(31, 12)},{20,new FontSizeInfo(35, 15)},{22,new FontSizeInfo(38, 16)},{24,new FontSizeInfo(42, 17)},{26,new FontSizeInfo(45, 19)},{28,new FontSizeInfo(48, 20)},{30,new FontSizeInfo(52, 21)},{32,new FontSizeInfo(56, 23)},{34,new FontSizeInfo(58, 24)},{36,new FontSizeInfo(62, 25)},{38,new FontSizeInfo(66, 27)},{40,new FontSizeInfo(68, 28)},{44,new FontSizeInfo(76, 32)},{48,new FontSizeInfo(82, 34)},{52,new FontSizeInfo(89, 37)},{56,new FontSizeInfo(96, 40)},{60,new FontSizeInfo(102, 43)},{64,new FontSizeInfo(109, 45)},{68,new FontSizeInfo(116, 49)},{72,new FontSizeInfo(123, 52)},{96,new FontSizeInfo(163, 69)},{128,new FontSizeInfo(218, 92)},{256,new FontSizeInfo(434, 184)},}},
            {"Calisto MT",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 6)},{9,new FontSizeInfo(16, 6)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(19, 8)},{12,new FontSizeInfo(21, 8)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(24, 10)},{15,new FontSizeInfo(26, 10)},{16,new FontSizeInfo(27, 11)},{17,new FontSizeInfo(29, 12)},{18,new FontSizeInfo(31, 12)},{20,new FontSizeInfo(34, 15)},{22,new FontSizeInfo(37, 16)},{24,new FontSizeInfo(40, 17)},{26,new FontSizeInfo(44, 19)},{28,new FontSizeInfo(46, 20)},{30,new FontSizeInfo(50, 21)},{32,new FontSizeInfo(54, 23)},{34,new FontSizeInfo(56, 24)},{36,new FontSizeInfo(60, 25)},{38,new FontSizeInfo(64, 27)},{40,new FontSizeInfo(66, 28)},{44,new FontSizeInfo(74, 32)},{48,new FontSizeInfo(80, 35)},{52,new FontSizeInfo(86, 37)},{56,new FontSizeInfo(93, 40)},{60,new FontSizeInfo(99, 43)},{64,new FontSizeInfo(106, 45)},{68,new FontSizeInfo(113, 49)},{72,new FontSizeInfo(119, 52)},{96,new FontSizeInfo(158, 69)},{128,new FontSizeInfo(211, 92)},{256,new FontSizeInfo(420, 185)},}},
            {"Cambria",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(14, 6)},{9,new FontSizeInfo(16, 7)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(19, 8)},{12,new FontSizeInfo(21, 9)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(24, 11)},{15,new FontSizeInfo(25, 11)},{16,new FontSizeInfo(27, 12)},{17,new FontSizeInfo(29, 14)},{18,new FontSizeInfo(30, 14)},{20,new FontSizeInfo(34, 16)},{22,new FontSizeInfo(36, 17)},{24,new FontSizeInfo(40, 19)},{26,new FontSizeInfo(44, 20)},{28,new FontSizeInfo(46, 21)},{30,new FontSizeInfo(50, 23)},{32,new FontSizeInfo(54, 25)},{34,new FontSizeInfo(56, 26)},{36,new FontSizeInfo(60, 28)},{38,new FontSizeInfo(63, 29)},{40,new FontSizeInfo(66, 31)},{44,new FontSizeInfo(73, 35)},{48,new FontSizeInfo(79, 37)},{52,new FontSizeInfo(85, 40)},{56,new FontSizeInfo(93, 44)},{60,new FontSizeInfo(99, 46)},{64,new FontSizeInfo(105, 50)},{68,new FontSizeInfo(112, 53)},{72,new FontSizeInfo(118, 56)},{96,new FontSizeInfo(157, 75)},{128,new FontSizeInfo(210, 101)},{256,new FontSizeInfo(418, 201)},}},
            {"Cambria Math",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(85, 6)},{9,new FontSizeInfo(93, 7)},{10,new FontSizeInfo(102, 7)},{11,new FontSizeInfo(117, 8)},{12,new FontSizeInfo(124, 9)},{13,new FontSizeInfo(132, 9)},{14,new FontSizeInfo(147, 11)},{15,new FontSizeInfo(154, 11)},{16,new FontSizeInfo(162, 12)},{17,new FontSizeInfo(179, 14)},{18,new FontSizeInfo(186, 14)},{20,new FontSizeInfo(209, 16)},{22,new FontSizeInfo(223, 17)},{24,new FontSizeInfo(248, 19)},{26,new FontSizeInfo(270, 20)},{28,new FontSizeInfo(285, 21)},{30,new FontSizeInfo(310, 23)},{32,new FontSizeInfo(332, 25)},{34,new FontSizeInfo(347, 26)},{36,new FontSizeInfo(371, 28)},{38,new FontSizeInfo(394, 29)},{40,new FontSizeInfo(409, 31)},{44,new FontSizeInfo(455, 35)},{48,new FontSizeInfo(493, 37)},{52,new FontSizeInfo(532, 40)},{56,new FontSizeInfo(579, 44)},{60,new FontSizeInfo(616, 46)},{64,new FontSizeInfo(655, 50)},{68,new FontSizeInfo(702, 53)},{72,new FontSizeInfo(739, 56)},{96,new FontSizeInfo(986, 75)},{128,new FontSizeInfo(1317, 101)},{256,new FontSizeInfo(2047, 201)},}},
            {"Century Gothic",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(18, 6)},{9,new FontSizeInfo(19, 7)},{10,new FontSizeInfo(18, 7)},{11,new FontSizeInfo(22, 8)},{12,new FontSizeInfo(23, 9)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(24, 11)},{15,new FontSizeInfo(25, 11)},{16,new FontSizeInfo(26, 12)},{17,new FontSizeInfo(28, 14)},{18,new FontSizeInfo(32, 14)},{20,new FontSizeInfo(35, 16)},{22,new FontSizeInfo(38, 17)},{24,new FontSizeInfo(41, 19)},{26,new FontSizeInfo(44, 20)},{28,new FontSizeInfo(46, 22)},{30,new FontSizeInfo(51, 23)},{32,new FontSizeInfo(54, 25)},{34,new FontSizeInfo(59, 26)},{36,new FontSizeInfo(61, 28)},{38,new FontSizeInfo(66, 29)},{40,new FontSizeInfo(66, 31)},{44,new FontSizeInfo(76, 35)},{48,new FontSizeInfo(82, 37)},{52,new FontSizeInfo(88, 40)},{56,new FontSizeInfo(94, 44)},{60,new FontSizeInfo(100, 46)},{64,new FontSizeInfo(106, 50)},{68,new FontSizeInfo(113, 53)},{72,new FontSizeInfo(119, 56)},{96,new FontSizeInfo(158, 75)},{128,new FontSizeInfo(214, 101)},{256,new FontSizeInfo(427, 201)},}},
            {"Century Schoolbook",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(17, 6)},{9,new FontSizeInfo(18, 7)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(19, 8)},{12,new FontSizeInfo(21, 9)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(24, 11)},{15,new FontSizeInfo(25, 11)},{16,new FontSizeInfo(27, 12)},{17,new FontSizeInfo(29, 14)},{18,new FontSizeInfo(30, 14)},{20,new FontSizeInfo(34, 16)},{22,new FontSizeInfo(36, 17)},{24,new FontSizeInfo(40, 19)},{26,new FontSizeInfo(44, 20)},{28,new FontSizeInfo(46, 22)},{30,new FontSizeInfo(50, 23)},{32,new FontSizeInfo(53, 25)},{34,new FontSizeInfo(56, 26)},{36,new FontSizeInfo(59, 28)},{38,new FontSizeInfo(63, 29)},{40,new FontSizeInfo(65, 31)},{44,new FontSizeInfo(73, 35)},{48,new FontSizeInfo(79, 38)},{52,new FontSizeInfo(85, 40)},{56,new FontSizeInfo(92, 44)},{60,new FontSizeInfo(98, 46)},{64,new FontSizeInfo(104, 50)},{68,new FontSizeInfo(112, 54)},{72,new FontSizeInfo(118, 56)},{96,new FontSizeInfo(157, 75)},{128,new FontSizeInfo(209, 101)},{256,new FontSizeInfo(416, 202)},}},
            {"Corbel",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 6)},{9,new FontSizeInfo(16, 6)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(20, 8)},{12,new FontSizeInfo(21, 8)},{13,new FontSizeInfo(23, 9)},{14,new FontSizeInfo(25, 10)},{15,new FontSizeInfo(26, 10)},{16,new FontSizeInfo(28, 11)},{17,new FontSizeInfo(30, 12)},{18,new FontSizeInfo(31, 14)},{20,new FontSizeInfo(35, 15)},{22,new FontSizeInfo(38, 16)},{24,new FontSizeInfo(42, 18)},{26,new FontSizeInfo(45, 19)},{28,new FontSizeInfo(48, 20)},{30,new FontSizeInfo(52, 22)},{32,new FontSizeInfo(56, 24)},{34,new FontSizeInfo(58, 25)},{36,new FontSizeInfo(62, 26)},{38,new FontSizeInfo(66, 28)},{40,new FontSizeInfo(68, 29)},{44,new FontSizeInfo(76, 33)},{48,new FontSizeInfo(82, 36)},{52,new FontSizeInfo(89, 38)},{56,new FontSizeInfo(96, 41)},{60,new FontSizeInfo(102, 44)},{64,new FontSizeInfo(109, 48)},{68,new FontSizeInfo(116, 51)},{72,new FontSizeInfo(123, 53)},{96,new FontSizeInfo(163, 71)},{128,new FontSizeInfo(218, 95)},{256,new FontSizeInfo(434, 190)},}},
            {"Courier New",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 7)},{9,new FontSizeInfo(16, 7)},{10,new FontSizeInfo(18, 8)},{11,new FontSizeInfo(20, 9)},{12,new FontSizeInfo(21, 10)},{13,new FontSizeInfo(23, 10)},{14,new FontSizeInfo(25, 11)},{15,new FontSizeInfo(26, 12)},{16,new FontSizeInfo(28, 14)},{17,new FontSizeInfo(30, 15)},{18,new FontSizeInfo(32, 15)},{20,new FontSizeInfo(35, 17)},{22,new FontSizeInfo(38, 18)},{24,new FontSizeInfo(42, 20)},{26,new FontSizeInfo(46, 22)},{28,new FontSizeInfo(48, 23)},{30,new FontSizeInfo(52, 25)},{32,new FontSizeInfo(56, 27)},{34,new FontSizeInfo(58, 28)},{36,new FontSizeInfo(62, 31)},{38,new FontSizeInfo(66, 33)},{40,new FontSizeInfo(69, 34)},{44,new FontSizeInfo(76, 37)},{48,new FontSizeInfo(83, 40)},{52,new FontSizeInfo(89, 43)},{56,new FontSizeInfo(97, 48)},{60,new FontSizeInfo(103, 51)},{64,new FontSizeInfo(109, 54)},{68,new FontSizeInfo(117, 58)},{72,new FontSizeInfo(123, 61)},{96,new FontSizeInfo(164, 82)},{128,new FontSizeInfo(219, 109)},{256,new FontSizeInfo(444, 218)},}},
            {"Garamond",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 5)},{9,new FontSizeInfo(16, 6)},{10,new FontSizeInfo(17, 6)},{11,new FontSizeInfo(20, 7)},{12,new FontSizeInfo(21, 8)},{13,new FontSizeInfo(22, 8)},{14,new FontSizeInfo(25, 9)},{15,new FontSizeInfo(26, 9)},{16,new FontSizeInfo(28, 10)},{17,new FontSizeInfo(30, 11)},{18,new FontSizeInfo(31, 11)},{20,new FontSizeInfo(35, 14)},{22,new FontSizeInfo(38, 15)},{24,new FontSizeInfo(41, 16)},{26,new FontSizeInfo(45, 17)},{28,new FontSizeInfo(48, 18)},{30,new FontSizeInfo(52, 20)},{32,new FontSizeInfo(55, 21)},{34,new FontSizeInfo(58, 22)},{36,new FontSizeInfo(62, 24)},{38,new FontSizeInfo(65, 25)},{40,new FontSizeInfo(68, 26)},{44,new FontSizeInfo(76, 29)},{48,new FontSizeInfo(82, 32)},{52,new FontSizeInfo(88, 34)},{56,new FontSizeInfo(96, 37)},{60,new FontSizeInfo(102, 40)},{64,new FontSizeInfo(108, 42)},{68,new FontSizeInfo(116, 45)},{72,new FontSizeInfo(122, 48)},{96,new FontSizeInfo(163, 63)},{128,new FontSizeInfo(217, 85)},{256,new FontSizeInfo(432, 170)},}},
            {"Georgia",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 8)},{9,new FontSizeInfo(16, 8)},{10,new FontSizeInfo(17, 9)},{11,new FontSizeInfo(19, 9)},{12,new FontSizeInfo(20, 10)},{13,new FontSizeInfo(22, 10)},{14,new FontSizeInfo(24, 12)},{15,new FontSizeInfo(26, 12)},{16,new FontSizeInfo(27, 14)},{17,new FontSizeInfo(30, 15)},{18,new FontSizeInfo(31, 16)},{20,new FontSizeInfo(34, 18)},{22,new FontSizeInfo(36, 19)},{24,new FontSizeInfo(40, 21)},{26,new FontSizeInfo(44, 22)},{28,new FontSizeInfo(46, 24)},{30,new FontSizeInfo(50, 26)},{32,new FontSizeInfo(53, 27)},{34,new FontSizeInfo(56, 29)},{36,new FontSizeInfo(60, 31)},{38,new FontSizeInfo(64, 33)},{40,new FontSizeInfo(66, 35)},{44,new FontSizeInfo(73, 38)},{48,new FontSizeInfo(79, 41)},{52,new FontSizeInfo(85, 44)},{56,new FontSizeInfo(93, 49)},{60,new FontSizeInfo(99, 52)},{64,new FontSizeInfo(105, 55)},{68,new FontSizeInfo(112, 59)},{72,new FontSizeInfo(118, 62)},{96,new FontSizeInfo(157, 84)},{128,new FontSizeInfo(209, 111)},{256,new FontSizeInfo(417, 222)},}},
            {"Gill Sans MT",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(18, 6)},{9,new FontSizeInfo(21, 6)},{10,new FontSizeInfo(20, 7)},{11,new FontSizeInfo(23, 8)},{12,new FontSizeInfo(26, 8)},{13,new FontSizeInfo(28, 9)},{14,new FontSizeInfo(29, 10)},{15,new FontSizeInfo(32, 10)},{16,new FontSizeInfo(33, 11)},{17,new FontSizeInfo(36, 12)},{18,new FontSizeInfo(37, 12)},{20,new FontSizeInfo(41, 15)},{22,new FontSizeInfo(43, 16)},{24,new FontSizeInfo(48, 17)},{26,new FontSizeInfo(51, 19)},{28,new FontSizeInfo(56, 20)},{30,new FontSizeInfo(59, 21)},{32,new FontSizeInfo(63, 23)},{34,new FontSizeInfo(67, 24)},{36,new FontSizeInfo(71, 25)},{38,new FontSizeInfo(75, 27)},{40,new FontSizeInfo(78, 28)},{44,new FontSizeInfo(85, 32)},{48,new FontSizeInfo(91, 34)},{52,new FontSizeInfo(101, 37)},{56,new FontSizeInfo(109, 40)},{60,new FontSizeInfo(116, 42)},{64,new FontSizeInfo(124, 45)},{68,new FontSizeInfo(132, 49)},{72,new FontSizeInfo(139, 51)},{96,new FontSizeInfo(184, 68)},{128,new FontSizeInfo(243, 91)},{256,new FontSizeInfo(421, 181)},}},
            {"Impact",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(18, 6)},{9,new FontSizeInfo(18, 7)},{10,new FontSizeInfo(19, 7)},{11,new FontSizeInfo(21, 8)},{12,new FontSizeInfo(22, 9)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(24, 10)},{15,new FontSizeInfo(27, 11)},{16,new FontSizeInfo(28, 11)},{17,new FontSizeInfo(29, 12)},{18,new FontSizeInfo(30, 14)},{20,new FontSizeInfo(36, 16)},{22,new FontSizeInfo(38, 17)},{24,new FontSizeInfo(40, 18)},{26,new FontSizeInfo(45, 20)},{28,new FontSizeInfo(46, 21)},{30,new FontSizeInfo(49, 23)},{32,new FontSizeInfo(55, 24)},{34,new FontSizeInfo(59, 25)},{36,new FontSizeInfo(63, 27)},{38,new FontSizeInfo(65, 29)},{40,new FontSizeInfo(67, 31)},{44,new FontSizeInfo(72, 34)},{48,new FontSizeInfo(83, 37)},{52,new FontSizeInfo(87, 39)},{56,new FontSizeInfo(96, 43)},{60,new FontSizeInfo(100, 45)},{64,new FontSizeInfo(104, 49)},{68,new FontSizeInfo(115, 52)},{72,new FontSizeInfo(119, 55)},{96,new FontSizeInfo(158, 73)},{128,new FontSizeInfo(212, 99)},{256,new FontSizeInfo(420, 196)},}},
            {"Rockwell",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 4)},{4,new FontSizeInfo(9, 4)},{5,new FontSizeInfo(11, 5)},{6,new FontSizeInfo(11, 6)},{7,new FontSizeInfo(12, 6)},{8,new FontSizeInfo(18, 6)},{9,new FontSizeInfo(17, 7)},{10,new FontSizeInfo(19, 7)},{11,new FontSizeInfo(21, 8)},{12,new FontSizeInfo(22, 9)},{13,new FontSizeInfo(23, 9)},{14,new FontSizeInfo(26, 10)},{15,new FontSizeInfo(27, 11)},{16,new FontSizeInfo(29, 11)},{17,new FontSizeInfo(31, 12)},{18,new FontSizeInfo(32, 14)},{20,new FontSizeInfo(36, 16)},{22,new FontSizeInfo(39, 17)},{24,new FontSizeInfo(43, 18)},{26,new FontSizeInfo(46, 20)},{28,new FontSizeInfo(50, 21)},{30,new FontSizeInfo(54, 23)},{32,new FontSizeInfo(57, 24)},{34,new FontSizeInfo(60, 25)},{36,new FontSizeInfo(64, 27)},{38,new FontSizeInfo(67, 29)},{40,new FontSizeInfo(70, 31)},{44,new FontSizeInfo(78, 34)},{48,new FontSizeInfo(85, 37)},{52,new FontSizeInfo(91, 39)},{56,new FontSizeInfo(99, 43)},{60,new FontSizeInfo(106, 45)},{64,new FontSizeInfo(111, 49)},{68,new FontSizeInfo(120, 52)},{72,new FontSizeInfo(127, 55)},{96,new FontSizeInfo(168, 73)},{128,new FontSizeInfo(225, 99)},{256,new FontSizeInfo(420, 196)},}},
            {"Rockwell Condensed",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 4)},{4,new FontSizeInfo(9, 4)},{5,new FontSizeInfo(11, 5)},{6,new FontSizeInfo(11, 6)},{7,new FontSizeInfo(12, 6)},{8,new FontSizeInfo(15, 5)},{9,new FontSizeInfo(16, 5)},{10,new FontSizeInfo(17, 6)},{11,new FontSizeInfo(20, 6)},{12,new FontSizeInfo(21, 7)},{13,new FontSizeInfo(22, 7)},{14,new FontSizeInfo(24, 8)},{15,new FontSizeInfo(26, 9)},{16,new FontSizeInfo(27, 9)},{17,new FontSizeInfo(29, 10)},{18,new FontSizeInfo(31, 10)},{20,new FontSizeInfo(34, 12)},{22,new FontSizeInfo(37, 12)},{24,new FontSizeInfo(41, 15)},{26,new FontSizeInfo(44, 16)},{28,new FontSizeInfo(47, 17)},{30,new FontSizeInfo(50, 18)},{32,new FontSizeInfo(54, 19)},{34,new FontSizeInfo(57, 20)},{36,new FontSizeInfo(60, 21)},{38,new FontSizeInfo(64, 23)},{40,new FontSizeInfo(66, 24)},{44,new FontSizeInfo(74, 26)},{48,new FontSizeInfo(80, 28)},{52,new FontSizeInfo(86, 31)},{56,new FontSizeInfo(94, 34)},{60,new FontSizeInfo(100, 36)},{64,new FontSizeInfo(106, 38)},{68,new FontSizeInfo(113, 41)},{72,new FontSizeInfo(120, 43)},{96,new FontSizeInfo(159, 58)},{128,new FontSizeInfo(212, 77)},{256,new FontSizeInfo(422, 155)},}},
            {"Times New Roman",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 6)},{9,new FontSizeInfo(16, 6)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(20, 8)},{12,new FontSizeInfo(21, 8)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(25, 10)},{15,new FontSizeInfo(26, 10)},{16,new FontSizeInfo(27, 11)},{17,new FontSizeInfo(30, 12)},{18,new FontSizeInfo(31, 12)},{20,new FontSizeInfo(35, 15)},{22,new FontSizeInfo(37, 16)},{24,new FontSizeInfo(41, 17)},{26,new FontSizeInfo(44, 19)},{28,new FontSizeInfo(47, 20)},{30,new FontSizeInfo(51, 21)},{32,new FontSizeInfo(54, 23)},{34,new FontSizeInfo(57, 24)},{36,new FontSizeInfo(61, 25)},{38,new FontSizeInfo(64, 27)},{40,new FontSizeInfo(67, 28)},{44,new FontSizeInfo(76, 32)},{48,new FontSizeInfo(82, 34)},{52,new FontSizeInfo(89, 37)},{56,new FontSizeInfo(96, 40)},{60,new FontSizeInfo(102, 42)},{64,new FontSizeInfo(109, 45)},{68,new FontSizeInfo(116, 49)},{72,new FontSizeInfo(122, 51)},{96,new FontSizeInfo(164, 68)},{128,new FontSizeInfo(216, 91)},{256,new FontSizeInfo(428, 181)},}},
            {"Trebuchet MS",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(18, 6)},{9,new FontSizeInfo(20, 6)},{10,new FontSizeInfo(20, 7)},{11,new FontSizeInfo(22, 8)},{12,new FontSizeInfo(24, 8)},{13,new FontSizeInfo(24, 9)},{14,new FontSizeInfo(25, 10)},{15,new FontSizeInfo(27, 10)},{16,new FontSizeInfo(28, 11)},{17,new FontSizeInfo(30, 12)},{18,new FontSizeInfo(31, 14)},{20,new FontSizeInfo(37, 15)},{22,new FontSizeInfo(38, 16)},{24,new FontSizeInfo(41, 18)},{26,new FontSizeInfo(45, 19)},{28,new FontSizeInfo(48, 20)},{30,new FontSizeInfo(51, 22)},{32,new FontSizeInfo(56, 24)},{34,new FontSizeInfo(58, 25)},{36,new FontSizeInfo(62, 26)},{38,new FontSizeInfo(66, 28)},{40,new FontSizeInfo(68, 29)},{44,new FontSizeInfo(75, 33)},{48,new FontSizeInfo(82, 36)},{52,new FontSizeInfo(89, 38)},{56,new FontSizeInfo(97, 41)},{60,new FontSizeInfo(102, 44)},{64,new FontSizeInfo(108, 48)},{68,new FontSizeInfo(117, 51)},{72,new FontSizeInfo(123, 53)},{96,new FontSizeInfo(164, 71)},{128,new FontSizeInfo(219, 95)},{256,new FontSizeInfo(418, 190)},}},
            {"Tw Cen MT",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(15, 6)},{9,new FontSizeInfo(16, 7)},{10,new FontSizeInfo(17, 7)},{11,new FontSizeInfo(19, 8)},{12,new FontSizeInfo(21, 9)},{13,new FontSizeInfo(22, 9)},{14,new FontSizeInfo(25, 10)},{15,new FontSizeInfo(26, 11)},{16,new FontSizeInfo(27, 12)},{17,new FontSizeInfo(30, 14)},{18,new FontSizeInfo(31, 14)},{20,new FontSizeInfo(34, 16)},{22,new FontSizeInfo(37, 17)},{24,new FontSizeInfo(40, 19)},{26,new FontSizeInfo(44, 20)},{28,new FontSizeInfo(47, 21)},{30,new FontSizeInfo(51, 23)},{32,new FontSizeInfo(54, 25)},{34,new FontSizeInfo(57, 26)},{36,new FontSizeInfo(60, 27)},{38,new FontSizeInfo(64, 29)},{40,new FontSizeInfo(67, 31)},{44,new FontSizeInfo(74, 35)},{48,new FontSizeInfo(80, 37)},{52,new FontSizeInfo(86, 40)},{56,new FontSizeInfo(93, 43)},{60,new FontSizeInfo(100, 46)},{64,new FontSizeInfo(106, 50)},{68,new FontSizeInfo(113, 53)},{72,new FontSizeInfo(119, 56)},{96,new FontSizeInfo(159, 75)},{128,new FontSizeInfo(212, 100)},{256,new FontSizeInfo(421, 199)},}},
            {"Tw Cen MT Condensed",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(14, 4)},{9,new FontSizeInfo(16, 4)},{10,new FontSizeInfo(17, 5)},{11,new FontSizeInfo(19, 5)},{12,new FontSizeInfo(21, 6)},{13,new FontSizeInfo(22, 6)},{14,new FontSizeInfo(25, 7)},{15,new FontSizeInfo(26, 7)},{16,new FontSizeInfo(27, 8)},{17,new FontSizeInfo(29, 8)},{18,new FontSizeInfo(31, 9)},{20,new FontSizeInfo(34, 10)},{22,new FontSizeInfo(37, 11)},{24,new FontSizeInfo(41, 12)},{26,new FontSizeInfo(45, 14)},{28,new FontSizeInfo(48, 15)},{30,new FontSizeInfo(51, 16)},{32,new FontSizeInfo(54, 17)},{34,new FontSizeInfo(57, 17)},{36,new FontSizeInfo(61, 19)},{38,new FontSizeInfo(64, 20)},{40,new FontSizeInfo(67, 20)},{44,new FontSizeInfo(74, 23)},{48,new FontSizeInfo(80, 24)},{52,new FontSizeInfo(87, 26)},{56,new FontSizeInfo(94, 28)},{60,new FontSizeInfo(100, 31)},{64,new FontSizeInfo(107, 33)},{68,new FontSizeInfo(114, 35)},{72,new FontSizeInfo(120, 37)},{96,new FontSizeInfo(160, 50)},{128,new FontSizeInfo(214, 66)},{256,new FontSizeInfo(424, 133)},}},
            {"Verdana",new Dictionary<float,FontSizeInfo>(){{3,new FontSizeInfo(8, 3)},{4,new FontSizeInfo(9, 3)},{5,new FontSizeInfo(11, 4)},{6,new FontSizeInfo(11, 5)},{7,new FontSizeInfo(12, 5)},{8,new FontSizeInfo(14, 7)},{9,new FontSizeInfo(15, 8)},{10,new FontSizeInfo(17, 8)},{11,new FontSizeInfo(19, 10)},{12,new FontSizeInfo(20, 10)},{13,new FontSizeInfo(21, 11)},{14,new FontSizeInfo(24, 12)},{15,new FontSizeInfo(27, 14)},{16,new FontSizeInfo(26, 14)},{17,new FontSizeInfo(29, 16)},{18,new FontSizeInfo(30, 16)},{20,new FontSizeInfo(33, 18)},{22,new FontSizeInfo(36, 19)},{24,new FontSizeInfo(39, 21)},{26,new FontSizeInfo(43, 23)},{28,new FontSizeInfo(47, 25)},{30,new FontSizeInfo(49, 26)},{32,new FontSizeInfo(53, 28)},{34,new FontSizeInfo(55, 31)},{36,new FontSizeInfo(61, 33)},{38,new FontSizeInfo(62, 34)},{40,new FontSizeInfo(67, 36)},{44,new FontSizeInfo(74, 40)},{48,new FontSizeInfo(80, 43)},{52,new FontSizeInfo(86, 46)},{56,new FontSizeInfo(93, 51)},{60,new FontSizeInfo(99, 54)},{64,new FontSizeInfo(103, 57)},{68,new FontSizeInfo(112, 61)},{72,new FontSizeInfo(117, 65)},{96,new FontSizeInfo(156, 86)},{128,new FontSizeInfo(207, 116)},{256,new FontSizeInfo(418, 230)},}},
        };
        public static void LoadAllFontsFromResource()
        {
            LoadFontsFromResource(null);
        }
        public static void LoadFontsFromResource(string fontName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream("OfficeOpenXml.resources.fontsize.zip");

            using (stream)
            {
                var zipStream = new ZipInputStream(stream);
                ZipEntry entry;
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (entry.FileName.Equals("fontsize.bin", StringComparison.OrdinalIgnoreCase))
                    {
                        ReadFontSize(zipStream, (int)entry.UncompressedSize, fontName);
                    }
                }
            }
        }

        private static void ReadFontSize(Stream stream, int uncompressedSize, string fontName)
        {
            var br = new BinaryReader(stream);
            var sp = 0;                
            while (sp < uncompressedSize)
            {
                var length = br.ReadUInt16();
                var nameSize = (short)br.ReadByte();
                var pos = 3;
                pos += nameSize;
                var name = Encoding.ASCII.GetString(br.ReadBytes(nameSize));                    
                if (fontName == null || fontName.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    var fontSize = new Dictionary<float, FontSizeInfo>();
                    while (pos <= length)
                    {
                        var s = br.ReadUInt16();
                        var w = br.ReadUInt16();
                        var h = br.ReadUInt16();
                        fontSize.Add(s / 100F, new FontSizeInfo(h, w));
                        pos += 6;
                    }
                    if (FontHeights.ContainsKey(name)==false)
                    {
                        FontHeights.Add(name, fontSize);
                    }
                    if (fontName != null) break;
                    sp += pos;
                }
                else
                {
                    br.ReadBytes(length - 2);
                    sp += length-2;
                }
            }
        }
    }
}
