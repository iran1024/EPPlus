﻿using System;
using System.Xml;

namespace OfficeOpenXml.Utils
{
    internal static class XmlReaderHelper
    {
        internal static bool ReadUntil(this XmlReader xr, int depth, params string[] tagName)
        {
            if (xr.EOF) return false;
            while ((xr.Depth == depth && Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag))) == false)
            {
                do
                {
                    xr.Read();
                    if (xr.EOF) return false;
                } while (xr.Depth != depth);
            }
            return (ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]));
        }

        internal static bool ReadUntil(this XmlReader xr, params string[] tagName)
        {
            if (xr.EOF) return false;
            do
            {
                if (xr.EOF) return false;
                xr.Read();
            } while ((Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag))) == false);

            return (ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]));
        }
    }
}