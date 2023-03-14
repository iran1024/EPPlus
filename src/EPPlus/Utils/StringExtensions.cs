using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace OfficeOpenXml.Utils
{
    internal static class StringExtensions
    {
        internal static string NullIfWhiteSpace(this string s) { return s == "" ? null : s; }

        internal static string GetSubstringStoppingAtSymbol(this string s, int index, string stopSymbol = "\"") 
        {
            if(!string.IsNullOrEmpty(s))
            {
                int charIndex = s.IndexOf(stopSymbol, index);

                if(charIndex > 0)
                {
                    return s.Substring(index, charIndex);
                }
            }

            return string.Empty;
        }
    }
}
