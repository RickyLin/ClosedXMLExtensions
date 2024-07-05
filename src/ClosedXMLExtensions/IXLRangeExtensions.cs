using System;
using ClosedXML.Excel;

namespace ClosedXMLExtensions
{
    public static class IXLRangeExtensions
    {
        public static IXLStyle SetInsideAndOutsideBorders(this IXLRange range, XLBorderStyleValues borderStyle)
            => range.Style.Border.SetInsideBorder(borderStyle).Border.SetOutsideBorder(borderStyle);
    }
}
