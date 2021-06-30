using OfficeOpenXml;
using System.Collections.Generic;

namespace WebApplication2.Extensions
{
    public static class WorksheetExtension
    {
        public static ExcelWorksheet GenerateWorksheet<T>(this ExcelWorksheet worksheet, List<T> data)
        {
            worksheet.Cells.LoadFromCollection(data, true);
            return worksheet;
        }
    }
}
