using OfficeOpenXml;
using System.Collections.Generic;

namespace WebApplication2.Extensions
{
    public static class WorksheetExtension
    {
        public static ExcelWorksheets GenerateWorksheet<T>(this ExcelWorksheets worksheet, List<T> data, string nameOfWorksheet)
        {
            if (data.Count != 0)
            {
                worksheet.Add(nameOfWorksheet).Cells.LoadFromCollection(data, true);
                return worksheet;
            }

            return null;
        }
    }
}
