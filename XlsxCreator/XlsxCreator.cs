using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxCreator
{
    public class XlsxCreator
    {
        public byte[] GetFileAsByteArray(IDataPreparator preparator, bool withHeaders)
        {
            using (var xlsxPackage = new ExcelPackage())
            {
                var worksheet = xlsxPackage.Workbook.Worksheets.Add("Report");
                IEnumerable<IEnumerable<object>> data;
                if (withHeaders)
                {
                    data = preparator.PrepareWithHeaders();
                    CreateHeaders(worksheet, data.First());
                    CreateContent(worksheet, data, withHeaders);
                }
                else
                {
                    data = preparator.Prepare();
                    CreateContent(worksheet, data, withHeaders);
                }

                worksheet.Cells.AutoFitColumns();
                return xlsxPackage.GetAsByteArray();
            }
        }

        private void CreateHeaders(ExcelWorksheet worksheet, IEnumerable<object> headers)
        {
            const int headerRow = 1;
            int column = 1;

            foreach (var header in headers)
            {
                worksheet.Cells[headerRow, column++].Value = header;
            }
        }

        private void CreateContent(ExcelWorksheet worksheet, IEnumerable<IEnumerable<object>> allData, bool withHeaders)
        {
            int row = 1;
            int column = 1;

            if (withHeaders)
            {
                allData = allData.Skip(1);
                row = 2;
            }

            foreach (var data in allData)
            {
                foreach (var rowInfo in data)
                {
                    TryFormatDateCell(worksheet, row, column, rowInfo);
                    worksheet.Cells[row, column++].Value = rowInfo;
                }
                column = 1;
                row++;
            }
        }

        private static void TryFormatDateCell(ExcelWorksheet worksheet, int row, int column, object rowInfo)
        {
            if (rowInfo is DateTime)
                worksheet.Cells[row, column].Style.Numberformat.Format = "yyyy-mm-dd";
        }
    }
}
