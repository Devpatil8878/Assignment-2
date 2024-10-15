using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "C:\\Users\\TE-65\\Downloads\\Book.xlsx";
        using (var workbook = new XLWorkbook(filePath))
        {
            var sheet1 = workbook.Worksheet(1);
            var sheet2 = workbook.Worksheet(2);

            var headers1 = GetHeaders(sheet1);
            var headers2 = GetHeaders(sheet2);

            var passwordsSheet1 = GetPasswords(sheet1, headers1, "PasswordHash");
            var passwordsSheet2 = GetPasswords(sheet2, headers2, "PasswordHash");

            List<int> list = ComparePasswords(passwordsSheet1, passwordsSheet2);
            int count = 0;
            List<List<string>> result = new List<List<string>>();

            foreach (var item in list)
            {
                GetDataFromRow(sheet1, item);
                result.Add(GetDataFromRow(sheet1, item));
                count++;
                Console.WriteLine(item);
            }

            Console.WriteLine(count);

            CreateNewExcelFile("C:\\Users\\TE-65\\Downloads\\Result.xlsx", result);
        }
    }
    private static void CreateNewExcelFile(string filePath, List<List<string>> records)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.AddWorksheet("Records");

            for (int row = 0; row < records.Count; row++)
            {
                for (int col = 0; col < records[row].Count; col++)
                {
                    worksheet.Cell(row + 1, col + 1).Value = records[row][col];
                }
            }

            workbook.SaveAs(filePath);
            Console.WriteLine($"Excel file created at: {filePath}");
        }
    }

    private static List<string> GetDataFromRow(IXLWorksheet sheet, int rowNumber)
    {
        var rowData = new List<string>();

        var row = sheet.Row(rowNumber);

        foreach (var cell in row.CellsUsed())
        {
            rowData.Add(cell.GetString());
            //Console.Write(cell.GetString());
        }
        //Console.WriteLine();

        return rowData;
    }



    private static Dictionary<string, string> GetHeaders(IXLWorksheet sheet)
    {
        var headers = new Dictionary<string, string>();
        var headerRow = sheet.FirstRow();

        foreach (var cell in headerRow.Cells())
        {
            var columnName = cell.Address.ColumnLetter;
            var headerValue = cell.GetString();
            headers[columnName] = headerValue;
        }

        return headers;
    }

    private static Dictionary<string, string> GetPasswords(IXLWorksheet sheet, Dictionary<string, string> headers, string passwordHeader)
    {
        var passwords = new Dictionary<string, string>();
        var passwordColumnIndex = headers.FirstOrDefault(h => h.Value == passwordHeader).Key;

        foreach (var row in sheet.RowsUsed().Skip(1))
        {
            var passwordCell = row.Cell(passwordColumnIndex);
            var password = passwordCell.GetString();
            var passwordCellReference = passwordCell.Address.ToString();

            if (!string.IsNullOrEmpty(password))
            {
                passwords[password] = passwordCellReference;
            }
        }

        return passwords;
    }


    private static List<int> ComparePasswords(Dictionary<string, string> passwords1, Dictionary<string, string> passwords2)
    {
        var uniqueToSheet1 = passwords1.Keys.Except(passwords2.Keys);
        List<int> list = new List<int>();

        foreach (var password in uniqueToSheet1)
        {
            int temp = int.Parse(passwords1[password].Substring(1));

            list.Add(temp);
        }

        return list;
    }
}
