using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

class User
{
    public string UserName { get; set; }
    public string PasswordHash { get; set; }
    public int CellReference { get; set; }

    public User(string userName, string passwordHash, int cellReference)
    {
        UserName = userName;
        PasswordHash = passwordHash;
        CellReference = cellReference;
    }

}

class Program
{
    static void Main(string[] args)
    {
        string filePath = "C:\\Users\\TE-65\\Downloads\\Book.xlsx";
        using (var workbook = new XLWorkbook(filePath))
        {
            List<User> users = new List<User>();

            var sheet1 = workbook.Worksheet(1);
            var sheet2 = workbook.Worksheet(2);

            var headers1 = GetHeaders(sheet1);
            var headers2 = GetHeaders(sheet2);

            var usersSheet1 = GetPasswords(sheet1, headers1, "PasswordHash");
            var usersSheet2 = GetPasswords(sheet2, headers2, "PasswordHash");

            List<User> list = ComparePasswords(usersSheet1, usersSheet2);
            int count = 0;
            List<List<string>> result = new List<List<string>>();

            foreach (var item in list)
            {
                GetDataFromRow(sheet1, item.CellReference);
                result.Add(GetDataFromRow(sheet1, item.CellReference));
                count++;
                Console.WriteLine(item.UserName);
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

    private static List<User> GetPasswords(IXLWorksheet sheet, Dictionary<string, string> headers, string passwordHeader)
    {
        //var passwords = new Dictionary<string, string>();
        List<User> users = new List<User>();
        var passwordColumnIndex = headers.FirstOrDefault(h => h.Value == passwordHeader).Key;
        var UsernameColumnIndex = headers.FirstOrDefault(h => h.Value == "UserName").Key;

        foreach (var row in sheet.RowsUsed().Skip(1))
        {
            var passwordCell = row.Cell(passwordColumnIndex);
            var password = passwordCell.GetString();
            var passwordCellReference = passwordCell.Address.ToString();
            
            var usernameCell = row.Cell(UsernameColumnIndex);
            var useranme = usernameCell.GetString();
            var usernameCellReference = usernameCell.Address.ToString();

            if (!string.IsNullOrEmpty(password))
            {
                //passwords[password] = passwordCellReference;
                users.Add(new User(useranme, password, int.Parse(passwordCellReference.Substring(1))));
            }
        }

        return users;
    }


    private static List<User> ComparePasswords(List<User> users1, List<User> users2)
    {
        //var uniqueToSheet1 = passwords1.Keys.Except(passwords2.Keys);
        List<User> passChanged = new List<User>();

        foreach(var user in users1)
        {
            foreach(var user2 in users2)
            {
                if (user.UserName == user2.UserName && user.PasswordHash != user2.PasswordHash)
                    passChanged.Add(user);
            }
                
        }


        return passChanged;
    }
}
