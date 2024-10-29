
using changeExcel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Data;
using System.Reflection;
using System.Text;

var currentDate = DateTime.Now;
string path = string.Empty;
string fileName = string.Empty;
int inputIdSheet = 0;
// Thiết lập mã hóa UTF-8 cho console
Console.OutputEncoding = Encoding.UTF8;
Console.InputEncoding = Encoding.UTF8;

Console.ForegroundColor = ConsoleColor.Magenta;
Console.WriteLine($"{Environment.NewLine}Hello Oanh Hoàng <3, bây giờ là {currentDate:d} - {currentDate:t}!");
Console.WriteLine("Em làm rồi nghỉ sớm nha");

//while (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(fileName))
//{
//    Console.ForegroundColor = ConsoleColor.Blue;
//    Console.WriteLine("Nhập thư mục chứa file, ví dụ E:\\zalo ");
//    path = Console.ReadLine();
//    Console.WriteLine("Nhập tên file");
//    fileName = Console.ReadLine();
//    Console.WriteLine("Nhập id sheet dùng để lấy data tình từ sheet 0");
//    inputIdSheet = Convert.ToInt32(Console.ReadLine());

//    if (!string.IsNullOrEmpty(path) && !string.IsNullOrEmpty(fileName))
//    {
//        Console.ForegroundColor = ConsoleColor.Red;
//        Console.WriteLine("Thiếu tên hoặc đường dẫn (:");
//    }

//    if (inputIdSheet < 0)
//    {
//        Console.ForegroundColor = ConsoleColor.Red;
//        Console.WriteLine("Id sheet không hợp lệ, vui lòng nhập lại");
//    }

//    //check format file
//    if (!fileName.EndsWith(".xlsx"))
//    {
//        Console.ForegroundColor = ConsoleColor.Red;
//        Console.WriteLine("File không đúng định dạng, vui lòng nhập lại");
//    }
//    // check file exist
//    if (!File.Exists(System.IO.Path.Combine(path, fileName)))
//    {
//        Console.ForegroundColor = ConsoleColor.Red;
//        Console.WriteLine("File không tồn tại, vui lòng nhập lại");
//    }
//}

//OfficeOpenXml.LicenseException
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var link = "E:\\zalo";
var name = "thang8.xlsx";
var filePath = System.IO.Path.Combine(link, name);

var data = ReadExcelFile(filePath, inputIdSheet);

Console.ForegroundColor = ConsoleColor.White;
Console.Write($"{Environment.NewLine}(:\t...Đợi tý sắp ra rồi... \t(:");
Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine("\n Đang chạy dữ liệu rồi nha \n");


//write data to file
Console.WriteLine("Nhập tên sheet mới đi em ");
//string sheetName = Console.ReadLine();
string sheetName = "ok";
if (!string.IsNullOrEmpty(sheetName))
{
   Console.WriteLine("không nhập tên, tạm thời để Sheet1 nha");
   sheetName = "Sheet1";
}



WriteExcelFile(filePath, data, sheetName);

for (int i = 0; i < 100; i++)
{
    Console.Write("\u2593");
    Thread.Sleep(50);
}

Console.ReadKey(true);

void WriteExcelFile(string filePath, List<RootData> data, string sheetName)
{
    using (var package = new ExcelPackage())
    {
        var worksheet = package.Workbook.Worksheets.Add(sheetName);

        for (int i = 0; i < data.Count; i++)
        {
            worksheet.Cells[i + 1, 1].Value = data[i];
        }

        File.WriteAllBytes(filePath, package.GetAsByteArray());
    }
}

List<RootData> ReadExcelFile(string filePath, int inputIdSheet)
{
    var data = new List<RootData>();
    inputIdSheet = 2;
    try
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[inputIdSheet];
            if (worksheet != null)
            {
                data = worksheet.ReadExcelToList<RootData>();
            }
        }
        //using (ExcelPackage package = new ExcelPackage(file.InputStream))
        //{
        //    ExcelWorkbook workbook = package.Workbook;
        //    if (workbook != null)
        //    {
        //        ExcelWorksheet worksheet = workbook.Worksheets.FirstOrDefault();
        //        if (worksheet != null)
        //        {
        //            list = worksheet.ReadExcelToList<Users>();
        //            //Your code
        //        }
        //    }
        //}
    }
    catch (Exception ex)
    {
        //Save error log
    }
    return data;
}

public static class ReadExcel
{
    public static List<T> ReadExcelToList<T>(this ExcelWorksheet worksheet) where T : new()
    {
        List<T> collection = new List<T>();
        try
        {
            DataTable dt = new DataTable();
            foreach (var firstRowCell in new T().GetType().GetProperties().ToList())
            {
                //Add table colums with properties of T
                dt.Columns.Add(firstRowCell.Name);
            }
            for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                DataRow row = dt.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }

            //Get the colums of table
            var columnNames = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

            //Get the properties of T
            List<PropertyInfo> properties = new T().GetType().GetProperties().ToList();

            collection = dt.AsEnumerable().Select(row =>
            {
                T item = Activator.CreateInstance<T>();
                foreach (var pro in properties)
                {
                    if (columnNames.Contains(pro.Name) || columnNames.Contains(pro.Name.ToUpper()))
                    {
                        PropertyInfo pI = item.GetType().GetProperty(pro.Name);
                        pro.SetValue(item, (row[pro.Name] == DBNull.Value) ? null : Convert.ChangeType(row[pro.Name], (Nullable.GetUnderlyingType(pI.PropertyType) == null) ? pI.PropertyType : Type.GetType(pI.PropertyType.GenericTypeArguments[0].FullName)));
                    }
                }
                return item;
            }).ToList();

        }
        catch (Exception ex)
        {
            //Save error log
        }

        return collection;
    }
}