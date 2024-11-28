using changeExcel.Handler;
using DocumentFormat.OpenXml.Vml;
using OfficeOpenXml;
using System.Text;

var currentDate = DateTime.Now;
string path = string.Empty;
string fileName = string.Empty;
int inputIdSheet = 0;
var invcHandler = new InvoiceHandler();
Console.OutputEncoding = Encoding.UTF8;
Console.InputEncoding = Encoding.UTF8;

Console.ForegroundColor = ConsoleColor.Magenta;
Console.WriteLine($"{Environment.NewLine}Hello Oanh Hoàng <3, bây giờ là {currentDate:d} - {currentDate:t}!");
//Console.WriteLine("Em làm rồi nghỉ sớm nha");

while (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(fileName))
{
    Console.ForegroundColor = ConsoleColor.Blue;
    Console.WriteLine("Nhập thư mục chứa file, ví dụ E:\\zalo ");
    path = Console.ReadLine();
    Console.WriteLine("Nhập tên file");
    fileName = Console.ReadLine();
    Console.WriteLine("Nhập id sheet dùng để lấy data tình từ sheet 0");
    inputIdSheet = Convert.ToInt32(Console.ReadLine());

    if (string.IsNullOrEmpty(path) && string.IsNullOrEmpty(fileName))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Thiếu tên hoặc đường dẫn (:");
    }

    if (inputIdSheet < 0)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Id sheet không hợp lệ, vui lòng nhập lại");
    }

    //check format file
    if (!fileName.EndsWith(".xlsx"))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("File không đúng định dạng, vui lòng nhập lại");
    }
    // check file exist
    if (!File.Exists(System.IO.Path.Combine(path, fileName)))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("File không tồn tại, vui lòng nhập lại");
    }
}

//OfficeOpenXml.LicenseException
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var filePath = System.IO.Path.Combine(path, fileName);

var excelReader = new ExcelReader(filePath, inputIdSheet);

var data = excelReader.ReadDataFromExcel();
var totalValue = data.Sum(x => x.Value * x.Quantity);

if (data != null)
{
    Console.ForegroundColor = ConsoleColor.White;
    Console.Write($"{Environment.NewLine}(:\t...Đợi tý sắp ra rồi... \t(:");
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("\n Dữ liệu đã được đọc từ file excel.\n Bây h nhập số tiền mong đợi trong tháng, tháng và năm");
    Console.WriteLine("\n ---------- Tổng tiền ------------");
    var total = Convert.ToDecimal(Console.ReadLine());

    while (totalValue < total) {

        Console.WriteLine("\n ---------- Nhập lại tổng tiền ------------");
        total = Convert.ToDecimal(Console.ReadLine());
    }
 
    Console.WriteLine("\n ---------- Tháng ------------");
    var month = Convert.ToInt32(Console.ReadLine());
    Console.WriteLine("\n ---------- Năm ------------");
    var year = Convert.ToInt32(Console.ReadLine());


    var invoices = invcHandler.CreateRandomInvoices(data, total == 0? Convert.ToDecimal("10000000"): total,
        1<= month && month>= 12? month: DateTime.Now.Month, year == 0? DateTime.Now.Year: year);

    Console.WriteLine("Nhập tên sheet mới đi em ");
    string sheetName = Console.ReadLine();
    if (string.IsNullOrEmpty(sheetName))
    {
        Console.WriteLine("không nhập tên, tạm thời để Sheet1 nha");
        sheetName = "Sheet1";
    }
    for (int i = 0; i < 30; i++)
    {
        Console.Write("\u2593");
        Thread.Sleep(50);
    }

    invcHandler.SaveInvoicesToExcel(invoices, filePath, sheetName);
}
else
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Không có dữ liệu nào được đọc từ file excel");
}


Console.WriteLine("Nhập bất kỳ để đóng");
Console.ReadKey(true);