﻿using changeExcel;
using OfficeOpenXml;
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
var link = "C:\\Users\\dopt\\SETUP\\ePPlus\\eppex\\aceess\\";
var name = "thang8.xlsx";
var filePath = System.IO.Path.Combine(link, name);

var excelReader = new ExcelReader(filePath, 2);
var data = excelReader.ReadDataFromExcel();
excelReader.PrintData(data);

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

for (int i = 0; i < 100; i++)
{
    Console.Write("\u2593");
    Thread.Sleep(50);
}

Console.ReadKey(true);