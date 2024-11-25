using changeExcel.Utils;
using OfficeOpenXml;

namespace changeExcel.Handler
{
    public class ExcelReader
    {
        public string FilePath;
        public int SheetIndex;

        public ExcelReader(string path, int sheetIndex = 0)
        {
            FilePath = path;
            SheetIndex = sheetIndex;
        }

        public List<Product> ReadDataFromExcel()
        {
            try
            {
                FileInfo existingFileInfo = new FileInfo($@"{FilePath}");

                if (existingFileInfo.Exists)
                {
                    using ExcelPackage excelPackage = new ExcelPackage(existingFileInfo);
                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[SheetIndex];

                    return excelWorksheet.ConvertSheetToObjects<Product>().ToList();
                }

                Console.WriteLine("\n\nLoad fail");

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("+++++ ERROR - Hiba!" + $" Lỗi code \n\nException Message {ex.Message}");
                return null;
            }
        }

        public void PrintData(List<Product> data)
        {
            if (data != null)
            {
                foreach (var item in data)
                {
                    Console.WriteLine($"Code: {item.Code}");
                    Console.WriteLine($"Name: {item.Name}");
                    Console.WriteLine($"Unit: {item.Unit}");
                    Console.WriteLine($"Quantity: {item.Quantity}");
                    Console.WriteLine($"Price: {item.Price}");
                    Console.WriteLine($"SalePrice: {item.SalePrice}");
                    Console.WriteLine($"TaxRate: {item.TaxRate}");
                    Console.WriteLine($"PriceCheck: {item.PriceCheck}");
                    Console.WriteLine($"TaxCheck: {item.TaxCheck}");
                    Console.WriteLine();
                }
            }
        }
    }
}