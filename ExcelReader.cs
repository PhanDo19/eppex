using changeExcel.Utils;
using OfficeOpenXml;

namespace changeExcel
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

        public List<RootData> ReadDataFromExcel()
        {
            try
            {
                FileInfo existingFileInfo = new FileInfo($@"{FilePath}");

                if (existingFileInfo.Exists)
                {
                    using (ExcelPackage excelPackage = new ExcelPackage(existingFileInfo))
                    {
                        ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[SheetIndex];

                        return excelWorksheet.ConvertSheetToObjects<RootData>().ToList();
                    }
                }

                Console.WriteLine("\n\nLoad fail");

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("+++++ ERROR - Hiba!" + $" Hiba az Excel beolvasása közben!\n\nException Message {ex.Message}");

                return null;
            }
        }

        public void PrintData(List<RootData> data)
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
                    Console.WriteLine($"GrossProfit: {item.GrossProfit}");
                    Console.WriteLine($"ProfitRate: {item.ProfitRate}");
                    Console.WriteLine();
                }
            }
        }   
    }
}