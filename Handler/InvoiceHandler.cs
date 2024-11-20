using changeExcel.Utils;
using OfficeOpenXml;

namespace changeExcel.Handler
{
    public class InvoiceHandler
    {
        public List<Invoice> CreateRandomInvoices(List<Product> products, decimal totalAmount, int month, int year)
        {
            List<Invoice> invoices = new List<Invoice>();
            decimal remainingAmount = totalAmount;
            Random random = new Random();
            DateTime currentDate = new DateTime(year, month, 1);
            // areadly exist remainingAmount = 0 for break loop
            while (remainingAmount > 0)
            {
                Invoice invoice = new Invoice
                {
                    InvoiceNumber = "BH" + random.Next(10000, 100000).ToString("D5"),
                    InvoiceDate = currentDate.AddDays(random.Next(0, DateTime.DaysInMonth(currentDate.Year, currentDate.Month))).ToString("MM/dd/yyyy"),
                    Items = new List<InvoiceItem>(),
                    TotalAmount = 0
                };

                decimal invoiceTotal = 0;
                int numItems = random.Next(1, 4); // Tạo từ 1 đến 3 mặt hàng trong mỗi đơn hàng
                for (int i = 0; i < numItems; i++)
                {
                    InvoiceItem item = new InvoiceItem
                    {
                        Product = products[random.Next(0, products.Count)],
                        Quantity = random.Next(1, 3) // Số lượng từ 1 đến 2
                    };
                    //check null before add
                    if (item.Product != null && invoice != null)
                    {   
                        invoice.Items.Add(item);
                    }
                    invoiceTotal += item.Product.Price * item.Quantity;
                }

                if (invoiceTotal <= remainingAmount)
                {
                    invoice.TotalAmount = invoiceTotal;
                    
                    remainingAmount = remainingAmount - invoiceTotal;
                    //If the remaining amount <0 then call the function again and find the invoice with the amount closest to the remaining amount.
                    if (remainingAmount < 0)
                    {
                        invoices.Add(invoice);
                        Console.WriteLine("Kiểm tra lại hóa đơn cuối cùng do só tiền lớn hơn đầu vào");
                        break;
                    }
                    invoices.Add(invoice);
                }
            }

            return invoices;
        }


        public void SaveInvoicesToExcel(List<Invoice> originalInvoices, string filePath, string sheetName)
        {
            FileInfo existingFileInfo = new FileInfo($@"{filePath}");

            if (existingFileInfo.Exists)
            {
                using (ExcelPackage excelPackage = new ExcelPackage(existingFileInfo))
                {

                    var worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);

                    // Định dạng header
                    worksheet.Cells["A1:F1"].Style.Font.Bold = true;
                    worksheet.Cells["A1"].Value = "Mã Chứng Từ";
                    worksheet.Cells["B1"].Value = "Tên Hàng";
                    worksheet.Cells["C1"].Value = "Số Lượng";
                    worksheet.Cells["D1"].Value = "Đơn Giá";
                    worksheet.Cells["E1"].Value = "Thành Tiền";
                    worksheet.Cells["F1"].Value = "Ngày Tháng";

                    // Nhóm các hóa đơn theo ngày
                    var invoicesByDay = originalInvoices.GroupBy(i => Convert.ToDateTime(i.InvoiceDate).Date);

                    // Lưu các hóa đơn vào worksheet
                    int row = 2;
                    foreach (var invoiceGroup in invoicesByDay)
                    {
                        // Tạo số lượng hóa đơn ngẫu nhiên trong ngày
                        int numInvoices = new Random().Next(1, invoiceGroup.Count() + 1);
                        var invoicesForDay = invoiceGroup.OrderBy(x => Guid.NewGuid()).Take(numInvoices).ToList();

                        foreach (var invoice in invoicesForDay)
                        {
                            foreach (var item in invoice.Items)
                            {
                                worksheet.Cells[row, 1].Value = invoice.InvoiceNumber;
                                worksheet.Cells[row, 2].Value = item.Product.Name;
                                worksheet.Cells[row, 3].Value = item.Quantity;
                                worksheet.Cells[row, 4].Value = item.Product.Price;
                                worksheet.Cells[row, 5].Value = item.Quantity * item.Product.Price;
                                worksheet.Cells[row, 6].Value = invoice.InvoiceDate;
                                row++;
                            }
                        }
                    }

                    // Điều chỉnh độ rộng cột
                    worksheet.Cells.AutoFitColumns();

                    // Lưu file Excel
                    excelPackage.SaveAs(existingFileInfo);
                }
            }
        }
    }
}
