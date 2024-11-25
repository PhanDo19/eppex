using changeExcel.Utils;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeOpenXml;

namespace changeExcel.Handler
{
    public class InvoiceHandler
    {
        public List<Invoice> CreateRandomInvoices(List<Product> products, decimal totalAmount, int month, int year)
        {
            List<Invoice> invoices = new List<Invoice>();
            decimal remainingAmount = 0;
            var rangeFrom = totalAmount - 5000000;
            var rangeTo = totalAmount + 5000000;
            Random random = new Random();
            DateTime currentDate = new DateTime(year, month, 1);
            // loop while reaminingAmount in range from totalAmount - 5,000,000 to totalAmount + 5,000,000
            while( remainingAmount >= rangeTo)
            {
                Console.WriteLine($"TotalPrice - {remainingAmount}");
                Invoice invoice = new Invoice
                {
                    InvoiceNumber = "BH" + month.ToString() + random.Next(0, 500).ToString("D3"),
                    InvoiceDate = currentDate.AddDays(random.Next(0, DateTime.DaysInMonth(currentDate.Year, currentDate.Month))).ToString("MM/dd/yyyy"),
                    Items = new List<InvoiceItem>(),
                    TotalAmount = 0
                };

                decimal invoiceTotal = 0;
                int numItems = random.Next(1, 4); // Tạo từ 1 đến 3 mặt hàng trong mỗi đơn hàng
                for (int i = 0; i < numItems; i++)
                {
                    var index = random.Next(0, products.Count - 1);
                    var product = products[index];
                    if (product == null || product.Quantity <= 0)
                    {
                        continue;
                    }

                    InvoiceItem item = new InvoiceItem
                    {
                        Product = products[index],
                        Quantity = GetQuantity(product.Quantity)
                    };
                    //check null before add
                    if (item.Product != null && invoice != null)
                    {   
                        invoice.Items.Add(item);
                        product.Quantity -= item.Quantity;
                    }
                    invoiceTotal += item.Product.Price * item.Quantity;
                }

                if (invoiceTotal > 0)
                {
                    invoice.TotalAmount = invoiceTotal;

                    remainingAmount = remainingAmount + invoiceTotal;

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
                    worksheet.Cells["A1"].Value = "STT";
                    worksheet.Cells["B1"].Value = "IDChungTu/MaBill";
                    worksheet.Cells["C1"].Value = "TenHangHoaDichVu";
                    worksheet.Cells["D1"].Value = "DonViTinh/ChietKhau";
                    worksheet.Cells["E1"].Value = "SoLuong";
                    worksheet.Cells["F1"].Value = "DonGia";
                    worksheet.Cells["G1"].Value = "ThanhTien";
                    worksheet.Cells["H1"].Value = "ThueSuat";
                    worksheet.Cells["I1"].Value = "TienThueGTGT";
                    worksheet.Cells["J1"].Value = "NgayThangNamHD";
                    worksheet.Cells["K1"].Value = "HoTenNguoiMua";
                    worksheet.Cells["L1"].Value = "TenDonVi";
                    worksheet.Cells["M1"].Value = "MaSoThue";
                    worksheet.Cells["N1"].Value = "DiaChi";
                    worksheet.Cells["O1"].Value = "SoTaiKhoan";
                    worksheet.Cells["P1"].Value = "HinhThucTT";
                    worksheet.Cells["Q1"].Value = "NhanBangEmail";
                    worksheet.Cells["R1"].Value = "DSEmail";
                    worksheet.Cells["S1"].Value = "NhanBangSMS";
                    worksheet.Cells["T1"].Value = "DSSMS";
                    worksheet.Cells["U1"].Value = "NhanBangBanIN";
                    worksheet.Cells["V1"].Value = "HoTenNguoiNhan";
                    worksheet.Cells["W1"].Value = "SoDienThoaiNguoiNhan";
                    worksheet.Cells["X1"].Value = "SoNha";
                    worksheet.Cells["Y1"].Value = "Tinh/ThanhPho";
                    worksheet.Cells["Z1"].Value = "Huyen/Quan/ThiXa";
                    worksheet.Cells["AA1"].Value = "Xa/Phuong/ThiTran";
                    worksheet.Cells["AB1"].Value = "GhiChu";
                    
                    // Write data to Excel file
                    int row = 2;
                    int stt = 1;
                    foreach (var invoice in originalInvoices)
                    {
                        foreach (var item in invoice.Items)
                        {
                            worksheet.Cells[$"A{row}"].Value = stt++;
                            worksheet.Cells[$"B{row}"].Value = invoice.InvoiceNumber;
                            worksheet.Cells[$"C{row}"].Value = item.Product.Name;
                            worksheet.Cells[$"D{row}"].Value = item.Product.Unit;
                            worksheet.Cells[$"E{row}"].Value = item.Quantity;
                            worksheet.Cells[$"F{row}"].Value = item.Product.Price;
                            worksheet.Cells[$"G{row}"].Value = item.Product.Price * item.Quantity;
                            worksheet.Cells[$"H{row}"].Value = item.Product.TaxRate;
                            worksheet.Cells[$"I{row}"].Value = item.Product.Price * item.Quantity * item.Product.TaxRate / 100;
                            worksheet.Cells[$"J{row}"].Value = invoice.InvoiceDate;
                            worksheet.Cells[$"K{row}"].Value = "Người mua không lấy hóa đơn";
                            worksheet.Cells[$"P{row}"].Value = "TM/CK";
                        }
                    }

                    // Điều chỉnh độ rộng cột
                    worksheet.Cells.AutoFitColumns();   

                    // Lưu file Excel
                    excelPackage.SaveAs(existingFileInfo);
                }
            }
        }

        private static int GetQuantity(int quantity)
        {
            if(quantity == 1 || quantity == 2)
                return quantity;
            if ( quantity > 2 && quantity < 10)
                return new Random().Next(2, 5);
            if (quantity >= 10)
                return new Random().Next(7, 15);
            return 0;
        }
    }
}
