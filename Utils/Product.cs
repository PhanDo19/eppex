namespace changeExcel.Utils
{

    public class Invoice
    {
        public List<InvoiceItem> Items { get; set; }
        public decimal TotalAmount { get; set; }
        public string InvoiceNumber { get; set; }
        public string InvoiceDate { get; set; }
    }

    public class InvoiceItem
    {
        public Product Product { get; set; }
        public int Quantity { get; set; }
    }
    public class Product
    {
        [Column(1, true)]
        public string Code { get; set; }
        [Column(2)]
        public string Name { get; set; }
        [Column(3)] 
        public string Unit { get; set; }
        [Column(4)]
        public int Quantity { get; set; }

        [Column(5)]
        public decimal Price { get; set; }

        // Giá trị
        [Column(6)]
        public decimal Value { get; set; }

        //mặt hàng
        [Column(7)]
        public string Item { get; set; }

        [Column(8)]
        public decimal SalePrice { get; set; }
        //thuế suất
        [Column(9)]
        public decimal TaxRate { get; set; }
        //đơn giá check
        [Column(10)]
        public decimal PriceCheck { get; set; }
        // thuế check 
        [Column(11)]
        public bool TaxCheck { get; set; }

    }
}
