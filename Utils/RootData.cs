namespace changeExcel.Utils
{
    public class RootData
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
        public double Price { get; set; }

        [Column(6)]
        public double SalePrice { get; set; }
        //thuế suất
        [Column(7)]
        public double TaxRate { get; set; }
        //đơn giá check
        [Column(8)]
        public double PriceCheck { get; set; }
        // thuế check 
        [Column(9)]
        public bool TaxCheck { get; set; }

        //lai gộp
        [Column(10)]
        public double GrossProfit { get; set; }
        // tỷ lệ lãi
        [Column(11)]
        public double ProfitRate { get; set; }
    }
}
