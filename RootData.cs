namespace changeExcel
{
    public class RootData
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
        public double SalePrice { get; set; }
        //thuế suất
        public double TaxRate { get; set; }
        //đơn giá check 
        public double PriceCheck { get; set; }
        // thuế check 
        public bool TaxCheck { get; set; }

        //lai gộp
        public double GrossProfit { get; set; }
        // tỷ lệ lãi
        public double ProfitRate { get; set; }
    }
}
