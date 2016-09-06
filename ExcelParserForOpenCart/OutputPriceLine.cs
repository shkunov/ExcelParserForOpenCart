namespace ExcelParserForOpenCart
{
    public class OutputPriceLine
    {
        /// <summary>
        /// Артикул
        /// </summary>
        public string VendorCode { get; set; }
        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; set; }
        public string Category1 { get; set; }
        public string Category2 { get; set; }
        /// <summary>
        /// Производитель
        /// </summary>
        public string Producer { get; set; }
        public string ProductDescription { get; set; }
        public string Cost { get; set; } // а не сделать ли числовые поля числовыми, int?

        public string Foto { get; set; } // возможно следует использовать другой тип
        public string Option { get; set; } // Опция Кронштейн веткоотсекателя
        /// <summary>
        /// Количество
        /// </summary>
        public string Qt { get; set; }
        /// <summary>
        /// Плюс с цене
        /// </summary>
        public string PlusThePrice { get; set; }
    }
}
