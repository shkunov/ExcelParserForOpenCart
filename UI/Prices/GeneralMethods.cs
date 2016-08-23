using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public abstract class GeneralMethods
    {
        protected readonly BackgroundWorker Worker;
        protected readonly DoWorkEventArgs E;

        public List<OutputPriceLine> ResultingPrice { get; private set; }

        protected GeneralMethods(object sender, DoWorkEventArgs e)
        {
            ResultingPrice = new List<OutputPriceLine>();
            Worker = sender as BackgroundWorker;
            E = e;
        }

        protected GeneralMethods()
        {
        }

        protected static string ConverterToString(dynamic obj)
        {
            string s;
            try
            {
                s = Convert.ToString(obj);
            }
            catch
            {
                s = string.Empty;
            }
            return s;
        }

        protected static decimal ConverterToDecimal(Range range)
        {
            if (range == null)
                return 0;
            var obj = range.Value2;
            if (obj == null)
                return 0;
            decimal d;
            try
            {
                d = Convert.ToDecimal(obj);
            }
            catch
            {
                d = 0;
            }
            return d;
        }

        protected static string ConverterToString(Range range)
        {
            if (range == null)
                return string.Empty;
            var obj = range.Value2;
            if (obj == null)
                return string.Empty;
            string s;
            try
            {
                s = Convert.ToString(obj);
            }
            catch
            {
                s = string.Empty;
            }
            return s;
        }
    }
}
