using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    public class GeneralMethods
    {
        protected BackgroundWorker Worker;
        protected DoWorkEventArgs E;

        public List<OutputPriceLine> ResultingList { get; private set; }

        public GeneralMethods()
        {
            ResultingList = new List<OutputPriceLine>();
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
