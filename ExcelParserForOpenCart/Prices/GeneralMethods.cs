using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace ExcelParserForOpenCart.Prices
{
    /// <summary>
    /// Общий абстрактный класс, который содержит:
    /// 1. BackgroundWorker в котором происходит обработка прайса
    /// 2. Методы для конвертации данных из ячеек
    /// </summary>
    public abstract class GeneralMethods
    {
        protected readonly BackgroundWorker Worker;
        protected readonly DoWorkEventArgs E;
        public List<Producers> Producers { get; private set; }

        public List<OutputPriceLine> ResultingPrice { get; private set; }

        public event Action<string> OnMsg;

        protected GeneralMethods(object sender, DoWorkEventArgs e)
        {
            ResultingPrice = new List<OutputPriceLine>();
            Producers = new List<Producers>();
            using (var baseConnecter = new BaseConnecter(OnBaseMsgAction))
            {
                Producers.AddRange(baseConnecter.GetProducers());
            }
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

        private void OnBaseMsgAction(string s)
        {
            if (OnMsg != null) OnMsg(s);
        }
    }
}
