using System;
using System.Data.SQLite;
using System.IO;

namespace ExcelParserForOpenCart
{
    public class BaseConnecter : IDisposable
    {
        private readonly SQLiteConnection _connection;
        private readonly bool _isConnected;
        // делегат для обработки ошибок
        private Action<string> _onMsgAction;

        public BaseConnecter(Action<string> onMsgAction)
        {
            _isConnected = true;
            _onMsgAction = onMsgAction;
            var databaseName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "base.sqlite");
            if (!File.Exists(databaseName))
            {
                _onMsgAction("Отсутствует файл базы данных");
				_isConnected = false;
                return;
            }
            _connection = new SQLiteConnection(string.Format("Data Source={0};", databaseName));
            try
            {
                _connection.Open();
            }
            catch
            {
                _isConnected = false;
                _onMsgAction("Ошибка подключения к базе данных");
            }

        }
        /// <summary>
        /// Метод перезаписывающий наименование категории записанную во множественном числе в единственное число
        /// Например: Бамперы силовые -> Бампер силовой
        /// Служит для формировании наименования продукции
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public string OJ_Composition(string source)
        {
            if (_isConnected == false) return source;
            const string commandText = "SELECT old, new FROM oj_rowscomparsion";            
            var myCommand = _connection.CreateCommand();
            myCommand.CommandText = commandText;
            var dataReader = myCommand.ExecuteReader();
            while (dataReader.Read())
            {
                var old = dataReader["old"].ToString();
                old = old.Replace("/n", "").Trim();
                if (!source.Contains(old)) continue;
                var @new = dataReader["new"].ToString();
                return source.Replace(old, @new);
            }
            return source;
        }

        public void Dispose()
        {
            if (_isConnected == false) return;
            _connection.Close();
            _connection.Dispose();
        }
    }
}
