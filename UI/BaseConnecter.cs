using System;
using System.Data.SQLite;
using System.IO;
using System.Windows;

namespace ExcelParserForOpenCart
{
    public class BaseConnecter : IDisposable
    {
        private readonly SQLiteConnection _connection;
        private readonly bool _isConnected;

        public BaseConnecter()
        {
            _isConnected = true;
            var databaseName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "base.sqlite");
            if (!File.Exists(databaseName))
            {
                MessageBox.Show("Отсутствует файл базы данных");
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
                MessageBox.Show("Ошибка");
            }

        }

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
            _connection.Close();
            _connection.Dispose();
        }
    }
}
