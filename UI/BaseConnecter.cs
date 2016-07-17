using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelParserForOpenCart
{
    public class BaseConnecter
    {
        public BaseConnecter()
        {
            var databaseName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "base.sqlite");
            if (!File.Exists(databaseName))
            {
                MessageBox.Show("Отсутствует файл базы данных");
                return;
            }
            var connection = new SQLiteConnection(string.Format("Data Source={0};", databaseName));

            const string commandText = "SELECT old, new FROM rowscomparsion";
            try
            {
                connection.Open();
                var myCommand = connection.CreateCommand();
                myCommand.CommandText = commandText;
                var dataReader = myCommand.ExecuteReader();
                while (dataReader.Read())
                {
                    var str1 = dataReader["old"].ToString();
                    var str2 = dataReader["new"].ToString();
                }
                connection.Close();
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }

        }
    }
}
