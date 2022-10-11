using MySql.Data.MySqlClient;
using System;

namespace SqlConn
{
    class SQLScripts
    {
        public static void DeleteAllFromTable()
        {
            MySqlConnection conn = DBUtils.GetDBConnection();
            try
            {
                conn.Open();
                string sql = "TRUNCATE TABLE matrixArray;"; //удаляет все записи и заодно обнуляет автоинкремент id
                MySqlCommand cmd = new MySqlCommand
                {
                    Connection = conn,
                    CommandText = sql
                };
                cmd.ExecuteNonQuery();
            }
            catch (Exception )
            {
                Console.WriteLine("Не удалось удалить записи с базы данных");
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
        }

        public static async void InsertToDataBase(string id, string name, string value, string valueTime)
        {
            MySqlConnection conn = DBUtils.GetDBConnection();
            try
            {
                await conn.OpenAsync();
                string sql = "INSERT INTO matrixArray (Id, Name, Debit, NormOfTime)  VALUES (@id, @name, @value, @time);";
                MySqlCommand cmd = new MySqlCommand
                {
                    Connection = conn,
                    CommandText = sql
                };
                cmd.Parameters.Add("@id", MySqlDbType.VarChar).Value = id;
                cmd.Parameters.Add("@name", MySqlDbType.VarChar).Value = name;
                cmd.Parameters.Add("@value", MySqlDbType.Float).Value = value;
                cmd.Parameters.Add("@time", MySqlDbType.Float).Value = valueTime; //date.Date.ToString("yyyy/MM/dd");
                // cmd.Parameters.Add("@date", MySqlDbType.Date).Value = DateImport;
                await cmd.ExecuteNonQueryAsync();
            }
            catch (Exception )
            {
               //Console.WriteLine("Какая-то ошибка при заполнении базы данных" + ex);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
        }
    }
}
