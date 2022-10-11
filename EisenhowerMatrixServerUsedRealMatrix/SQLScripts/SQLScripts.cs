using MySql.Data.MySqlClient;
using System;


namespace SqlConn
{
    class SQLScripts
    {
        public static void DeleteAllFromTable(MySqlConnection conn)
        {
            try
            {
                string sql = "DELETE FROM matrixArray;";
                //sql = "TRUNCATE TABLE matrixArray;"; //удаляет все записи и заодно обнуляет автоинкремент id
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
            }
        }

        public static void Conn(MySqlConnection conn)
        {
            try
            {
                conn.Open();
                string sql = $"SET SESSION transaction ISOLATION LEVEL REPEATABLE READ";
                MySqlCommand cmd = new MySqlCommand
                {
                    Connection = conn,
                    CommandText = sql
                };
                cmd.ExecuteNonQuery();
                sql = " START TRANSACTION; ";
                cmd = new MySqlCommand
                {
                    Connection = conn,
                    CommandText = sql
                };
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {
                Console.WriteLine("Не уподключиться к базе данных");
            }
            finally
            {
            }
        }

        public static void ConnClose(MySqlConnection conn)
        {
            try
            {
                string sql = $"Commit;"; 
                MySqlCommand cmd = new MySqlCommand
                {
                    Connection = conn,
                    CommandText = sql
                };
                cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {
                Console.WriteLine("Не удалось завершить транзакции");
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
        }


        public static void InsertToDataBase(string id, string name, string value, string valueTime, MySqlConnection conn)
        {
           
            try
            {
                string sql = "INSERT INTO matrixArray (Id, Name, Debit, NormOfTime)  VALUES (@id, @name, @value, @time);";
                MySqlCommand cmd = new MySqlCommand
                {
                    Connection = conn,
                    CommandText = sql
                };
                cmd.Parameters.Add("@id", MySqlDbType.VarChar).Value = id;
                cmd.Parameters.Add("@name", MySqlDbType.VarChar).Value = name;
                cmd.Parameters.Add("@value", MySqlDbType.Float).Value = value;
                cmd.Parameters.Add("@time", MySqlDbType.Float).Value = valueTime; 
                cmd.ExecuteNonQuery();
            }
            catch (Exception )
            {
                Console.WriteLine("Не вставить данных в бд");
            }
            finally
            {
            }
        }
    }
}
