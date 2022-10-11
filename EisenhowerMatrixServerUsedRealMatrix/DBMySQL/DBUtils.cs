using MySql.Data.MySqlClient;
using System;
using System.Configuration;

namespace SqlConn
{
    class DBUtils
    {
        public static MySqlConnection GetDBConnection()//параметры строки подключения к БД
        {
            string host = ConfigurationManager.AppSettings.Get("host");
            int port = Convert.ToInt32(ConfigurationManager.AppSettings.Get("port"));
            string database = ConfigurationManager.AppSettings.Get("database");
            string username = ConfigurationManager.AppSettings.Get("username");
            string password = ConfigurationManager.AppSettings.Get("password");

            return DBMySQLUtils.GetDBConnection(host, port, database, username, password);
        }
    }
}
