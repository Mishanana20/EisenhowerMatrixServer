using System;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;


namespace SqlConn
{
    class DBMySQLUtils
    {
        public static MySqlConnection
                 GetDBConnection(string host, int port, string database, string username, string password)
        {
            String connString = "Server=" + host + ";Database=" + database
                + ";port=" + port + ";User Id=" + username + ";password=" + password + "; convert zero datetime=True; Connection Timeout=43200;";
            MySqlConnection conn = new MySqlConnection(connString);

            return conn;
        }
    }
}
