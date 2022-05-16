using MySqlConnector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookToMariadb
{
    public class MySqlServer
    {
        public MySqlConnectionStringBuilder ConnectionBuilder;

        public MySqlServer(string server, string user, string password, string database)
        {
            ConnectionBuilder = new MySqlConnectionStringBuilder()
            {
                Server = server,
                UserID = user,
                Password = password,
                Database = database,
            };
        }

        public MySqlConnection Connect(bool async = false)
        {
            var res = new MySqlConnection(ConnectionBuilder.ConnectionString);
            if (async)
                res.OpenAsync();
            else
                res.Open();
            return res;
        }

        public void TruncateTable(string name)
        {
            ExecNonQuery($"truncate table {name}");
        }

        public void ExecNonQuery(string sql)
        {
            var conn = Connect();
            var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
            conn.Close();
        }
    }
}
