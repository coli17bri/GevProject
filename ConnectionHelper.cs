using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{
    public static class ConnectionHelper
    {
        private static String connectionParameter="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\\risorse\\Gev_Cagnola.accdb"; 
        private static OleDbConnection connection;

        public static  OleDbConnection createConnection() {
           connection = new OleDbConnection(connectionParameter); 
           return connection;
        }

    }
}
