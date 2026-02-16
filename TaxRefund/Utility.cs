using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaxRefund
{
    class Utility
    {
        public static SqlConnection conn;       

        public SqlConnection OpenDB()
        {
            IPHostEntry host;
            host = Dns.GetHostEntry(Dns.GetHostName());

            foreach (IPAddress ip in host.AddressList)
            {
                if(ip.ToString().Contains("10.228."))
                {
                    conn = new SqlConnection(@"Data Source= 10.228.48.200, 1433; Initial Catalog=TaxRefund; User ID = ducnq100152; Password = 123456a@");
                }
                else
                {
                    conn = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=TaxRefund;Integrated Security=True");
                }               
            }
            return conn;
        } 
    }
}
