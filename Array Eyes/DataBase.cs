using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace Array_Eyes
{
    class DataBase
    {
        OleDbConnection oledbcon, oledbcon1;
        OleDbCommand oledbcmd;
        OleDbDataAdapter oledbda;
        DataSet myds;

        public string DataSource = "";

        public OleDbConnection getCon()
        {
            // DataSource = Form21.LocalDataSource;

            oledbcon = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;");
            return oledbcon;
        }

        public void getCmd(string strSql)
        {
            oledbcon = getCon();
            oledbcmd = new OleDbCommand(strSql, oledbcon);
            oledbcon.Open();
            oledbcmd.ExecuteNonQuery();
            oledbcon.Close();
        }

        public DataSet getDs(string strSql)
        {
            oledbcon = getCon();
            oledbda = new OleDbDataAdapter(strSql, oledbcon);
            myds = new DataSet();
            oledbda.Fill(myds);
            return myds;
        }
    }
}
