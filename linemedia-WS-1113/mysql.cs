using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;
namespace linemedia_WS_1113
{
    class mysql
    {
        public MySqlConnection con;
        public MySqlCommand cmd = new MySqlCommand();
        public int irows = 0;
        public string srr = "";
        public string error = "";
        public String constr = "";
        public mysql()
        {
            try
            {

                con = new MySqlConnection(constr);
                con.Open();
            }
            catch (MySqlException ex)
            {
                srr = "";
                srr = ex.Message.ToString();
            }
        }

        public MySqlConnection mysqlconnection()
        {
            try
            {
                con = new MySqlConnection(constr);
                con.Open();
            }
            catch (MySqlException ex)
            {
                srr = "";
                srr = ex.Message.ToString();
            }
            return con;
        }
        public int executeDMLSQL(String s)
        {
            try
            {
                MySqlCommand cmd = new MySqlCommand(s, con);
                irows = cmd.ExecuteNonQuery();
            }
            catch (MySqlException ex)
            {
                srr = ex.Message.ToString();
            }
            return irows;
        }
        public DataSet executeSQL_dset(string ssql)
        {
            string srr = " ";
            DataSet ds = new DataSet();
            MySqlDataAdapter mda = new MySqlDataAdapter();
            try
            {
                MySqlCommand cmd = new MySqlCommand();
                cmd.CommandText = ssql;
                cmd.Connection = con;
                cmd.CommandTimeout = 6000;
                mda.SelectCommand = cmd;
                mda.Fill(ds);
            }
            catch (Exception e)
            {
                srr = e.Message.ToString();
            }
            return ds;
        }
        public void closeconnection()
        {
            con.Close();
        }


    }
}