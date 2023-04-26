//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace W_Opera.DAO
{
    public class DataProvider
    {
        private static DataProvider instance;
        public static DataProvider Instance
        {
            get 
            { 
                if (instance == null)
                    instance = new DataProvider();
                return instance; 
            }
            private set => instance = value; 
        }

        public DataTable executeQuery(string str, string Query, object[] parameter = null)
        {
            DataTable dt = new DataTable();
            using(SqlConnection con = new SqlConnection(str))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(Query,con);
                if(parameter != null)
                {
                    string[] listpara = Query.Split(' ');
                    int i = 0;
                    foreach(string item in listpara)
                    {
                        if(item.Contains('@'))
                        {
                            cmd.Parameters.Add(new SqlParameter(item,SqlDbType.NVarChar)).Value = parameter[i];
                            i++;
                        }
                    }
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    con.Close();
                }
                return dt;
            }
        }

        public DataTable ExecuteSP(string str, string query, object[] paramater = null)
        {
            DataTable dt = new DataTable();
            using(SqlConnection con = new SqlConnection(str))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(query,con);
                if(paramater != null)
                {
                    string[] listpara = query.Split(' ');
                    int i = 0;
                    foreach(string item in listpara)
                    {
                        if(item.Contains('@'))
                        {
                            cmd.Parameters.Add(new SqlParameter(item,SqlDbType.NVarChar)).Value = paramater[i];
                            i++;
                        }
                    }
                }
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                con.Close();
            }
            return dt;
        }

        public object ExecuteScalar(string str, string query, object[] parameter = null)
        {
            object a = 0;
            using (SqlConnection con = new SqlConnection(str))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(query, con);
                if (parameter != null)
                {
                    string[] listpara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listpara)
                    {
                        if (item.Contains('@'))
                        {
                            cmd.Parameters.Add(new SqlParameter(item, SqlDbType.NVarChar)).Value = parameter[i];
                            i++;
                        }
                    }
                }
                a = cmd.ExecuteScalar();
                con.Close();
            }
            return a;
        }

        public List<string> GetList(string str, string query, object[] parameter = null)
        {
            List<string> list = new List<string>();
            using (SqlConnection con = new SqlConnection(str))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(query, con);
                if (parameter != null)
                {
                    string[] listpara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listpara)
                    {
                        if (item.Contains('@'))
                        {
                            cmd.Parameters.Add(new SqlParameter(item, SqlDbType.NVarChar)).Value = parameter[i];
                            i++;
                        }
                    }
                }
                SqlDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    int i = 0;
                    list.Add(dr.GetString(0));
                    i += 1;
                }
                con.Close();
            }
            return list;
        }
    }
}
