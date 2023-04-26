using DataHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace W_Opera
{
    public class Helper_AccessManger : INotifyPropertyChanged
    {       
        private int _id;
        private int _samno;
        private string _UserApprove;
        private string _NameApprove;
        private string _DepApprove;
        private string _CreatApprove;
        private string _ApproveApprove;       
        private string _ProRun;
        private string _DateApprove;
        private PinValue _color;
        public int ID { get { return _id; } set { if (_id != value) { _id = value; NotifyPropertyChanged("ID"); } } }
        public int SamNo { get { return _samno; } set { if (_samno != value) { _samno = value; NotifyPropertyChanged("SamNo"); } } }
        public string UserApprove { get { return _UserApprove; } set { if (_UserApprove != value) { _UserApprove = value; NotifyPropertyChanged("UserApprove"); } } }
        public string NameApprove { get { return _NameApprove; } set { if (_NameApprove != value) { _NameApprove = value;NotifyPropertyChanged("NameApprove"); } } }
        public string DepApprove { get { return _DepApprove; } set { if (_DepApprove != value) { _DepApprove = value; NotifyPropertyChanged("DepApprove"); } } }
        public string CreatApprove { get { return _CreatApprove; } set { if (_CreatApprove != value) { _CreatApprove = value; NotifyPropertyChanged("CreatApprove"); } } }
        public string ApproveApprove { get { return _ApproveApprove; } set { if (_ApproveApprove != value) { _ApproveApprove = value; NotifyPropertyChanged("ApproveApprove"); } } }
        public string ProRun { get { return _ProRun; } set { if (_ProRun != value) { _ProRun = value; NotifyPropertyChanged("ProRun"); } } }

        public string DateApprove { get { return _DateApprove; } set { if (_DateApprove != value) { _DateApprove = value; NotifyPropertyChanged("DateApprove"); } } }     
        private void NotifyPropertyChanged(string v)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(v));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public string Read_MaxSamno(string path_sql, string nameTable)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    string max = "";
                    string command = "SELECT max(samno) FROM tbSampleAccess";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        max = (int.Parse(cmd.ExecuteScalar().ToString()) + 1).ToString();
                        
                    }
                    conn.Close();
                    return max;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper AccessManger/Read_MaxSamno", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }
        public List<Helper_AccessManger> Read_AccessApprove(string path_sql, string nameTable,string command)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_AccessManger> listAccessApprove = new List<Helper_AccessManger>();
                    //var command = "SELECT * from " + nameTable + "";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_AccessManger _access = new Helper_AccessManger();
                                if (dr[0] != null)
                                {
                                    _access.SamNo = int.Parse(dr[2].ToString());
                                    _access.UserApprove = dr[3].ToString();
                                    _access.NameApprove = dr[4].ToString();
                                    _access.DepApprove = dr[5].ToString();
                                    _access.CreatApprove = dr[6].ToString();
                                    _access.ApproveApprove = dr[7].ToString();
                                    _access.ProRun = dr[8].ToString();
                                    //_access.DateApprove = dr[18].ToString();
                                }
                                listAccessApprove.Add(_access);
                            }
                        }
                    }
                    conn.Close();
                    return listAccessApprove;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper AccessManger/Read_AccessApprove", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public void Insert_AccessApprove(string path_sql, string nameTable,string samno,string user,string pass,string dep,string creat,string approve,string run,string date)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();                   
                    var command = "INSERT tbSampleAccess(cmpcode,bizdiv,samno,UserLogin,NameLogin,Department,Creater,Approver,ProRun,insdt) VALUES " +
                        " ('02','300','"+samno+"',N'" + user + "',N'" + pass + "','" + dep + "'," +
                        "'" + creat + "','" + approve + "','"+run+"','" + date + "')";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {                       
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper AccessManger/Insert_AccessApprove", MessageBoxButton.OK, MessageBoxImage.Error);               
            }
        }

        public void Delete_AccessApprove(string path_sql, string nameTable,string ID)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "DELETE tbSampleAccess WHERE samno = '" + ID + "'";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper AccessManger/Delete_AccessApprove", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void Update_AccessApprove(string path_sql,string nameTable,string ID, string user, string name, string dep, string creat, string approve,string run,string date)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "UPDATE tbSampleAccess SET UserLogin = '"+user+ "', NameLogin = N'" + name + "', Department = '" + dep + "', " +
                        "Creater = '" + creat + "', Approver = '" + approve + "',ProRun ='"+run+"', insdt = '" + date + "' WHERE samno = '" + ID + "'";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper AccessManger/Update_AccessApprove", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        

    }
}
