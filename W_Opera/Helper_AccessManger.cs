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

        private string _AddMan;
        private string _DelMan;
        private string _EditMan;
        private string _SaveMan;
        private string _AddBox;
        private string _DelBox;
        private string _EditBox;
        private string _SaveBox;

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
        public string AddMan { get { return _AddMan; } set { if (_AddMan != value) { _AddMan = value; NotifyPropertyChanged("AddMan"); } } }
        public string DelMan { get { return _DelMan; } set { if (_DelMan != value) { _DelMan = value; NotifyPropertyChanged("DelMan"); } } }
        public string EditMan { get { return _EditMan; } set { if (_EditMan != value) { _EditMan = value; NotifyPropertyChanged("EditMan"); } } }
        public string SaveMan { get { return _SaveMan; } set { if (_SaveMan != value) { _SaveMan = value; NotifyPropertyChanged("SaveMan"); } } }
        public string AddBox { get { return _AddBox; } set { if (_AddBox != value) { _AddBox = value; NotifyPropertyChanged("AddBox"); } } }
        public string DelBox { get { return _DelBox; } set { if (_DelBox != value) { _DelBox = value; NotifyPropertyChanged("DelBox"); } } }
        public string EditBox { get { return _EditBox; } set { if (_EditBox != value) { _EditBox = value; NotifyPropertyChanged("EditBox"); } } }
        public string SaveBox { get { return _SaveBox; } set { if (_SaveBox != value) { _SaveBox = value; NotifyPropertyChanged("ProRun"); } } }

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
                                    _access.AddMan = dr[9].ToString();
                                    _access.DelMan = dr[10].ToString();
                                    _access.EditMan = dr[11].ToString();
                                    _access.SaveMan = dr[12].ToString();
                                    _access.AddBox = dr[13].ToString();
                                    _access.DelBox = dr[14].ToString();
                                    _access.EditBox = dr[15].ToString();
                                    _access.SaveBox = dr[16].ToString();
                                    _access.DateApprove = dr[18].ToString();
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

        public void Insert_AccessApprove(string path_sql, string nameTable,string samno,string user,string pass,string dep,string creat,string approve,string run,string add, string del, string edit, string save, string addBox, string delBox, string editBox, string saveBox, string date)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();                   
                    var command = "INSERT tbSampleAccess(cmpcode,bizdiv,samno,UserLogin,NameLogin,Department,Creater,Approver,ProRun,AddMan,DelMan,EditMan,SaveMan,AddBox,DelBox,EditBox,SaveBox,insdt) VALUES " +
                        " ('02','300','"+samno+"',N'" + user + "',N'" + pass + "','" + dep + "'," +
                        "'" + creat + "','" + approve + "','"+run+"','"+ add + "','" + del + "'," +
                        "'" + edit + "','" + save + "','" + addBox + "','" + delBox + "','" + editBox + "','" + saveBox + "','" + date + "')";
                     
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

        public void Update_AccessApprove(string path_sql,string nameTable,string ID, string user, string name, string dep, string creat, string approve,string run, string add, string del, string edit, string save, string addBox, string delBox, string editBox, string saveBox, string date)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "UPDATE tbSampleAccess SET UserLogin = '"+user+ "', NameLogin = N'" + name + "', Department = '" + dep + "', " +
                        "Creater = '" + creat + "', Approver = '" + approve + "',ProRun ='"+run+ "',AddMan ='" + add + "',DelMan ='" + del  + "'," +
                        "EditMan = '" + edit + "', SaveMan = '" + save + "',AddBox ='" + addBox + "',DelBox ='" + delBox + "',EditBox ='" + editBox + "'," +
                        "SaveBox = '" + saveBox + "',insdt = '" + date + "' WHERE samno = '" + ID + "'";

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
