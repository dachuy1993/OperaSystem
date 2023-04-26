using DataHelper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Page_RejectSample.xaml
    /// </summary>
    public partial class Page_RejectSample : Window
    {
        public static string samno;
        public static string typeSample;
        public static string modelcode;
        string path_sql = MainWindow.path_sql;
        string date = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");
        public Page_RejectSample()
        {
            InitializeComponent();
            Loaded += Page_RejectSample_Loaded;
        }

        private void Page_RejectSample_Loaded(object sender, RoutedEventArgs e)
        {
            path_sql = "Data Source=" + MainWindow.ip + ";Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
            Read_RejectSample();
        }

        
        public string Read_SeqSampleReject(string samNo, string typeSample)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "SELECT max(seq) FROM tbSampleReject Where samno='" + samNo + "' and typeSample='" + typeSample + "'";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        string seq = cmd.ExecuteScalar().ToString();
                        if (seq != "")
                        {
                            seq = (int.Parse(seq) + 1).ToString("0000");
                        }
                        else
                        {
                            seq = ("0001");
                        }
                        conn.Close();
                        return seq;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Page RejectSample/Read_SeqSampleReject", MessageBoxButton.OK, MessageBoxImage.Error);
                return "null";
            }
        }
        public void Inser_RejectSample()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                    var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                    string seq = Read_SeqSampleReject(samno, typeSample);
                    string remark = txt_Reject.Text;
                    string imsempcode = MainWindow.UserLogin;
                    string insdt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                    string updempcode = MainWindow.UserLogin;
                    string upddt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                    Helper_TaixinDB_SampleBox taixinDB = new Helper_TaixinDB_SampleBox();
                    var command = "INSERT tbSampleReject(cmpcode,bizdiv,samno,seq,typeSample,modelcode,remark,imsempcode,insdt,updempcode,upddt) Values('02','300','" + samno + "','" + seq + "','" + typeSample + "','" + modelcode + "',N'" + remark + "','" + imsempcode + "','" + insdt + "','" + updempcode + "','" + upddt + "')";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                    //MessageBox.Show("Phản hồi thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Page RejectSample/Inser_RejectSample", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void Update_ValueBoxManual()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                    var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                    string seq = Read_SeqSampleReject(samno, typeSample);
                    string remark = txt_Reject.Text;
                    string imsempcode = MainWindow.UserLogin;
                    string insdt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                    string updempcode = MainWindow.UserLogin;
                    string upddt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                    Helper_TaixinDB_SampleBox taixinDB = new Helper_TaixinDB_SampleBox();
                    var command = "";
                    if (typeSample == "Box")
                    {
                        command = "Update tbSampleBox set reject ='1' where samno ='" + samno + "'";
                    }
                    if (typeSample == "Manual")
                    {
                        command = "Update tbSampleManual set reject ='1' where samno ='" + samno + "'";
                    }
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Page RejectSample/Update_ValueBoxManual", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public List<Helper_TaixinDB_RejectSample> Read_RejectSample()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    List<Helper_TaixinDB_RejectSample> listSampleReject = new List<Helper_TaixinDB_RejectSample>();

                    var command = "SELECT * from tbSampleReject WHERE samno = '" + samno + "' and typeSample ='" + typeSample + "' order by insdt asc";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                if (dr[0] != null)
                                {
                                    Helper_TaixinDB_RejectSample taixinDB = new Helper_TaixinDB_RejectSample();
                                    taixinDB.samno = dr[2].ToString();
                                    taixinDB.seq = dr[3].ToString();
                                    taixinDB.typeSample = dr[4].ToString();
                                    taixinDB.modelcode = dr[5].ToString();
                                    taixinDB.remark = dr[16].ToString();
                                    taixinDB.insempcode = dr[17].ToString();
                                    taixinDB.insdt = dr[18].ToString();
                                    taixinDB.updempcode = dr[19].ToString();
                                    taixinDB.upddt = dr[20].ToString();
                                    listSampleReject.Add(taixinDB);
                                }
                            }
                        }
                    }
                    conn.Close();
                    lvSampleReject.ClearValue(ListView.ItemsSourceProperty);
                    lvSampleReject.ItemsSource = listSampleReject;
                    return listSampleReject;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Page RejectSample/Read_RejectSample", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private void btn_Reject_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn lưu phản hồi không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                Inser_RejectSample();
                Read_RejectSample();
                Update_ValueBoxManual();
            }
        }
    }
}
