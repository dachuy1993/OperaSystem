using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
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
using System.Net;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Login.xaml
    /// </summary>
    public partial class Login : Window
    {
        public static string pathFileIni;
        public static string pathFolderIni;
        string pathSql = "";
        string ip = "";
        string verSQL, bufferExe, fileType, libraryName;
        //public static bool SaveLimitValue = false;
        string verFramework = "";
        string db_user = "sa";
        string db_pass = "oneuser1!";
        string nameApplication = "taixin_db";
        string nameFolderIni = "W_Opera";
        string nameFolderExe = @"C:\OperaSystem\OperaSystem\";
        bool CheckRememberLogin = false;
        bool CheckShowPass = false;
        bool checkLogin = false;
        public static string UserLogin;
        public static string PassLogin;

        public Login()
        {
            InitializeComponent();
            pb_Pass.Visibility = Visibility.Visible;
            txtPass.Visibility = Visibility.Hidden;
            btnShowPass.Visibility = Visibility.Hidden;
            btnHidenPass.Visibility = Visibility.Visible;
            pathSql = @"Data Source= 192.168.2.10;Initial Catalog=" + nameApplication + ";Persist Security Info=True;User ID=" + db_user + ";Password=" + db_pass + "";
            //pathSql = "Data Source=192.168.2.10;Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        }

        public void Db_Read_Version()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(pathSql))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("select top(1)* from FileUpdate_Opera order by ID desc", conn))
                    {
                        using (IDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                if (dr[0] != null)
                                {
                                    libraryName = dr[6].ToString();
                                    fileType = dr[2].ToString();
                                    verSQL = dr[4].ToString();
                                    bufferExe = dr[5].ToString();
                                }
                            }
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReadVersion_SQLserver" + ex.Message, "Login/MainWindow", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void Db_Read_Employee()
        {
            try
            {
                checkLogin = false;
                //string path_TaixinWeb = "server=192.168.2.10;Port=3307;user id=txadmin;database=LiveSystem;password=Taixinweb1!";
                List<string> list = new List<string>();
                using (SqlConnection conn = new SqlConnection(pathSql))
                {
                    conn.Open();
                    {
                        var command = "SELECT emp_code,password FROM pf_employee";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            using (IDataReader dr = cmd.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    list.Add(dr[0].ToString());
                                    if (dr[0] != null)
                                    {
                                        if (txt_User.Text.ToUpper() == dr[0].ToString().Trim().ToUpper() && (txtPass.Text.ToUpper() == dr[1].ToString().Trim().ToUpper() || pb_Pass.Password.ToUpper() == dr[1].ToString().Trim().ToUpper()))
                                        {
                                            checkLogin = true;
                                        }
                                    }
                                }

                            }
                        }

                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ReadVersion_SQLserver" + ex.Message, "Login/MainWindow", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private string GetIPAddress()
        {
            string IPAddress = string.Empty;
            IPHostEntry Host = default(IPHostEntry);
            string Hostname = null;
            Hostname = System.Environment.MachineName;
            Host = Dns.GetHostEntry(Hostname);
            foreach (IPAddress IP in Host.AddressList)
            {
                if (IP.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    IPAddress = Convert.ToString(IP);
                }
            }
            return IPAddress;
        }

        public void Process_UpdateExeFile()
        {
            Db_Read_Version();
            string fileExeOperasystem = nameFolderExe + nameFolderIni + ".exe";
            Process.Start(fileExeOperasystem);
        }


        public void User_Login()
        {
            try
            {
                Db_Read_Employee();
                if (checkLogin == true)
                {
                    GetIPAddress();
                    MainWindow Main = new MainWindow();
                    Login login = new Login();
                    UserLogin = txt_User.Text.ToUpper();
                    PassLogin = pb_Pass.Password.ToUpper();
                    //Process_UpdateExeFile();  
                    Main.Show();
                    login.Hide();
                    //this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CheckLoginInput : " + ex.Message, "Login/MainWindow", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Btn_Login_Click(object sender, RoutedEventArgs e)
        {
            User_Login();
        }
        private void Pb_Pass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                User_Login();
            }
        }

        private void Txt_User_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void Txt_User_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                pb_Pass.Focus();
            }
        }

        private void TxtPass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                User_Login();
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }



        private void CkbRemember_Unchecked(object sender, RoutedEventArgs e)
        {
            CheckRememberLogin = false;
            txtPass.Text = "";
            txt_User.Text = "";
            pb_Pass.Password = "";
            ckbRemember.IsChecked = false;
            txt_User.IsEnabled = true;
            txtPass.IsEnabled = true;
            pb_Pass.IsEnabled = true;
            btnHidenPass.IsEnabled = true;
        }

        private void CkbRemember_Checked(object sender, RoutedEventArgs e)
        {
            CheckRememberLogin = true;
            ckbRemember.IsChecked = true;
            txt_User.IsEnabled = false;
            txtPass.IsEnabled = false;
            pb_Pass.IsEnabled= false;


        }

        private void Btn_ShowPass_Click(object sender, RoutedEventArgs e)
        {
            pb_Pass.Password = txtPass.Text;
            pb_Pass.Visibility = Visibility.Visible;
            txtPass.Visibility = Visibility.Hidden;
            btnShowPass.Visibility = Visibility.Hidden;
            btnHidenPass.Visibility = Visibility.Visible;
            CheckShowPass = false;
        }

        private void Btn_HidenPass_Click(object sender, RoutedEventArgs e)
        {
            txtPass.Text = pb_Pass.Password;
            pb_Pass.Visibility = Visibility.Hidden;
            txtPass.Visibility = Visibility.Visible;
            btnShowPass.Visibility = Visibility.Visible;
            btnHidenPass.Visibility = Visibility.Hidden;
            CheckShowPass = true;
        }

        public static bool IsUserAdministrator()
        {
            bool isAdmin;
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException ex)
            {
                isAdmin = false;
            }
            catch (Exception ex)
            {
                isAdmin = false;
            }
            return isAdmin;
        }
    }
}
