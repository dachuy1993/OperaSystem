using DataHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Tulpep.NotificationWindow;
using W_Opera.DAO;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        //public static string strSearch = "";
        public static string Ver = "V7.6";
        string str = "open";
        string IPUser = "";
        Thickness margin, margin1;
        double width_form, height_form;
        public static List<Helper_TabItem> listTabItem = new List<Helper_TabItem>();
        public static string ip = "192.168.2.10";
        Page_Sample_Manual Page_Sample_Manual = new Page_Sample_Manual();
        Page_Sample_Box Page_Sample_Box;
        Page_Setting PageSetting = new Page_Setting();

        TabControl tabMainControl = new TabControl();
        List<Helper_DataButton> ListButton_Left   = new List<Helper_DataButton>();
        List<Helper_DataButton> ListButton_Header = new List<Helper_DataButton>();
        DataBaseHelper db = new DataBaseHelper();
        public static string path_sql_attach = "Data Source=192.168.2.10;Initial Catalog=taixin_attach;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        public static string path_sql = "Data Source=192.168.2.10;Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        //public static string path_sql = "Data Source=192.168.2.10;Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        public static string txbloading="Loading Data";
        public static Print print = new Print();
        public static string pl_Print="Box";
        public static string at_Samno = "";   
        public static string at_ModelCode = "";
        public static string path_SaveEXL = "";
        public static Helper_TaixinDB_Input print_DB;
        public static bool checkPrint = false;
        public static string UserLogin;
        public static string PassLogin;
        public MainWindow()
        {
            this.InitializeComponent();           
            ApplyLanguage();
            GetUserLogin();
            tabMainControl.SelectedIndex = 0;
            Loaded += MainWindow_Loaded;           
        }
        
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                width_form = System.Windows.SystemParameters.PrimaryScreenWidth;
                height_form = System.Windows.SystemParameters.PrimaryScreenHeight;
            }
            else if (this.WindowState == WindowState.Minimized)
            {
                width_form = 780;
                height_form = 1366;
            }
            //Login login = new Login();
            //UserLogin = Login.UserLogin;
            //PassLogin = Login.PassLogin;

            Page_Sample_Box = new Page_Sample_Box();          
            CreatButton_Header();
            stackMainControl.Children.Add(tabMainControl);
            ChangeSize();
            StyleButton_Remove();
            Style_TextBlock();
            GetDataCmbTypeManual();


        }

        private void GetUserLogin()
        {
            IPUser = GetIPAddress();
            List<string> list = new List<string>();
            List<string> listTime = new List<string>();
            using (SqlConnection conn = new SqlConnection(path_sql))
            {
                conn.Open();
                {
                   
                    string query = "SPCheckTimeLogin @IPUser ";

                    DataTable ListTimeLogin = new DataTable();
                    ListTimeLogin = DataProvider.Instance.ExecuteSP(path_sql, query, new object[] { IPUser });
                    
                    foreach(DataRow item in ListTimeLogin.Rows)
                    {
                        UserLogin = item[0].ToString();
                        lbUserLogin.Content = "User:" + UserLogin;
                    }

                    //var command = "SELECT TOP 1 UserLogin ,DateLogin FROM TblUserIPLogin WHERE IPLogin = '" + IPUser + "' ORDER BY ID Desc";
                    //using (SqlCommand cmd = new SqlCommand(command, conn))
                    //{
                    //    using(IDataReader dr = cmd.ExecuteReader())
                    //    {
                    //        while(dr.Read())
                    //        {
                    //            list.Add(dr[0].ToString());
                    //            listTime.Add(dr[1].ToString());
                    //            if (dr[0] != null)
                    //            {
                    //                UserLogin = dr[0].ToString();
                    //                lbUserLogin.Content = "User:" + UserLogin;
                                    
                    //            }    
                    //        }    
                    //    }    
                    //}    
                }
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


        PopupNotifier popup = new PopupNotifier();
        public static string timeApproveManual = "";
        public static string timeApproveBox = "";
        

       
        public void ChangeSize()
        {
            if (str == "open")
            {
                stackMenuControl.Visibility = Visibility.Visible;
                margin = stackMenuControl.Margin;
                margin1 = stackMainControl.Margin;
                margin1.Left = 150;
                stackMenuControl.Margin = margin;
                stackMainControl.Margin = margin1;
                //border2.Width = w - 225;
            }
            if (str == "close")
            {
                stackMenuControl.Visibility = Visibility.Hidden;
                margin = stackMenuControl.Margin;
                margin1 = stackMainControl.Margin;
                margin1.Left = 0;
                stackMenuControl.Margin = margin;
                stackMainControl.Margin = margin1;
            }

        }

        private void ApplyLanguage(string cultureName = null)
        {
            if (cultureName != null)
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(cultureName);

            ResourceDictionary dict = new ResourceDictionary();
            switch (Thread.CurrentThread.CurrentCulture.ToString())
            {
                case "vi-VN":
                    dict.Source = new Uri("..\\Lang\\VietNam.xaml", UriKind.Relative);
                    break;
                case "ko-KR":
                    dict.Source = new Uri("..\\Lang\\Korea.xaml", UriKind.Relative);
                    break;
                // ...
                default:
                    dict.Source = new Uri("..\\Lang\\English.xaml", UriKind.Relative);
                    break;
            }
            this.Resources.MergedDictionaries.Add(dict);
        }

       

       
        public void CreatButton_Header()
        {
            ListButton_Header.Add(new Helper_DataButton {
                ID = 1, ContentButton = "Menu",
                ImageSource="Image/Dep/Home.png",
                BackGroundColor = PinValue.OFF });
            ListButton_Header.Add(new Helper_DataButton
            {
                ID = 2,
                ContentButton = "HR",
                ImageSource = "Image/Dep/HR.png",
                BackGroundColor = PinValue.OFF
            });
            ListButton_Header.Add(new Helper_DataButton
            {
                ID = 3,
                ContentButton = "ACC",
                ImageSource = "Image/Dep/Acc.png",
                BackGroundColor = PinValue.OFF
            });
            ListButton_Header.Add(new Helper_DataButton
            {
                ID = 4,
                ContentButton = "PRO",
                ImageSource = "Image/Dep/Pro.png",
                BackGroundColor = PinValue.OFF
            });
            ListButton_Header.Add(new Helper_DataButton
            {
                ID = 5,
                ContentButton = "QC",
                ImageSource = "Image/Dep/QC.png",
                BackGroundColor = PinValue.OFF
            });
            ListButton_Header.Add(new Helper_DataButton
            {
                ID = 6,
                ContentButton = "EQUIT",
                ImageSource = "Image/Dep/Equiment.png",
                BackGroundColor = PinValue.OFF

            });
            ListButton_Header.Add(new Helper_DataButton
            {
                ID = 7,
                ContentButton = "IT",
                ImageSource = "Image/Dep/IT.png",
                BackGroundColor = PinValue.OFF
            });
            ListButton_Header.Add(new Helper_DataButton
            {
                ID = 8,
                ContentButton = "SETTING",
                ImageSource = "Image/Dep/Setting.png",
                BackGroundColor = PinValue.OFF
            });
            foreach (var button in ListButton_Header)
            {
                lvButtonTop.Items.Add(button);
            }
        }

        
        public void ProcessButtonDep_Pro()
        {
            lvListItemMenu.Items.Clear();
            ListButton_Left.Clear();
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 1,
                ContentButton = "Nhập Kho Thường",
                ImageSource = "Image/Dep/Acc.png",
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 2,
                ContentButton = "Nhập Kho Khác",
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 3,
                ContentButton = "Xuất Kho Thường",
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 4,
                ContentButton = "Xuất Kho Khác",
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 5,
                ContentButton = "Nhập NVL",
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 6,
                ContentButton = "Xuất NVL",
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            foreach (var item in ListButton_Left)
            {
                lvListItemMenu.Items.Add(item);
            }            
        }

        public void ProcessButtonSetting()
        {
            lvListItemMenu.Items.Clear();
            ListButton_Left.Clear();
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 1,
                ContentButton = "Sample Setting",               
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            foreach (var item in ListButton_Left)
            {
                lvListItemMenu.Items.Add(item);
            }
        }

        public void ProcessButtonSample()
        {
            lvListItemMenu.Items.Clear();
            ListButton_Left.Clear();
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 1,
                ContentButton = "Sample Manual",               
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });
            ListButton_Left.Add(new Helper_DataButton
            {
                ID = 2,
                ContentButton = "Sample Box",
                CheckCreatTab = PinValue.OFF,
                BackGroundColor = PinValue.OFF
            });

            foreach (var item in ListButton_Left)
            {
                lvListItemMenu.Items.Add(item);
            }
           
        }

       
        private void ButtonItemMenu_Click(object sender, RoutedEventArgs e)
        {
            bool check = false;
            var click = sender as Button;
            var clickItem = click.DataContext as Helper_DataButton;
            if (clickItem != null)
            {
               
                foreach (var _Tab in listTabItem)
                {
                    if (_Tab.Name == clickItem.ContentButton)
                    {
                        check = true;
                        tabMainControl.SelectedItem = _Tab.TabItemCreat; 
                    }
                }
                if (check == false)
                {
                    CreatTabItem(clickItem.ContentButton.ToString());
                }
                foreach (var button in ListButton_Left)
                {
                    button.BackGroundColor = PinValue.OFF;
                    if (button.ContentButton == clickItem.ContentButton)
                    {
                        button.BackGroundColor = PinValue.ON;
                    }
                }
            }
        }

       
        private void ButtonTop_Click(object sender, RoutedEventArgs e)        
        {
            var click = sender as Button;
            var clickItem = click.DataContext as Helper_DataButton;
            if(clickItem != null)
            {
                switch (clickItem.ContentButton)
                {
                    case "Menu":
                        {
                            //MessageBox.Show(clickItem.ContentButton.ToString());
                            break;
                        }
                    case "PRO":
                        {                            
                            ProcessButtonSample();
                            break;
                        }
                    case "ACC":
                        {
                            ProcessButtonDep_Pro();
                            //ProcessButtonSample();
                            break;
                        }
                    case "SETTING":
                        {
                            ProcessButtonSetting();
                            break;
                        }

                }
                foreach (var button in ListButton_Header)
                {
                    button.BackGroundColor = PinValue.OFF;
                    if(button.ContentButton == clickItem.ContentButton)
                    {
                        button.BackGroundColor = PinValue.ON;
                    }
                }
               
            }
           
        }

       
        private void BtnOpenMenuControl_Click(object sender, RoutedEventArgs e)
        {
            str = "open";
            ChangeSize();           
        }

       
        private void BtnCloseMenuControl_Click(object sender, RoutedEventArgs e)
        {
            str = "close";
            ChangeSize();            
        }
        public static List<Helper_Combobox> _ListTypeD = new List<Helper_Combobox>();
        public static List<Helper_Combobox> _ListTypeM = new List<Helper_Combobox>();
        public  void GetDataCmbTypeManual()
        {
            try
            {
                _ListTypeD.Clear();
                ////lây dữ liệu lên cbb Year
                //string cbYear = "";
                //string queryYear = "SPGetDataCmbTypeDetailManual @cbYear ";

                //// Lấy dữ liệu và hiển thị
                //DataTable listCmbYear = new DataTable();

                //listCmbYear = DataProvider.Instance.ExecuteSP(path_sql, queryYear, new object[] { cbYear });


                //Helper_Combobox item = new Helper_Combobox();

                //foreach (DataRow Row in listCmbYear.Rows)
                //{
                //    item.code = Row["ChageChar1"].ToString();
                //    item.Name_loc = Row["Name_loc"].ToString();
                //    _ListTypeD.Add(item);
                //}


                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT Name_loc, ChageChar1 FROM temstetc WHERE GubunCode='962' UNION SELECT '' AS Name_loc, '000' AS ChageChar1", conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                if (dr[0] != null)
                                {
                                    Helper_Combobox item = new Helper_Combobox();
                                    item.code = dr[1].ToString();
                                    item.Name_loc = dr[0].ToString();
                                    _ListTypeD.Add(item);
                                }
                            }
                        }
                    }
                    conn.Close();
                }
                //cbbTypeCertification.ItemsSource = listResultYear;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void BtnMainExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void BtnAboutPage_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnMainGlobal_Click(object sender, RoutedEventArgs e)
        {
            Page_LanguageSetting langSetting = new Page_LanguageSetting();
            langSetting.Owner = this;
            if (langSetting.ShowDialog() == true)
            {
                ApplyLanguage();
            }
        }

        //===================================================================================

        Style styleButton_Remove = new Style(typeof(Button));
        Style styleTextBlock_Remove = new Style(typeof(TextBlock));
        Style styleTabItem = new Style(typeof(Button));
        public void StyleButton_Remove()
        {
            styleButton_Remove.Setters.Add(new Setter(Button.BackgroundProperty, Brushes.Transparent));
            styleButton_Remove.Setters.Add(new Setter(Button.WidthProperty, 20.0));
            styleButton_Remove.Setters.Add(new Setter(Button.HeightProperty, 20.0));
            styleButton_Remove.Setters.Add(new Setter(Button.ContentProperty, "x"));
            styleButton_Remove.Setters.Add(new Setter(Button.BorderBrushProperty, Brushes.Transparent));
            styleButton_Remove.Setters.Add(new Setter(Button.VerticalAlignmentProperty, VerticalAlignment.Center));
            styleButton_Remove.Setters.Add(new Setter(Button.HorizontalAlignmentProperty, HorizontalAlignment.Center));
            Resources.Add(typeof(Button), styleButton_Remove);
        }
        public void Style_TextBlock()
        {
            styleTextBlock_Remove.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
            styleTextBlock_Remove.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Center));
            styleTextBlock_Remove.Setters.Add(new Setter(TextBlock.MarginProperty, margin));
            Resources.Add(typeof(TextBlock), styleTextBlock_Remove);
        }

        public void CreatTabItem(string _HeadName)
        {
            Helper_TabItem creatItem = new Helper_TabItem();
            TabItem tabItem = new TabItem();
            Frame frame_TabItem = new Frame();
            StackPanel stackHeader_TapItem = new StackPanel();
            TextBlock txb_HeaderName = new TextBlock();
            Button btnRemoveTab = new Button();
            int count = 0;
            //Tạo ID cho từng TabItem
            foreach (var item in ListButton_Left)
            {
                if (item.ContentButton != "")
                {
                    count++;
                }
            }
            if (count == 0)
            {
                creatItem.ID = 1;
            }
            if (count > 0)
            {
                creatItem.ID = count + 1;
            }
            creatItem.Run = true;
            creatItem.Name = _HeadName;            
            frame_TabItem.NavigationUIVisibility = NavigationUIVisibility.Hidden;            
            txb_HeaderName.Text = _HeadName;
            txb_HeaderName.Style = styleTextBlock_Remove;
            btnRemoveTab.Click += BtnRemoveTab_Click;
            btnRemoveTab.Style = styleButton_Remove;
            stackHeader_TapItem.Orientation = Orientation.Horizontal;              
            stackHeader_TapItem.Children.Add(txb_HeaderName);
            stackHeader_TapItem.Children.Add(btnRemoveTab);
            tabItem.Header = stackHeader_TapItem;
            tabMainControl.SelectionChanged += TabMainControl_SelectionChanged;
            creatItem.TabItemCreat = tabItem;
            btnRemoveTab.DataContext = creatItem;
            tabItem.Background = Brushes.DodgerBlue;
            
            switch (_HeadName)
            {
               
                case "Sample Manual":
                    {
                        frame_TabItem.Navigate(Page_Sample_Manual);                        
                        break;
                    }
                case "Sample Box":
                    {
                        frame_TabItem.Navigate(Page_Sample_Box);
                        break;
                    }
                case "Sample Setting":
                    {
                        if(MainWindow.UserLogin.ToUpper()=="IT")
                        {
                            frame_TabItem.Navigate(PageSetting);
                        }
                        else
                        {
                            MessageBox.Show("Bạn không có quyền truy cập", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        break;
                    }
            }            
            tabItem.Content = frame_TabItem;           
            tabMainControl.Items.Add(tabItem);
            tabMainControl.VerticalAlignment = VerticalAlignment.Stretch;
            tabMainControl.HorizontalAlignment = HorizontalAlignment.Stretch;
            listTabItem.Add(creatItem);
            tabMainControl.SelectedItem = tabItem;           
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void btnMainAbout_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Opera System Version : " + Ver, "Version" , MessageBoxButton.OK, MessageBoxImage.Information);
            Window_About about = new Window_About();
            about.Show();
        }

        private void TabMainControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void BtnRemoveTab_Click(object sender, RoutedEventArgs e)
        {
            var click = sender as Button;
            var clickItem = click.DataContext as Helper_TabItem;           
            if (clickItem != null)
            {                
              
                foreach (var item in listTabItem)
                {
                    if(item.Name == clickItem.Name)
                    {
                        listTabItem.Remove(item);
                        tabMainControl.Items.Remove(clickItem.TabItemCreat);
                        break;
                    }
                }
            }            
        }


        

    }
    
    public class ModelCode
    {
        public static Helper_TaixinDB_Model ModelCodeStatus = new Helper_TaixinDB_Model();
    }


    public class DOCode
    {
        public static Helper_TaixinDB_DO DOCodeStatus = new Helper_TaixinDB_DO();
    }
    public class Helper_TabItem
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public TabItem TabItemCreat { get; set; }

        public bool Run { get; set; }
    }

   
}
