using System;
using System.Collections.Generic;
using System.IO;
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
    /// Interaction logic for Window_About.xaml
    /// </summary>
    public partial class Window_About : Window
    {
        public Window_About()
        { 
            InitializeComponent();
            txt_Version.Text = "Opera System IP : " + MainWindow.ip + "\rVersion : " + MainWindow.Ver;

        }        

        private void cbb_IpSetting_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var click = sender as ComboBox;
            var clickItem = click.SelectedItem as ComboBoxItem;
            string user = MainWindow.UserLogin;
            string pass = MainWindow.PassLogin;

            if (user.ToUpper() != "IT" || pass.ToUpper() != "ITSYSTEM")
            {
                lab_IpSetting.Visibility = Visibility.Hidden;
                cbb_IpSetting.Visibility = Visibility.Hidden;
            }
            else
            {
                string ip = clickItem.Content.ToString();
                MainWindow.ip = ip;
                if(ip=="192.168.2.10")
                {
                    MainWindow.path_sql_attach = "Data Source=192.168.2.10;Initial Catalog=taixin_attach;Persist Security Info=True;User ID=sa;Password= oneuser1!";
                }  
                else
                {
                    MainWindow.path_sql_attach = "Data Source=192.168.2.5;Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!"; 
                }          
            }
        }
    }

}
