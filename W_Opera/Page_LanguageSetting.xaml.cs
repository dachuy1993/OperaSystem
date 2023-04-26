using System;
using System.Collections.Generic;
using System.Linq;
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
using System.Windows.Shapes;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for LanguageSetting.xaml
    /// </summary>
    public partial class Page_LanguageSetting : Window
    {
        public Page_LanguageSetting()
        {
            InitializeComponent();
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
           
        }

        
        private void RbLanguage_Eng_Click(object sender, RoutedEventArgs e)
        {
            if (rbLanguage_Eng.IsChecked == true)
            {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            }                         
            DialogResult = true;
        }

        private void RbLanguage_Kor_Click(object sender, RoutedEventArgs e)
        {
            if (rbLanguage_Kor.IsChecked == true)
            {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ko-KR");               
            }
            DialogResult = true;
        }

        private void RbLanguage_Vni_Click(object sender, RoutedEventArgs e)
        {
            if (rbLanguage_Vni.IsChecked == true)
            {
                Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("vi-VN");               
            }                         
            DialogResult = true;
        }
    }
}
