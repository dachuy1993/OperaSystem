using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for DelayLoadingData.xaml
    /// </summary>
    public partial class Page_LoadingData : Page
    {
        public Page_LoadingData()
        {
            InitializeComponent();
            Loaded += DelayLoadingData_Loaded;
        }
        public static bool checkDelayLoading = false;
        private void DelayLoadingData_Loaded(object sender, RoutedEventArgs e)
        {
            txbLoading.Text = MainWindow.txbloading;
            checkDelayLoading = true;
        }
        DispatcherTimer dt = new DispatcherTimer();
        
        



    }
}
