using DataHelper;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using System.Windows.Threading;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using Tulpep.NotificationWindow;
using Newtonsoft.Json;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Page_Sample.xaml
    /// </summary>
    public partial class Page_Sample : Page
    {
        #region Khai báo
        string path_sql = "Data Source=192.168.2.5;Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password=oneuser1!";
        List<Helper_TaixinDB_Model> ListAllDataSample = new List<Helper_TaixinDB_Model>();
        List<Helper_DataButton> listButtonTop = new List<Helper_DataButton>();
        Helper_TaixinDB_Model ApprovalClickItem = new Helper_TaixinDB_Model();
        Helper_AccessManger access_db = new Helper_AccessManger();
        List<Helper_AccessManger> list_access = new List<Helper_AccessManger>();
        PopupNotifier popup = new PopupNotifier();
        DispatcherTimer dt = new DispatcherTimer();
        string date = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");
        string str_cbbFilterApprove = "Tìm kiếm All";
        string dateFilterStart = "";
        string dateFilterFinish = "";
        bool checkPopup = false;
        string dateApproval = "";
        string department = "";
        string str_CMPCODE;
        string str_BIZDIV;
        string str_SAMNO;
        string str_MODELCODE;
        string str_MODELNAME;
        string str_APPLYDT;
        string str_VERSION;
        string str_CUSTPARTCODE;
        string str_CUSTPART_VERSION;
        string str_CUSTPARTCODE_VER;
        string str_CUSTMODELCODE;
        string str_REPMODELCODE;
        string str_USEFLAG;
        string str_MODELGROUP;
        string str_MODELDIV;
        string str_MODELTYPE;
        string str_MODELCHILD;
        string str_CUST_GB;
        string str_TA;
        string str_BUYER;

        string str_COLORIn;
        string str_MODELSPECIn;
        string str_MODELLENGTHIn;
        string str_MODELWIDTHIn;
        string str_MODELHEIGHTIn;
        string str_MODELUNFOLDLENGTHIn;
        string str_MODELUNFOLDWIDTHIn;

        string str_COLOROut;
        string str_MODELSPECOut;
        string str_MODELLENGTHOut;
        string str_MODELWIDTHOut;
        string str_MODELHEIGHTOut;
        string str_MODELUNFOLDLENGTHOut;
        string str_MODELUNFOLDWIDTHOut;

        string str_CUSTCODE;
        string str_CUSTSHORTCODE;
        string str_PAGECNT;
        string str_TYPE;
        string str_PAPERGUBUNIn;
        string str_PAPERGUBUNOut;


        string str_WEIGHTIn;
        string str_WIDTHIn;
        string str_HEIGHTIn;
        string str_SIDEGUBUNIn;
        string str_FRONTCCIn;
        string str_BACKCCIn;
        string str_FRONTBCOLORIn;
        string str_BACKBCOLORIn;
        string str_BCOLORCODEIn;
        string str_PHCOUNTIn;
        string str_TOTALPAGEIN;
        string str_FOLDINGPAGEIN;
        string str_TAYIN;
        string str_TAYPAGEIN;
        string str_TOTALPAGEOUT;
        string str_FOLDINGPAGEOUT;
        string str_TAYOUT;
        string str_TAYPAGEOUT;

        string str_WEIGHTOut;
        string str_WIDTHOut;
        string str_HEIGHTOut;
        string str_SIDEGUBUNOut;
        string str_FRONTCCOut;
        string str_BACKCCOut;
        string str_FRONTBCOLOROut;
        string str_BACKBCOLOROut;
        string str_BCOLORCODEOut;
        string str_PHCOUNTOut;
        string str_MA;
        string str_SX;
        string str_CL;
        string str_RD;
        string str_KD;
        string str_DATEINPUT;
        string str_DATEAPPROVE;
        string str_PAPERNAMEIn;
        string str_PAPERNAMEOut;
        string str_PAPERNAMEFULLIn;
        string str_PAPERNAMEFULLOut;

        string str_VERSIONUP;
        string str_INSEMPCODE;
        string str_DATESTARTAPPROVE;
        string str_DATEFINISHAPPROVE;
        string str_NOTE1;
        string str_NOTE2;
        string str_PROCESSWORDK;
        string str_QTYREQUEST;

        #endregion

        public Page_Sample()
        {
            InitializeComponent();        
            Loaded += Page_Sample_Loaded;
        }
        private void Page_Sample_Loaded(object sender, RoutedEventArgs e)
        {          
           
            Page_Sample_Box sample_Box = new Page_Sample_Box();
            frameSample.Navigate(sample_Box);
        }
        
    }

    public class SamplePaper
    {
        public static Helper_TaixinDB_Model PaperIn = new Helper_TaixinDB_Model();
        public static Helper_TaixinDB_Model PaperOut = new Helper_TaixinDB_Model();
        public static List<Helper_TaixinDB_Model> UnitBox = new List<Helper_TaixinDB_Model>();
    }
}
