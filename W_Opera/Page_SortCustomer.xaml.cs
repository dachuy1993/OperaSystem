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
using System.Windows.Shapes;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for ListCustomerPartCode.xaml
    /// </summary>
    public partial class Page_SortCustomer : Window
    {
        public Page_SortCustomer()
        {
            InitializeComponent();
            Loaded += ListCustomerPartCode_Loaded;
            
        }
        public string  path_sql;

        //public string path_sql = "Data Source=(local);Initial Catalog=taixin_db;Integrated Security=True;";
        public string nameTableModelCode = "TMSTMODEL";
        public string strCustomerSort = "";
        List<Helper_TaixinDB_Model> listModelCode = new List<Helper_TaixinDB_Model>();
        List<Helper_TaixinDB_Model> listModelCodeSearch = new List<Helper_TaixinDB_Model>();
        //Page_Sample_Manual Page_Sample_Manual = new Page_Sample_Manual();
        
        private void ListCustomerPartCode_Loaded(object sender, RoutedEventArgs e)
        {
            path_sql = MainWindow.path_sql;
            try
            {
                

                strCustomerSort = Page_Sample_Manual.CustomerCode;
                lvCustomerPartCode.Items.Clear();
                DataBaseHelper db = new DataBaseHelper();               
                listModelCode = db.Read_TaxinDb_ModelCode(path_sql, nameTableModelCode, "custpartcode",strCustomerSort);
                foreach (var _modelCode in listModelCode)
                {
                    listModelCodeSearch.Add(_modelCode);
                }
                lvCustomerPartCode.ItemsSource = listModelCodeSearch;
                //lvCustomerPartCode.ItemsSource = listModelCode;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ListCustomerPartCode :"+"\r\n" + ex.Message, "List Customer Infomation", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
        {

        }

        public static string strModelCode = "";  
        
        private void LvCustomerPartCode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var click = sender as ListView;
            var clickItem = click.SelectedItem as Helper_TaixinDB_Model;
            if(clickItem != null)
            {
                strModelCode = clickItem.TMSTMODEL_ModelCode.ToString();

                ModelCode.ModelCodeStatus.TMSTMODEL_ModelCode = clickItem.TMSTMODEL_ModelCode.ToString();
                ModelCode.ModelCodeStatus.TMSTMODEL_CustomerPartVer = clickItem.TMSTMODEL_CustomerPartVer.ToString();
                ModelCode.ModelCodeStatus.TMSTMODEL_ModelName = clickItem.TMSTMODEL_ModelName.ToString();
                ModelCode.ModelCodeStatus.TMSTMODEL_CustomerPartCode = clickItem.TMSTMODEL_CustomerPartCode.ToString();

                //ModelCode.ModelCodeStatus.CMPCODE = clickItem.CMPCODE.ToString();
                //ModelCode.ModelCodeStatus.BIZDIV = clickItem.BIZDIV.ToString();
                //ModelCode.ModelCodeStatus.MODELCODE = clickItem.MODELCODE.ToString();
                //ModelCode.ModelCodeStatus.MODELNAME = clickItem.MODELNAME.ToString();
                //ModelCode.ModelCodeStatus.APPLYDT = clickItem.APPLYDT.ToString();
                //ModelCode.ModelCodeStatus.VERSION = clickItem.VERSION.ToString();
                //ModelCode.ModelCodeStatus.CUSTPARTCODE = clickItem.CUSTPARTCODE.ToString();
                //ModelCode.ModelCodeStatus.CUSTPART_VERSION = clickItem.CUSTPART_VERSION.ToString();
                //ModelCode.ModelCodeStatus.CUSTPARTCODE_VER = clickItem.CUSTPARTCODE_VER.ToString();
                //ModelCode.ModelCodeStatus.CUSTMODELCODE = clickItem.CUSTMODELCODE.ToString();
                //ModelCode.ModelCodeStatus.REPMODELCODE = clickItem.REPMODELCODE.ToString();
                //ModelCode.ModelCodeStatus.USEFLAG = clickItem.USEFLAG.ToString();
                //ModelCode.ModelCodeStatus.MODELGROUP = clickItem.MODELGROUP.ToString();
                //ModelCode.ModelCodeStatus.MODELDIV = clickItem.MODELDIV.ToString();
                //ModelCode.ModelCodeStatus.MODELTYPE = clickItem.MODELTYPE.ToString();
                //ModelCode.ModelCodeStatus.MODELCHILD = clickItem.MODELCHILD.ToString();
                //ModelCode.ModelCodeStatus.CUST_GB = clickItem.CUST_GB.ToString();
                //ModelCode.ModelCodeStatus.TA = clickItem.TA.ToString();
                //ModelCode.ModelCodeStatus.BUYER = clickItem.BUYER.ToString();
                //ModelCode.ModelCodeStatus.COLOR = clickItem.COLOR.ToString();
                //ModelCode.ModelCodeStatus.MODELSPECIn = clickItem.MODELSPECIn.ToString();
                //ModelCode.ModelCodeStatus.MODELLENGTHIn = clickItem.MODELLENGTHIn.ToString();
                //ModelCode.ModelCodeStatus.MODELWIDTHIn = clickItem.MODELWIDTHIn.ToString();
                //ModelCode.ModelCodeStatus.MODELHEIGHTIn = clickItem.MODELHEIGHTIn.ToString();
                //ModelCode.ModelCodeStatus.MODELUNFOLDLENGTH = clickItem.MODELUNFOLDLENGTH.ToString();
                //ModelCode.ModelCodeStatus.MODELUNFOLDWIDTH = clickItem.MODELUNFOLDWIDTH.ToString();
                //ModelCode.ModelCodeStatus.CUSTCODE = clickItem.CUSTCODE.ToString();
                //ModelCode.ModelCodeStatus.CUSTSHORTCODE = clickItem.CUSTSHORTCODE.ToString();
                //ModelCode.ModelCodeStatus.PAGECNT = clickItem.PAGECNT.ToString();               
                //ModelCode.ModelCodeStatus.SEQ = clickItem.SEQ.ToString();
                //ModelCode.ModelCodeStatus.TYPE = clickItem.TYPE.ToString();
                //ModelCode.ModelCodeStatus.PAPERGUBUNIn = clickItem.PAPERGUBUNIn.ToString();
                //ModelCode.ModelCodeStatus.WEIGHTIn = clickItem.WEIGHTIn.ToString();
                //ModelCode.ModelCodeStatus.WIDTHIn = clickItem.WIDTHIn.ToString();
                //ModelCode.ModelCodeStatus.HEIGHTIn = clickItem.HEIGHTIn.ToString();
                //ModelCode.ModelCodeStatus.SIDEGUBUNIn = clickItem.SIDEGUBUNIn.ToString();
                //ModelCode.ModelCodeStatus.FRONTCCIn = clickItem.FRONTCCIn.ToString();
                //ModelCode.ModelCodeStatus.BACKCCIn = clickItem.BACKCCIn.ToString();
                //ModelCode.ModelCodeStatus.FRONTBCOLORIn = clickItem.FRONTBCOLORIn.ToString();
                //ModelCode.ModelCodeStatus.BACKBCOLORIn = clickItem.BACKBCOLORIn.ToString();
                //ModelCode.ModelCodeStatus.BCOLORCODEIn = clickItem.BCOLORCODEIn.ToString();
                //ModelCode.ModelCodeStatus.PHCOUNTIn = clickItem.PHCOUNTIn.ToString();
                //ModelCode.ModelCodeStatus.SEQ = clickItem.SEQ.ToString();
                //ModelCode.ModelCodeStatus.TYPE = clickItem.TYPE.ToString();
                //ModelCode.ModelCodeStatus.PAPERGUBUNIn = clickItem.PAPERGUBUNIn.ToString();
                //ModelCode.ModelCodeStatus.WEIGHTIn = clickItem.WEIGHTIn.ToString();
                //ModelCode.ModelCodeStatus.WIDTHIn = clickItem.WIDTHIn.ToString();
                //ModelCode.ModelCodeStatus.HEIGHTIn = clickItem.HEIGHTIn.ToString();
                //ModelCode.ModelCodeStatus.SIDEGUBUNIn = clickItem.SIDEGUBUNIn.ToString();
                //ModelCode.ModelCodeStatus.FRONTCCIn = clickItem.FRONTCC.ToString();
                //ModelCode.ModelCodeStatus.BACKCCIn = clickItem.BACKCC.ToString();
                //ModelCode.ModelCodeStatus.FRONTBCOLORIn = clickItem.FRONTBCOLOR.ToString();
                //ModelCode.ModelCodeStatus.BACKBCOLORIn = clickItem.BACKBCOLOR.ToString();
                //ModelCode.ModelCodeStatus.BCOLORCODEIn = clickItem.BCOLORCODE.ToString();
                //ModelCode.ModelCodeStatus.PHCOUNTIn = clickItem.PHCOUNT.ToString();
                //ModelCode.ModelCodeStatus.VERSIONUP = clickItem.VERSIONUP.ToString();
                //ModelCode.ModelCodeStatus.PAPERNAME = clickItem.PAPERNAME.ToString();  
                if(clickItem.MODELCODE != null)
                FilterPaperInBox(clickItem.MODELCODE);
            }
        }

        public string ReadProcessWork(string modelcode)
        {            
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();                    
                    var command = "SELECT b.workcodename_loc FROM tmstmodelwork a left outer JOIN tmstprocwork b on a.workcode = b.workcode WHERE a.modelcode = '"+modelcode+"' ORDER by seq ASC";
                    string process = "";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {                              
                                if (dr[0] != null)
                                {
                                    process += dr[0].ToString() + "->";
                                }
                            }
                        }
                    }
                    conn.Close();
                    return process;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ": DataBase.ReadTable", "ReadTable", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }



        }

        public List<Helper_TaixinDB_Model> ListPaper = new List<Helper_TaixinDB_Model>();
        public void FilterPaperInBox(string strModelCode)
        {
            try
            {

                DataBaseHelper db = new DataBaseHelper();
                ListPaper = db.Read_TaxinDb_ModelCode(MainWindow.path_sql, "TMSTMODEL", "ModelCode", strModelCode);
                int index = 0;
                SamplePaper.PaperIn.CMPCODE = "";
                SamplePaper.PaperIn.BIZDIV = "";
                SamplePaper.PaperIn.MODELCODE = "";
                SamplePaper.PaperIn.MODELNAME = "";
                SamplePaper.PaperIn.APPLYDT = "";
                SamplePaper.PaperIn.VERSION = "";
                SamplePaper.PaperIn.CUSTPARTCODE = "";
                SamplePaper.PaperIn.CUSTPART_VERSION = "";
                SamplePaper.PaperIn.CUSTPARTCODE_VER = "";
                SamplePaper.PaperIn.CUSTMODELCODE = "";
                SamplePaper.PaperIn.REPMODELCODE = "";
                SamplePaper.PaperIn.USEFLAG = "";
                SamplePaper.PaperIn.MODELGROUP = "";
                SamplePaper.PaperIn.MODELDIV = "";
                SamplePaper.PaperIn.MODELTYPE = "";
                SamplePaper.PaperIn.MODELCHILD = "";
                SamplePaper.PaperIn.CUST_GB = "";
                SamplePaper.PaperIn.TA = "";
                SamplePaper.PaperIn.BUYER = "";
                SamplePaper.PaperIn.COLOR = "";
                SamplePaper.PaperIn.MODELSPECIn = "";
                SamplePaper.PaperIn.MODELLENGTHIn = "";
                SamplePaper.PaperIn.MODELWIDTHIn = "";
                SamplePaper.PaperIn.MODELHEIGHTIn = "";
                SamplePaper.PaperIn.MODELUNFOLDLENGTH = "";
                SamplePaper.PaperIn.MODELUNFOLDWIDTH = "";
                SamplePaper.PaperIn.CUSTCODE = "";
                SamplePaper.PaperIn.CUSTSHORTCODE = "";
                SamplePaper.PaperIn.PAGECNT = "";
                SamplePaper.PaperIn.SEQ = "";
                SamplePaper.PaperIn.TYPE = "";
                SamplePaper.PaperIn.PAPERGUBUNIn = "";
                SamplePaper.PaperIn.WEIGHTIn = "";
                SamplePaper.PaperIn.WIDTHIn = "";
                SamplePaper.PaperIn.HEIGHTIn = "";
                SamplePaper.PaperIn.SIDEGUBUNIn = "";
                SamplePaper.PaperIn.FRONTCCIn = "";
                SamplePaper.PaperIn.BACKCCIn = "";
                SamplePaper.PaperIn.FRONTBCOLORIn = "";
                SamplePaper.PaperIn.BACKBCOLORIn = "";
                SamplePaper.PaperIn.BCOLORCODEIn = "";
                SamplePaper.PaperIn.PHCOUNTIn = "";
                SamplePaper.PaperIn.VERSIONUP = "";
                SamplePaper.PaperIn.PAPERNAMEIn = "";
                SamplePaper.PaperIn.PAPERNAME_FullIn = "";
                SamplePaper.PaperIn.TOTALPAGEIN = "";

                SamplePaper.PaperOut.CMPCODE = "";
                SamplePaper.PaperOut.BIZDIV = "";
                SamplePaper.PaperOut.IDNumber = "";
                SamplePaper.PaperOut.MODELCODE = "";
                SamplePaper.PaperOut.MODELNAME = "";
                SamplePaper.PaperOut.APPLYDT = "";
                SamplePaper.PaperOut.VERSION = "";
                SamplePaper.PaperOut.CUSTPARTCODE = "";
                SamplePaper.PaperOut.CUSTPART_VERSION = "";
                SamplePaper.PaperOut.CUSTPARTCODE_VER = "";
                SamplePaper.PaperOut.CUSTMODELCODE = "";
                SamplePaper.PaperOut.REPMODELCODE = "";
                SamplePaper.PaperOut.USEFLAG = "";
                SamplePaper.PaperOut.MODELGROUP = "";
                SamplePaper.PaperOut.MODELDIV = "";
                SamplePaper.PaperOut.MODELTYPE = "";
                SamplePaper.PaperOut.MODELCHILD = "";
                SamplePaper.PaperOut.CUST_GB = "";
                SamplePaper.PaperOut.TA = "";
                SamplePaper.PaperOut.BUYER = "";
                SamplePaper.PaperOut.COLOR = "";
                SamplePaper.PaperOut.MODELSPECOut = "";
                SamplePaper.PaperOut.MODELLENGTHOut = "";
                SamplePaper.PaperOut.MODELWIDTHOut = "";
                SamplePaper.PaperOut.MODELHEIGHTOut = "";
                SamplePaper.PaperOut.MODELUNFOLDLENGTH = "";
                SamplePaper.PaperOut.MODELUNFOLDWIDTH = "";
                SamplePaper.PaperOut.CUSTCODE = "";
                SamplePaper.PaperOut.CUSTSHORTCODE = "";
                SamplePaper.PaperOut.PAGECNT = "";
                SamplePaper.PaperOut.SEQ = "";
                SamplePaper.PaperOut.TYPE = "";
                SamplePaper.PaperOut.PAPERGUBUNOut = "";
                SamplePaper.PaperOut.WEIGHTOut = "";
                SamplePaper.PaperOut.WIDTHOut = "";
                SamplePaper.PaperOut.HEIGHTOut = "";
                SamplePaper.PaperOut.SIDEGUBUNOut = "";
                SamplePaper.PaperOut.FRONTCCOut = "";
                SamplePaper.PaperOut.BACKCCOut = "";
                SamplePaper.PaperOut.FRONTBCOLOROut = "";
                SamplePaper.PaperOut.BACKBCOLOROut = "";
                SamplePaper.PaperOut.BCOLORCODEOut = "";
                SamplePaper.PaperOut.PHCOUNTOut = "";
                SamplePaper.PaperOut.ETC1 = "";
                SamplePaper.PaperOut.TOTALPAGEOUT = "";

                SamplePaper.PaperOut.FOLDINGPAGEOUT = "";
                SamplePaper.PaperOut.TAYOUT = "";
                SamplePaper.PaperOut.TAYPAGEOUT = "";
                SamplePaper.PaperOut.VERSIONUP = "";
                SamplePaper.PaperOut.PAPERNAMEOut = "";
                SamplePaper.PaperOut.INSEMPCODE = "";
                SamplePaper.PaperOut.PAPERNAME_FullOut = "";
                foreach (var clickItem in ListPaper)
                {
                    index++;

                    if (clickItem.SEQ == "1" || clickItem.SEQ=="")
                    {                        
                        SamplePaper.PaperOut.CMPCODE = clickItem.CMPCODE;
                        SamplePaper.PaperOut.BIZDIV = clickItem.BIZDIV;
                        SamplePaper.PaperOut.IDNumber = clickItem.IDNumber;
                        SamplePaper.PaperOut.MODELCODE = clickItem.MODELCODE;
                        SamplePaper.PaperOut.MODELNAME = clickItem.MODELNAME;
                        SamplePaper.PaperOut.APPLYDT = clickItem.APPLYDT;
                        SamplePaper.PaperOut.VERSION = clickItem.VERSION;
                        SamplePaper.PaperOut.CUSTPARTCODE = clickItem.CUSTPARTCODE;
                        SamplePaper.PaperOut.CUSTPART_VERSION = clickItem.CUSTPART_VERSION;
                        SamplePaper.PaperOut.CUSTPARTCODE_VER = clickItem.CUSTPARTCODE_VER;
                        SamplePaper.PaperOut.CUSTMODELCODE = clickItem.CUSTMODELCODE;
                        SamplePaper.PaperOut.REPMODELCODE = clickItem.REPMODELCODE;
                        SamplePaper.PaperOut.USEFLAG = clickItem.USEFLAG;
                        SamplePaper.PaperOut.MODELGROUP = clickItem.MODELGROUP;
                        SamplePaper.PaperOut.MODELDIV = clickItem.MODELDIV;
                        SamplePaper.PaperOut.MODELTYPE = clickItem.MODELTYPE;
                        SamplePaper.PaperOut.MODELCHILD = clickItem.MODELCHILD;
                        SamplePaper.PaperOut.CUST_GB = clickItem.CUST_GB;
                        SamplePaper.PaperOut.TA = clickItem.TA;
                        SamplePaper.PaperOut.BUYER = clickItem.BUYER;
                        SamplePaper.PaperOut.COLOR = clickItem.COLOR;
                        SamplePaper.PaperOut.MODELSPECOut = clickItem.MODELSPECOut;
                        SamplePaper.PaperOut.MODELLENGTHOut = clickItem.MODELLENGTHOut;
                        SamplePaper.PaperOut.MODELWIDTHOut = clickItem.MODELLENGTHOut + "*" + clickItem.MODELWIDTHOut;
                        SamplePaper.PaperOut.MODELHEIGHTOut = clickItem.MODELHEIGHTOut;
                        SamplePaper.PaperOut.MODELUNFOLDLENGTH = clickItem.MODELUNFOLDLENGTH;
                        SamplePaper.PaperOut.MODELUNFOLDWIDTH = clickItem.MODELUNFOLDWIDTH;
                        SamplePaper.PaperOut.CUSTCODE = clickItem.CUSTCODE;
                        SamplePaper.PaperOut.CUSTSHORTCODE = clickItem.CUSTSHORTCODE;
                        SamplePaper.PaperOut.PAGECNT = clickItem.PAGECNT;
                        SamplePaper.PaperOut.SEQ = clickItem.SEQ;
                        SamplePaper.PaperOut.TYPE = clickItem.TYPE;
                        SamplePaper.PaperOut.PAPERGUBUNOut = clickItem.PAPERGUBUNOut;
                        SamplePaper.PaperOut.WEIGHTOut = clickItem.WEIGHTOut;
                        SamplePaper.PaperOut.WIDTHOut = clickItem.WIDTHOut;
                        SamplePaper.PaperOut.HEIGHTOut = clickItem.WIDTHOut + "*" +clickItem.HEIGHTOut;
                        SamplePaper.PaperOut.SIDEGUBUNOut = clickItem.SIDEGUBUNOut;
                        SamplePaper.PaperOut.FRONTCCOut = clickItem.FRONTCCOut;
                        SamplePaper.PaperOut.BACKCCOut = clickItem.BACKCCOut;
                        SamplePaper.PaperOut.FRONTBCOLOROut = clickItem.FRONTBCOLOROut;
                        SamplePaper.PaperOut.BACKBCOLOROut = clickItem.BACKBCOLOROut;
                        SamplePaper.PaperOut.BCOLORCODEOut = clickItem.BCOLORCODEOut;
                        if(clickItem.PHCOUNTOut != "")
                        SamplePaper.PaperOut.PHCOUNTOut = "1*"+clickItem.PHCOUNTOut;
                        string process = ReadProcessWork(clickItem.MODELCODE).ToUpper();
                        if(process != "")
                        SamplePaper.PaperOut.ETC1 = process.Substring(0, process.Length - 2);
                        if(clickItem.FOLDINGPAGEOUT != "")
                        {
                            if (clickItem.TOTALPAGEOUT != null && clickItem.FOLDINGPAGEOUT != null && clickItem.FOLDINGPAGEOUT.Length > 0)
                            {
                                if (int.Parse(clickItem.FOLDINGPAGEOUT) > 0)
                                {
                                    int totalOut = int.Parse(clickItem.TOTALPAGEOUT);
                                    int tayOut = int.Parse(clickItem.FOLDINGPAGEOUT);
                                    SamplePaper.PaperOut.TOTALPAGEOUT = tayOut + "p*" + totalOut / tayOut + "+" + totalOut % tayOut + "p";
                                }
                            }
                            else
                            {
                                SamplePaper.PaperOut.TOTALPAGEOUT = "";
                            }
                        }                          
                        
                        SamplePaper.PaperOut.FOLDINGPAGEOUT = clickItem.FOLDINGPAGEOUT;
                        SamplePaper.PaperOut.TAYOUT = clickItem.TAYOUT;
                        SamplePaper.PaperOut.TAYPAGEOUT = clickItem.TAYPAGEOUT;
                        switch (clickItem.VERSIONUP)
                        {
                            case "01":
                                {
                                    SamplePaper.PaperOut.VERSIONUP = "Thay đổi BOM";
                                    break;
                                }
                            case "02":
                                {
                                    SamplePaper.PaperOut.VERSIONUP = "Thay đổi NVL";
                                    break;
                                }
                            case "03":
                                {
                                    SamplePaper.PaperOut.VERSIONUP = "Thay đổi Spec";
                                    break;
                                }
                            default:
                                {
                                    SamplePaper.PaperOut.VERSIONUP = "Khác";
                                    break;
                                }

                        }
                        if (clickItem.PAPERNAMEOut != "")
                            SamplePaper.PaperOut.PAPERNAMEOut = clickItem.PAPERNAMEOut + " " + clickItem.WEIGHTOut + "g";
                        SamplePaper.PaperOut.INSEMPCODE = MainWindow.UserLogin;
                        SamplePaper.PaperOut.PAPERNAME_FullOut = clickItem.PAPERNAME_FullOut;
                    }
                    else if (clickItem.SEQ == "2")
                    {
                        SamplePaper.PaperIn.CMPCODE = clickItem.CMPCODE;
                        SamplePaper.PaperIn.BIZDIV = clickItem.BIZDIV;
                        SamplePaper.PaperIn.MODELCODE = clickItem.MODELCODE;
                        SamplePaper.PaperIn.MODELNAME = clickItem.MODELNAME;
                        SamplePaper.PaperIn.APPLYDT = clickItem.APPLYDT;
                        SamplePaper.PaperIn.VERSION = clickItem.VERSION;
                        SamplePaper.PaperIn.CUSTPARTCODE = clickItem.CUSTPARTCODE;
                        SamplePaper.PaperIn.CUSTPART_VERSION = clickItem.CUSTPART_VERSION;
                        SamplePaper.PaperIn.CUSTPARTCODE_VER = clickItem.CUSTPARTCODE_VER;
                        SamplePaper.PaperIn.CUSTMODELCODE = clickItem.CUSTMODELCODE;
                        SamplePaper.PaperIn.REPMODELCODE = clickItem.REPMODELCODE;
                        SamplePaper.PaperIn.USEFLAG = clickItem.USEFLAG;
                        SamplePaper.PaperIn.MODELGROUP = clickItem.MODELGROUP;
                        SamplePaper.PaperIn.MODELDIV = clickItem.MODELDIV;
                        SamplePaper.PaperIn.MODELTYPE = clickItem.MODELTYPE;
                        SamplePaper.PaperIn.MODELCHILD = clickItem.MODELCHILD;
                        SamplePaper.PaperIn.CUST_GB = clickItem.CUST_GB;
                        SamplePaper.PaperIn.TA = clickItem.TA;
                        SamplePaper.PaperIn.BUYER = clickItem.BUYER;
                        SamplePaper.PaperIn.COLOR = clickItem.COLOR;
                        SamplePaper.PaperIn.MODELSPECIn = clickItem.MODELSPECIn;
                        SamplePaper.PaperIn.MODELLENGTHIn = clickItem.MODELLENGTHIn;
                        SamplePaper.PaperIn.MODELWIDTHIn = clickItem.MODELWIDTHIn;
                        SamplePaper.PaperIn.MODELHEIGHTIn = clickItem.MODELHEIGHTIn;
                        SamplePaper.PaperIn.MODELUNFOLDLENGTH = clickItem.MODELUNFOLDLENGTH;
                        SamplePaper.PaperIn.MODELUNFOLDWIDTH = clickItem.MODELUNFOLDWIDTH;
                        SamplePaper.PaperIn.CUSTCODE = clickItem.CUSTCODE;
                        SamplePaper.PaperIn.CUSTSHORTCODE = clickItem.CUSTSHORTCODE;
                        SamplePaper.PaperIn.PAGECNT = clickItem.PAGECNT;
                        SamplePaper.PaperIn.SEQ = clickItem.SEQ;
                        SamplePaper.PaperIn.TYPE = clickItem.TYPE;
                        SamplePaper.PaperIn.PAPERGUBUNIn = clickItem.PAPERGUBUNIn;
                        SamplePaper.PaperIn.WEIGHTIn = clickItem.WEIGHTIn;
                        SamplePaper.PaperIn.WIDTHIn = clickItem.WIDTHIn;
                        SamplePaper.PaperIn.HEIGHTIn = clickItem.WIDTHIn + "*" +clickItem.HEIGHTIn;
                        SamplePaper.PaperIn.SIDEGUBUNIn = clickItem.SIDEGUBUNIn;
                        SamplePaper.PaperIn.FRONTCCIn = clickItem.FRONTCCIn;
                        SamplePaper.PaperIn.BACKCCIn = clickItem.BACKCCIn;
                        SamplePaper.PaperIn.FRONTBCOLORIn = clickItem.FRONTBCOLORIn;
                        SamplePaper.PaperIn.BACKBCOLORIn = clickItem.BACKBCOLORIn;
                        SamplePaper.PaperIn.BCOLORCODEIn = clickItem.BCOLORCODEIn;                        
                        SamplePaper.PaperIn.PHCOUNTIn = "1*"+clickItem.PHCOUNTIn;
                        if(clickItem.FOLDINGPAGEOUT != "")
                        {
                            if (clickItem.TOTALPAGEIN != null && clickItem.FOLDINGPAGEIN != null && clickItem.FOLDINGPAGEIN.Length > 0)
                            {
                                if (int.Parse(clickItem.FOLDINGPAGEIN) > 0)
                                {
                                    int totalIn = int.Parse(clickItem.TOTALPAGEIN);
                                    int tayIn = int.Parse(clickItem.FOLDINGPAGEIN);
                                    SamplePaper.PaperIn.TOTALPAGEIN = tayIn + "p*" + totalIn / tayIn + "+" + totalIn % tayIn + "p";
                                }
                            }
                            else
                            {
                                SamplePaper.PaperIn.TOTALPAGEIN = "";
                            }
                        }    
                            
                        SamplePaper.PaperIn.FOLDINGPAGEIN = clickItem.FOLDINGPAGEIN;
                        SamplePaper.PaperIn.TAYIN = clickItem.TAYIN;
                        SamplePaper.PaperIn.TAYPAGEIN = clickItem.TAYPAGEIN;
                        SamplePaper.PaperIn.VERSIONUP = clickItem.VERSIONUP;
                        if(clickItem.PAPERNAMEIn !="")
                        SamplePaper.PaperIn.PAPERNAMEIn= clickItem.PAPERNAMEIn + " " + clickItem.WEIGHTIn + "g";
                        SamplePaper.PaperIn.INSEMPCODE = clickItem.INSEMPCODE;
                        SamplePaper.PaperIn.PAPERNAME_FullIn = clickItem.PAPERNAME_FullIn;
                    }
                }
                if (index < 2)
                {
                    SamplePaper.PaperIn.CMPCODE = "";
                    SamplePaper.PaperIn.BIZDIV = "";
                    SamplePaper.PaperIn.MODELCODE = "";
                    SamplePaper.PaperIn.MODELNAME = "";
                    SamplePaper.PaperIn.APPLYDT = "";
                    SamplePaper.PaperIn.VERSION = "";
                    SamplePaper.PaperIn.CUSTPARTCODE = "";
                    SamplePaper.PaperIn.CUSTPART_VERSION = "";
                    SamplePaper.PaperIn.CUSTPARTCODE_VER = "";
                    SamplePaper.PaperIn.CUSTMODELCODE = "";
                    SamplePaper.PaperIn.REPMODELCODE = "";
                    SamplePaper.PaperIn.USEFLAG = "";
                    SamplePaper.PaperIn.MODELGROUP = "";
                    SamplePaper.PaperIn.MODELDIV = "";
                    SamplePaper.PaperIn.MODELTYPE = "";
                    SamplePaper.PaperIn.MODELCHILD = "";
                    SamplePaper.PaperIn.CUST_GB = "";
                    SamplePaper.PaperIn.TA = "";
                    SamplePaper.PaperIn.BUYER = "";
                    SamplePaper.PaperIn.COLOR = "";
                    SamplePaper.PaperIn.MODELSPECIn = "";
                    SamplePaper.PaperIn.MODELLENGTHIn = "";
                    SamplePaper.PaperIn.MODELWIDTHIn = "";
                    SamplePaper.PaperIn.MODELHEIGHTIn = "";
                    SamplePaper.PaperIn.MODELUNFOLDLENGTH = "";
                    SamplePaper.PaperIn.MODELUNFOLDWIDTH = "";
                    SamplePaper.PaperIn.CUSTCODE = "";
                    SamplePaper.PaperIn.CUSTSHORTCODE = "";
                    SamplePaper.PaperIn.PAGECNT = "";
                    SamplePaper.PaperIn.SEQ = "";
                    SamplePaper.PaperIn.TYPE = "";
                    SamplePaper.PaperIn.PAPERGUBUNIn = "";
                    SamplePaper.PaperIn.WEIGHTIn = "";
                    SamplePaper.PaperIn.WIDTHIn = "";
                    SamplePaper.PaperIn.HEIGHTIn = "";
                    SamplePaper.PaperIn.SIDEGUBUNIn = "";
                    SamplePaper.PaperIn.FRONTCCIn = "";
                    SamplePaper.PaperIn.BACKCCIn = "";
                    SamplePaper.PaperIn.FRONTBCOLORIn = "";
                    SamplePaper.PaperIn.BACKBCOLORIn = "";
                    SamplePaper.PaperIn.BCOLORCODEIn = "";
                    SamplePaper.PaperIn.PHCOUNTIn = "";                    
                    SamplePaper.PaperIn.VERSIONUP = "";
                    SamplePaper.PaperIn.PAPERNAMEIn = "";
                    SamplePaper.PaperIn.PAPERNAME_FullIn = "";
                    SamplePaper.PaperIn.TOTALPAGEIN = "";
                    index = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ListCustomerPartCode :" + "\r\n" + ex.Message, "List Customer Infomation", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
            {
                this.Close();
            }
        }
    }
    
}
