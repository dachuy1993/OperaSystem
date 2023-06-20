using DataHelper;
using Microsoft.Win32;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.Style;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Windows.Xps.Packaging;
using Tulpep.NotificationWindow;
using W_Opera.DAO;
using ZXing.QrCode.Internal;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Page_Sample_Manual.xaml
    /// </summary>
    public partial class Page_Sample_Manual : Page
    {
        #region Khai báo
        string path_sql_attach = "Data Source=192.168.2.10;Initial Catalog=taixin_attach;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        string path_sql = "Data Source=192.168.2.10;Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        List<Helper_TaixinDB_Model> ListAllDataSample = new List<Helper_TaixinDB_Model>();
        List<Helper_DataButton> listButtonTop = new List<Helper_DataButton>();
        Helper_TaixinDB_Model ApprovalClickItem = new Helper_TaixinDB_Model();
        Helper_AccessManger access_db = new Helper_AccessManger();
        List<Helper_AccessManger> list_access = new List<Helper_AccessManger>();
        List<Helper_DataExcel> list_Excel = new List<Helper_DataExcel>();
        PopupNotifier popup = new PopupNotifier();
        DispatcherTimer dt = new DispatcherTimer();
        DataBaseHelper db = new DataBaseHelper();
        string pathFileExcel = @"TempFile//ExcelFile.xlsx";

        public static string CustomerCode = "";

        string date = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");
        string str_cbbFilterApprove = "Tìm kiếm All";
        string dateFilterStart = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyy-MM-dd") + " 00:00:00";
        string dateFilterFinish = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyy-MM-dd") + " 23:59:59";   
        bool checkPopup = false;
        bool checkView = false;
        string processButton = "";
        public string str_depCreate = "";
        string str_FilterProduct = "Manual";
        string qtyAttachFile = "0";
        string accessProcess = "Approve";
        OpenFileDialog ofd;
        int qtyFileUpload = 0;
        bool checkDowload = false;
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
        string str_PAGETYPEM;
        string str_PAGETYPED;
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
        string str_qtyAttach;

        #endregion

        public Page_Sample_Manual()
        {
            InitializeComponent();
            CreatAllButtonEdit();
            Loaded += Page_Sample_Loaded;
        }      
        private void Page_Sample_Loaded(object sender, RoutedEventArgs e)
        {
            path_sql = "Data Source=" + MainWindow.ip + ";Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
            path_sql_attach = MainWindow.path_sql_attach;
            dpkStartApprove.SelectedDate = DateTime.Now;
            dpkFinishApprove.SelectedDate = DateTime.Now.AddDays(1);
            dp_DateStart.SelectedDate = DateTime.Now;
            dp_DateFinish.SelectedDate = DateTime.Now;
            rb_Manual.IsChecked = true;
            stackBox.Visibility = Visibility.Hidden;
            stackManul.Visibility = Visibility.Visible;
            str_APPLYDT = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");
            GetDataDeptUser();
            AccessManage();                    
            ColorRowListView(Filter_Sample_All());
            


        }

        public void GetDataDeptUser()
        {
            try
            {
                
                List<string> list = new List<string>();
                using (SqlConnection conn = new SqlConnection(MainWindow.path_sql))
                {
                    conn.Open();
                    {
                        var command = "SELECT Department FROM tbSampleAccess where UserLogin = '" + MainWindow.UserLogin + "'";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            using (IDataReader dr = cmd.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    list.Add(dr[0].ToString());
                                    if (dr[0] != null)
                                    {
                                        str_depCreate = dr[0].ToString();
                                        //if (txt_User.Text.ToUpper() == dr[0].ToString().Trim().ToUpper() && (txtPass.Text.ToUpper() == dr[1].ToString().Trim().ToUpper() || pb_Pass.Password.ToUpper() == dr[1].ToString().Trim().ToUpper()))
                                        //{
                                        //    checkLogin = true;
                                        //}
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


        

        public void CreatAllButtonEdit()
        {
            lvButtonTop.Items.Clear();
            listButtonTop.Clear();
            listButtonTop.Add(new Helper_DataButton
            {
                ID = 1,
                ContentButton = "Add",
                ImageSource = "Image/Edit/add.png",
                BackGroundColor = PinValue.OFF
            });
            listButtonTop.Add(new Helper_DataButton
            {
                ID = 2,
                ContentButton = "Del",
                ImageSource = "Image/Edit/delete.png",
                BackGroundColor = PinValue.OFF
            });
            listButtonTop.Add(new Helper_DataButton
            {
                ID = 3,
                ContentButton = "Edit",
                ImageSource = "Image/Edit/edit.png",
                BackGroundColor = PinValue.OFF
            });
            listButtonTop.Add(new Helper_DataButton
            {
                ID = 4,
                ContentButton = "Save",
                ImageSource = "Image/Edit/save.png",
                BackGroundColor = PinValue.OFF
            });
            listButtonTop.Add(new Helper_DataButton
            {
                ID = 5,
                ContentButton = "Print",
                ImageSource = "Image/Edit/printer.png",
                BackGroundColor = PinValue.OFF
            });
            foreach (var button in listButtonTop)
            {
                lvButtonTop.Items.Add(button);
            }

        }
        private void ButtonTop_Click(object sender, RoutedEventArgs e)
        {
            var click = sender as Button;
            var clickItem = click.DataContext as Helper_DataButton;
            if (clickItem != null)
            {
                switch (clickItem.ContentButton)
                {
                    case "Add":
                        {
                            processButton = "Add";
                            ProcessButtonEdit_Add();
                            break;
                        }
                    case "Del":
                        {
                            processButton = "Del";
                            ProcessButtonEdit_Del();
                            break;
                        }
                    case "Edit":
                        {
                            processButton = "Edit";
                            ProcessButtonEdit_Edit();
                            break;
                        }
                    case "Save":
                        {
                            processButton = "Save";
                            ProcessButtonEdit_Save();
                            break;
                        }
                    case "Print":
                        {
                            processButton = "Print";
                            ProcessButtonEdit_Printer();
                            break;
                        }
                    case "Run":
                        {
                            ProcessButtonEdit_Run();
                            break;
                        }

                }
                foreach (var button in listButtonTop)
                {
                    button.BackGroundColor = PinValue.OFF;
                    if (button.ContentButton == clickItem.ContentButton)
                    {
                        button.BackGroundColor = PinValue.ON;
                    }
                }

            }
        }
        public string ReadSampleNo()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "SELECT max(samno) FROM tbSampleManual Where applydt='" + date + "'";
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        string SampleNoMax = cmd.ExecuteScalar().ToString();
                        string samNo;
                        if (SampleNoMax != "")
                        {
                            string noMax = SampleNoMax.Substring(9, 4);
                            samNo = "SA" + date.Substring(2, date.Length - 2) + "-" + (int.Parse(noMax) + 1).ToString("0000");
                        }
                        else
                        {
                            samNo = "SA" + date.Substring(2, date.Length - 2) + "-" + ("0001");
                        }
                        conn.Close();
                        return samNo;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ReadSampleNo", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }
        public static List<Helper_TaixinDB_Model> listSampleManual = new List<Helper_TaixinDB_Model>();
        public List<Helper_TaixinDB_Model> ReadDataBase(string command)
        {
            try
            {
                listSampleManual.Clear();
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();                   
                    
                    //List<Helper_TaixinDB_SampleBox> list_print = db.Read_TaxinDb_SamplePrint(path_sql, "Manual");
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_TaixinDB_Model _sample = new Helper_TaixinDB_Model();

                                _sample.CMPCODE = dr[0].ToString();
                                _sample.BIZDIV = dr[1].ToString();
                                _sample.IDNumber = dr[2].ToString();
                                _sample.MODELCODE = dr[3].ToString();
                                _sample.MODELNAME = dr[4].ToString();
                                _sample.APPLYDT = dr[5].ToString();
                                _sample.VERSION = dr[6].ToString();
                                _sample.CUSTPARTCODE = dr[7].ToString();
                                _sample.CUSTPART_VERSION = dr[8].ToString();
                                _sample.CUSTPARTCODE_VER = dr[9].ToString();
                                _sample.CUSTMODELCODE = dr[10].ToString();
                                _sample.REPMODELCODE = dr[11].ToString();
                                _sample.USEFLAG = dr[12].ToString();
                                _sample.MODELGROUP = dr[13].ToString();
                                _sample.MODELDIV = dr[14].ToString();
                                _sample.MODELTYPE = dr[15].ToString();
                                _sample.MODELCHILD = dr[16].ToString();
                                _sample.CUST_GB = dr[17].ToString();
                                _sample.TA = dr[18].ToString();
                                _sample.BUYER = dr[19].ToString();
                                _sample.COLOR = dr[20].ToString();
                                _sample.MODELSPECIn = dr[21].ToString();
                                _sample.MODELLENGTHIn = dr[22].ToString();
                                _sample.MODELWIDTHIn = dr[23].ToString();
                                _sample.MODELHEIGHTIn = dr[24].ToString();
                                _sample.MODELSPECOut = dr[25].ToString();
                                _sample.MODELLENGTHOut = dr[26].ToString();
                                _sample.MODELWIDTHOut = dr[27].ToString();
                                _sample.MODELHEIGHTOut = dr[28].ToString();
                                _sample.MODELUNFOLDLENGTH = dr[29].ToString();
                                _sample.MODELUNFOLDWIDTH = dr[30].ToString();
                                _sample.CUSTCODE = dr[31].ToString();
                                _sample.CUSTSHORTCODE = dr[32].ToString();
                                _sample.PAGECNT = dr[33].ToString();
                                _sample.PAGETYPEM = dr[101].ToString();
                                _sample.PAGETYPED = dr[102].ToString();
                                _sample.SEQ = dr[34].ToString();
                                _sample.TYPE = dr[35].ToString();
                                _sample.PAPERGUBUNIn = dr[36].ToString();
                                _sample.WEIGHTIn = dr[37].ToString();
                                _sample.WIDTHIn = dr[38].ToString();
                                _sample.HEIGHTIn = dr[39].ToString();
                                _sample.SIDEGUBUNIn = dr[40].ToString();
                                _sample.FRONTCCIn = dr[41].ToString();
                                _sample.BACKCCIn = dr[42].ToString();
                                _sample.FRONTBCOLORIn = dr[43].ToString();
                                _sample.BACKBCOLORIn = dr[44].ToString();
                                _sample.BCOLORCODEIn = dr[45].ToString();
                                _sample.PHCOUNTIn = dr[46].ToString();
                                _sample.PAPERGUBUNOut = dr[47].ToString();
                                _sample.WEIGHTOut = dr[48].ToString();
                                _sample.WIDTHOut = dr[49].ToString();
                                _sample.HEIGHTOut = dr[50].ToString();
                                _sample.SIDEGUBUNOut = dr[51].ToString();
                                _sample.FRONTCCOut = dr[52].ToString();
                                _sample.BACKCCOut = dr[53].ToString();
                                _sample.FRONTBCOLOROut = dr[54].ToString();
                                _sample.BACKBCOLOROut = dr[55].ToString();
                                _sample.BCOLORCODEOut = dr[56].ToString();
                                _sample.PHCOUNTOut = dr[57].ToString();
                                _sample.VERSIONUP = dr[58].ToString();
                                _sample.PAPERNAMEIn = dr[59].ToString();
                                _sample.PAPERNAME_FullIn = dr[60].ToString();
                                _sample.PAPERNAMEOut = dr[61].ToString();
                                _sample.PAPERNAME_FullOut = dr[62].ToString();
                                //_sample.MA = dr[63].ToString();
                                //_sample.SX = dr[64].ToString();
                                //_sample.CL = dr[65].ToString();
                                //_sample.RD = dr[66].ToString();
                                //_sample.KD = dr[67].ToString();
                                _sample.TOTALPAGEIN = dr[68].ToString();
                                _sample.FOLDINGPAGEIN = dr[69].ToString();
                                _sample.TAYIN = dr[70].ToString();
                                _sample.TAYPAGEIN = dr[71].ToString();
                                _sample.TOTALPAGEOUT = dr[72].ToString();
                                _sample.FOLDINGPAGEOUT = dr[73].ToString();
                                _sample.TAYOUT = dr[74].ToString();
                                _sample.TAYPAGEOUT = dr[75].ToString();
                                _sample.QTYREQUEST = dr[76].ToString();
                                _sample.NOTE2 = dr[77].ToString();
                                _sample.NOTE1 = dr[78].ToString();
                                _sample.ETC1 = dr[79].ToString();
                                _sample.ETC3 = dr[80].ToString();                               
                                _sample.reject = dr[81].ToString();
                                _sample.depCreat = dr[82].ToString();                               
                                _sample.printed = dr[83].ToString();
                                _sample.DATESTARTAPPROVE = dr[93].ToString();
                                _sample.DATEFINISHAPPROVE = dr[94].ToString();
                                _sample.INSEMPCODE = dr[95].ToString();
                                _sample.DATEINPUT = dr[96].ToString();
                                _sample.INSEMPCODEUP = dr[97].ToString();
                                _sample.UPDT = dr[98].ToString();

                                if (dr[63].ToString() == "Y")
                                {
                                    _sample.MA = "DodgerBlue";
                                }
                                else if (dr[63].ToString() == "N")
                                {
                                    _sample.MA = "Red";
                                }
                                else
                                {
                                    _sample.MA = "LightGray";
                                }
                                //7
                                if (dr[64].ToString() == "Y")
                                {
                                    _sample.SX = "DodgerBlue";
                                }
                                else if (dr[64].ToString() == "N")
                                {
                                    _sample.SX = "Red";
                                }
                                else
                                {
                                    _sample.SX = "LightGray";
                                }
                                //8
                                if (dr[65].ToString() == "Y")
                                {
                                    _sample.CL = "DodgerBlue";
                                }
                                else if (dr[65].ToString() == "N")
                                {
                                    _sample.CL = "Red";
                                }
                                else
                                {
                                    _sample.CL = "LightGray";
                                }
                                //9
                                if (dr[66].ToString() == "Y")
                                {
                                    _sample.RD = "DodgerBlue";
                                }
                                else if (dr[66].ToString() == "N")
                                {
                                    _sample.RD = "Red";
                                }
                                else
                                {
                                    _sample.RD = "LightGray";
                                }
                                //10
                                if (dr[67].ToString() == "Y")
                                {
                                    _sample.KD = "DodgerBlue";
                                }
                                else if (dr[67].ToString() == "N")
                                {
                                    _sample.KD = "Red";
                                }
                                else
                                {
                                    _sample.KD = "LightGray";
                                }
                                //_sample.NumberCheck = int.Parse(dr[11].ToString());
                                listSampleManual.Add(_sample);
                            }
                        }
                        conn.Close();
                        return listSampleManual;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ReadDataBase", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_Model> Filter_Sample_OK()
        {
            var command = "SELECT * from tbSampleManual where (dep_MAR = 'Y' or dep_MAR='O') and ( dep_PRO='Y' or dep_PRO = 'O') and " +
                "(dep_QC='Y' or dep_QC ='O') and (dep_RND='Y' or dep_RND ='O') and (dep_PUR = 'Y'  or dep_PUR ='O') and reject ='0' ORDER BY Insdt DESC";
            return ReadDataBase(command);
        }

        public List<Helper_TaixinDB_Model> Filter_Sample_NG()
        {
            var command = "SELECT * from tbSampleManual where (dep_MAR = 'N' or dep_PRO='N' or dep_QC='N' or dep_RND='N' or dep_PUR='N') and reject ='0' ORDER BY Insdt DESC";
            return ReadDataBase(command);
        }

        public List<Helper_TaixinDB_Model> Filter_Sample_All()
        {
            var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
            var jsonDateFilterStart = JsonConvert.SerializeObject(DateTime.Parse(dateFilterStart), settings);
            string FilterStart = jsonDateFilterStart.Substring(1, jsonDateFilterStart.Length - 2);
            var jsonDateFilterFinish = JsonConvert.SerializeObject(DateTime.Parse(dateFilterFinish), settings);
            string FilterFinish = jsonDateFilterFinish.Substring(1, jsonDateFilterFinish.Length - 2);
            var command = "SELECT * from tbSampleManual where INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "' ORDER by Insdt desc";
            return ReadDataBase(command);
        }

        public List<Helper_TaixinDB_Model> Filter_Sample_RE()
        {
            var command = "SELECT * from tbSampleManual where reject='1' ORDER BY Insdt DESC";
            return ReadDataBase(command);
        }

        public List<Helper_TaixinDB_Model> Filter_Run_OK()
        {
            var command = "SELECT * from tbSampleManual where printed='Y' ORDER BY Insdt DESC";
            return ReadDataBase(command);
        }
        public List<Helper_TaixinDB_Model> Filter_Run_NG()
        {
            var command = "SELECT * from tbSampleManual where printed is null ORDER BY Insdt DESC";
            return ReadDataBase(command);
        }

        public void Add_NewData()
        {
            try
            {
                var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                string dateInput = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                str_CMPCODE = "02";
                str_BIZDIV = "300";
                str_SAMNO = ReadSampleNo();
                str_MODELCODE = SamplePaper.PaperOut.MODELCODE;
                str_MODELNAME = SamplePaper.PaperOut.MODELNAME;
                str_CUSTPARTCODE_VER = SamplePaper.PaperOut.CUSTPARTCODE_VER;
                str_REPMODELCODE = SamplePaper.PaperOut.REPMODELCODE;
                str_USEFLAG = SamplePaper.PaperOut.USEFLAG;
                str_MODELGROUP = SamplePaper.PaperOut.MODELGROUP;
                str_MODELDIV = SamplePaper.PaperOut.MODELDIV;
                str_MODELTYPE = SamplePaper.PaperOut.TYPE;
                str_MODELCHILD = SamplePaper.PaperOut.MODELCHILD;
                str_TA = SamplePaper.PaperOut.TA;
                str_BUYER = SamplePaper.PaperOut.BUYER;
                str_MODELSPECIn = SamplePaper.PaperIn.MODELSPECIn;
                str_MODELLENGTHIn = SamplePaper.PaperIn.MODELLENGTHIn;
                str_MODELWIDTHIn = SamplePaper.PaperIn.WIDTHIn;

                str_MODELUNFOLDLENGTHIn = SamplePaper.PaperIn.MODELUNFOLDLENGTH;
                str_MODELUNFOLDWIDTHIn = SamplePaper.PaperIn.MODELUNFOLDWIDTH;

                str_COLOROut = SamplePaper.PaperOut.COLOR;

                str_MODELLENGTHOut = SamplePaper.PaperOut.MODELLENGTHOut;
                str_MODELWIDTHOut = SamplePaper.PaperOut.WIDTHOut;

                str_MODELUNFOLDLENGTHOut = SamplePaper.PaperOut.MODELUNFOLDLENGTH;
                str_MODELUNFOLDWIDTHOut = SamplePaper.PaperOut.MODELUNFOLDWIDTH;

                str_CUSTCODE = SamplePaper.PaperOut.CUSTCODE;
                str_CUSTSHORTCODE = SamplePaper.PaperOut.CUSTSHORTCODE;

                str_TYPE = SamplePaper.PaperOut.TYPE;
                str_PAPERGUBUNIn = SamplePaper.PaperOut.PAPERGUBUNIn;
                str_PAPERGUBUNOut = SamplePaper.PaperOut.PAPERGUBUNOut;
                str_WEIGHTIn = SamplePaper.PaperIn.WEIGHTIn;
                str_WIDTHIn = SamplePaper.PaperIn.WIDTHIn;
                str_HEIGHTIn = SamplePaper.PaperIn.HEIGHTIn;
                str_SIDEGUBUNIn = SamplePaper.PaperIn.SIDEGUBUNIn;
                str_FRONTCCIn = SamplePaper.PaperIn.FRONTCCIn;
                str_BACKCCIn = SamplePaper.PaperIn.BACKCCIn;
                str_FRONTBCOLORIn = SamplePaper.PaperIn.FRONTBCOLORIn;
                str_BACKBCOLORIn = SamplePaper.PaperIn.BACKBCOLORIn;

                str_FOLDINGPAGEIN = SamplePaper.PaperIn.FOLDINGPAGEIN;
                str_TAYIN = SamplePaper.PaperIn.TAYIN;
                str_TAYPAGEIN = SamplePaper.PaperIn.TAYPAGEIN;

                str_INSEMPCODE = txt_INSEMPCODE.Text;

                str_WEIGHTOut = SamplePaper.PaperOut.WEIGHTOut;
                str_WIDTHOut = SamplePaper.PaperOut.WIDTHOut;
                str_HEIGHTOut = SamplePaper.PaperOut.HEIGHTOut;
                str_SIDEGUBUNOut = SamplePaper.PaperOut.SIDEGUBUNOut;
                str_FRONTCCOut = SamplePaper.PaperOut.FRONTCCOut;
                str_BACKCCOut = SamplePaper.PaperOut.BACKCCOut;
                str_FRONTBCOLOROut = SamplePaper.PaperOut.FRONTBCOLOROut;
                str_BACKBCOLOROut = SamplePaper.PaperOut.BACKBCOLOROut;

                str_FOLDINGPAGEOUT = SamplePaper.PaperOut.FOLDINGPAGEOUT;
                str_TAYOUT = SamplePaper.PaperOut.TAYOUT;
                str_TAYPAGEOUT = SamplePaper.PaperOut.TAYPAGEOUT;

                str_PAPERNAMEIn = SamplePaper.PaperIn.PAPERNAMEIn;
                str_PAPERNAMEOut = SamplePaper.PaperOut.PAPERNAMEOut;
                str_PAPERNAMEFULLIn = SamplePaper.PaperIn.PAPERNAME_FullIn;
                str_PAPERNAMEFULLOut = SamplePaper.PaperOut.PAPERNAME_FullOut;


                txb_NameFileUpload.Text = "";
                str_CUSTMODELCODE = txt_CUSTMODELCODE.Text;
                str_CUSTPARTCODE = txt_CUSTPARTCODE.Text;
                str_CUST_GB = txt_CUST_GB.Text;
                str_VERSION = txt_VERSION.Text;
                str_APPLYDT = date;
                str_CUSTPART_VERSION = txt_VERSION.Text;
                str_COLORIn = txt_PaperIn.Text;
                str_COLOROut = txt_PaperOut.Text;
                str_MODELSPECOut = txt_FullSize.Text;
                str_MODELHEIGHTIn = txt_SizeOut.Text;
                str_MODELHEIGHTOut = txt_SizeIn.Text;
                str_PAGECNT = txt_PaperNumber.Text;
                str_PAGETYPEM = cbbTypeCertification.Text;
                str_PAGETYPED = cbbTypeDetail.Text;
                str_BCOLORCODEIn = txt_ColorOut.Text;
                str_PHCOUNTIn = txt_RatioOut.Text;
                str_TOTALPAGEIN = txt_foldIn.Text;
                str_VERSIONUP = txt_VERSIONUP.Text;
                str_BCOLORCODEOut = txt_ColorIn.Text;
                str_PHCOUNTOut = txt_RatioIn.Text;
                str_TOTALPAGEOUT = txt_foldOut.Text;
                str_NOTE1 = txt_Note1.Text;
                str_NOTE2 = txt_Note2.Text;
                str_PROCESSWORDK = txt_ProcessWork.Text;
                str_QTYREQUEST = txt_Request.Text;
                str_DATEAPPROVE = dateInput;
                str_DATEINPUT = dateInput;
                dateApproval = dateInput;
                str_qtyAttach = qtyFileUpload.ToString();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Manual/Add_NewData", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            


        }

        public void ProcessButtonEdit_Add()
        {
            try
            {
                SamplePaper.PaperOut.CMPCODE = "";
                SamplePaper.PaperOut.BIZDIV = "";
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
                SamplePaper.PaperOut.PAGETYPEM = "";
                SamplePaper.PaperOut.PAGETYPED = "";
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
                SamplePaper.PaperOut.VERSIONUP = "";
                SamplePaper.PaperOut.PAPERNAMEOut = "";
                SamplePaper.PaperOut.PAPERNAME_FullOut = "";
                SamplePaper.PaperOut.TOTALPAGEOUT = "";
                SamplePaper.PaperOut.NOTE1 = "";
                SamplePaper.PaperOut.NOTE2 = "";
                SamplePaper.PaperOut.QTYREQUEST = "";
                SamplePaper.PaperOut.ETC1 = "";

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
                SamplePaper.PaperIn.PAGETYPEM = "";
                SamplePaper.PaperIn.PAGETYPED = "";
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

                //txt_CUSTMODELCODE.Text = "";
                //txt_CUSTPARTCODE.Text = "";
                //txt_VERSIONUP.Text = "";
                //txt_VERSION.Text = "";
                //txt_CUST_GB.Text = "";
                //txt_INSEMPCODE.Text = "";
                //txt_PaperIn.Text = "";
                //txt_SizeIn.Text = "";
                //txt_RatioIn.Text = "";
                //txt_PaperNameIn.Text = "";
                //txt_PaperOut.Text = "";
                //txt_SizeOut.Text = "";
                //txt_RatioOut.Text = "";
                //txt_PaperNameOut.Text = "";
                //txt_FullSize.Text = "";
                //txt_ColorIn.Text = "";
                //txt_ColorOut.Text = "";
                //txt_PaperNumber.Text = "";
                //txt_Note1.Text = "";
                //txt_Note2.Text = "";
                //txt_Request.Text = "";
                //txt_foldIn.Text = "";
                //txt_foldOut.Text = "";
                SamplePaper.PaperOut.ETC3 = "0 File";
                MainWindow.at_Samno = "";
                ckbApprove_Ma.IsChecked = false;
                ckbApprove_SX.IsChecked = false;
                ckbApprove_CL.IsChecked = false;
                ckbApprove_RD.IsChecked = false;
                ckbApprove_KD.IsChecked = false;
                dpkStartApprove.SelectedDate = DateTime.Now;
                dpkFinishApprove.SelectedDate = DateTime.Now.AddDays(1);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Manual/ProcessButtonEdit_Add", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            

        }

        public void ProcessSampleHistory(string samno,string ope)
        {           
            using (SqlConnection conn = new SqlConnection(path_sql))
            {
                conn.Open();
                var command = "Insert tbSampleHistory(cmpcode,bizdiv,samno,modelcode,typeSample,applydt,operator,imsempcode,insdt) values( '02','300','" + samno + "','" + txt_CUSTMODELCODE.Text + "','Manual','"+str_APPLYDT+"','" + ope + "','"+MainWindow.UserLogin+"','"+str_DATEINPUT+"')";
                using (SqlCommand cmd = new SqlCommand(command, conn))
                {
                    cmd.CommandTimeout = 100;
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        public void ProcessButtonEdit_Del()
        {
            try
            {
                if (MessageBoxResult.Yes == MessageBox.Show("Bạn muốn xóa dữ liệu?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question))
                {
                    using (SqlConnection conn = new SqlConnection(path_sql))
                    {
                        conn.Open();
                        var command = "DELETE tbSampleManual WHERE SAMNO =" + "'" + ApprovalClickItem.IDNumber + "'";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            cmd.CommandTimeout = 100;
                            cmd.ExecuteNonQuery();
                        }                        
                        conn.Close();
                    }

                    using (SqlConnection conn = new SqlConnection(path_sql_attach))
                    {
                        conn.Open();
                        var command = "DELETE tbSampleAttach WHERE typeSample ='Manual' and SAMNO =" + "'" + ApprovalClickItem.IDNumber + "'";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            cmd.CommandTimeout = 100;
                            cmd.ExecuteNonQuery();
                        }                        
                        conn.Close();
                    }
                    using (SqlConnection conn = new SqlConnection(path_sql))
                    {
                        conn.Open();
                        var command = "DELETE tbSampleReject WHERE typeSample ='Manual' and SAMNO =" + "'" + ApprovalClickItem.IDNumber + "'";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            cmd.CommandTimeout = 100;
                            cmd.ExecuteNonQuery();
                        }
                        conn.Close();
                    }
                    ProcessSampleHistory(ApprovalClickItem.IDNumber,"Delete");
                    ColorRowListView(Filter_Sample_All());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ProcessButtonEdit_Del", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void ProcessButtonEdit_Edit()
        {
            try
            {
                if (MessageBoxResult.Yes == MessageBox.Show("Bạn muốn sửa dữ liệu?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question))
                {
                    Add_NewData();
                    using (SqlConnection conn = new SqlConnection(path_sql))
                    {
                        conn.Open();
                        
                        //sx
                        if (ckbApprove_SX.IsChecked == true)
                        {
                            str_SX = "N";
                        }
                        else
                        {
                            str_SX = "O";
                        }
                        //cl
                        if (ckbApprove_CL.IsChecked == true)
                        {
                            str_CL = "N";
                        }
                        else
                        {
                            str_CL = "O";
                        }
                        //rd
                        if (ckbApprove_RD.IsChecked == true)
                        {
                            str_RD = "N";
                        }
                        else
                        {
                            str_RD = "O";
                        }
                        //kd
                        if (ckbApprove_KD.IsChecked == true)
                        {
                            str_KD = "N";
                        }
                        else
                        {
                            str_KD = "O";
                        }
                        //str_CUSTMODELCODE = txt_CUSTMODELCODE.Text;
                        //str_CUSTPARTCODE = txt_CUSTPARTCODE.Text;
                        //str_CUST_GB = txt_CUST_GB.Text;
                        //str_VERSION = txt_VERSION.Text;
                        //str_APPLYDT = date;
                        //str_CUSTPART_VERSION = txt_VERSION.Text;
                        //str_COLORIn = txt_PaperIn.Text;
                        //str_COLOROut = txt_PaperOut.Text;
                        //str_MODELSPECOut = txt_FullSize.Text;
                        //str_MODELHEIGHTIn = txt_SizeOut.Text;
                        //str_MODELHEIGHTOut = txt_SizeIn.Text;
                        //str_PAGECNT = txt_FullSize.Text;
                        //str_BCOLORCODEIn = txt_ColorOut.Text;
                        //str_PHCOUNTIn = txt_RatioOut.Text;
                        //str_TOTALPAGEIN = txt_foldIn.Text;
                        //str_VERSIONUP = txt_VERSIONUP.Text;
                        //str_BCOLORCODEOut = txt_ColorIn.Text;
                        //str_PHCOUNTOut = txt_RatioIn.Text;
                        //str_TOTALPAGEOUT = txt_foldOut.Text;
                        //str_NOTE1 = txt_Note1.Text;
                        //str_NOTE2 = txt_Note2.Text;
                        //str_PROCESSWORDK = txt_ProcessWork.Text;
                        //str_QTYREQUEST = txt_Request.Text;
                        //str_DATEAPPROVE = dateInput;
                        //str_DATEINPUT = dateInput;
                        //dateApproval = dateInput;
                        //str_qtyAttach = qtyFileUpload.ToString();
                        //var command = "UPDATE SAMPLEAPPROVE SET model = '" + model + "',modelcode = '" + modelcode + "',customer = '" + customer + "',ver = '" + ver + "'" +
                        //    ",sx = '" + sx + "',cl = '" + cl + "',rd = '" + rd + "',kd = '" + kd + "' WHERE ID = '"+ApproverClickItem.ID+"'";                        
                        var command = "UPDATE tbSampleManual SET custmodelcode = '" + str_CUSTMODELCODE + "',custpartcode = '"
                            + str_CUSTPARTCODE + "',cust_gb = '" + str_CUST_GB + "',version = N'" + str_VERSION + "',versionup=N'" + str_VERSIONUP+ "',papernameOut=N'" + txt_PaperIn.Text+ "',heightOut=N'" + txt_SizeIn.Text+ "',phcountOut=N'" + txt_RatioIn.Text+ "',papernameFullOut=N'" + txt_PaperNameIn.Text+ "',papernameIn=N'" + txt_PaperOut.Text+ "',heightIn=N'" + txt_SizeOut.Text+ "',phcountIn=N'" + txt_RatioOut.Text+ "',papernameFullIn=N'" + txt_PaperNameOut.Text+ "',modelspecOut=N'" + txt_FullSize.Text+ "',pagecnt=N'" + txt_PaperNumber.Text+ "',pagetypem=N'" + cbbTypeCertification.Text + "',pagetyped=N'" + cbbTypeDetail.Text + "',process=N'"+txt_ProcessWork.Text+ "',bcolorcodeOut=N'" + txt_ColorIn.Text+ "',bcolorcodeIn=N'" + txt_ColorOut.Text+ "',totalpageOut=N'" + txt_foldOut.Text+ "',totalpageIn=N'" + txt_foldIn.Text+ "',remark=N'"+txt_Note2.Text+ "',remarkdif=N'"+txt_Note1.Text+"',Updempcode = '" + str_INSEMPCODE+ "',UPddt ='" + dateApproval + "', reject='0'" +
                           " WHERE SAMNO = '" + ApprovalClickItem.IDNumber + "'";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            cmd.CommandTimeout = 100;
                            cmd.ExecuteNonQuery();
                        }
                        ColorRowListView(Filter_Sample_All());
                        conn.Close();
                        MessageBox.Show("Dữ liệu sửa Thành Công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    ProcessSampleHistory(ApprovalClickItem.IDNumber, "Edit");
                }                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ProcessButtonEdit_Edit", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        public async Task ProcessButtonEdit_Save()
        {
            try
            {                
                Add_NewData();
                await Process_AttachFile();
                var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                var jsonDateStart = JsonConvert.SerializeObject(DateTime.Parse(str_DATESTARTAPPROVE), settings);
                var jsonDateFinish = JsonConvert.SerializeObject(DateTime.Parse(str_DATEFINISHAPPROVE), settings);
                string dateInput = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                string dateStart = jsonDateStart.Substring(1, jsonDateStart.Length - 2);
                string dateFinish = jsonDateFinish.Substring(1, jsonDateFinish.Length - 2);

                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    str_MA = "N";
                    //sx
                    if (ckbApprove_SX.IsChecked == true)
                    {
                        str_SX = "N";
                    }
                    else
                    {
                        str_SX = "O";
                    }
                    //cl
                    if (ckbApprove_CL.IsChecked == true)
                    {
                        str_CL = "N";
                    }
                    else
                    {
                        str_CL = "O";
                    }
                    //rd
                    if (ckbApprove_RD.IsChecked == true)
                    {
                        str_RD = "N";
                    }
                    else
                    {
                        str_RD = "O";
                    }
                    //kd
                    if (ckbApprove_KD.IsChecked == true)
                    {
                        str_KD = "N";
                    }
                    else
                    {
                        str_KD = "O";
                    }
                    var command = "";
                    if (str_FilterProduct == "Manual")
                    {
                        command = "INSERT tbSampleManual(cmpcode,bizdiv,samno,modelcode,modelname," +
                        "applydt,version,custpartcode,custpart_version," +
                        "custpartcode_ver,custmodelcode,repmodelcode,useflag,modelgroup,modeldiv,modeltype," +
                        "modelchild,cust_gb,ta,buyer,color,modelspecIn,modellengthIn,modelwidthIn," +
                        "modelheightIn,modelspecOut,modellengtOut,modelwidthOut,modelheightOut,modelunfoldlength," +
                        "modelunfoldwidth,custcode,custshortcode,pagecnt,seq,type,papergubunIn,weightIn," +
                        "widthIn,heightIn,sidegubunIn,frontccIn,backccIn,frontbcolorIn,backbcolorIn,bcolorcodeIn," +
                        "phcountIn,papergubunOut,weightOut,widthOut,heightOut,sidegubunOut,frontccOut,backccOut," +
                        "frontbcolorOut,backbcolorOut,bcolorcodeOut,phcountOut,versionup,papernameIn,papernameFullIn," +
                        "papernameOut,papernameFullOut,dep_mar,dep_pro,dep_qc,dep_rnd,dep_pur,totalpageIn," +
                        "foldingpageIn,tayIn,taypageIn,totalpageOut,foldingpageOut,tayOut,taypageOut,qtyRequest,Remark," +
                        "RemarkDif,process,qtyAttach,reject,depCreate,Insempcode,Insdt,Sadt,Fadt,pagetypem,pagetyped)" +
                       " VALUES (N'" + str_CMPCODE + "',N'" + str_BIZDIV + "',N'" + str_SAMNO + "',N'" + str_MODELCODE + "',N'" + str_MODELNAME + "',N'"
                    + str_APPLYDT + "',N'" + str_VERSION + "',N'" + str_CUSTPARTCODE + "',N'" + str_CUSTPART_VERSION + "',N'"
                    + str_CUSTPARTCODE_VER + "',N'" + str_CUSTMODELCODE + "',N'" + str_REPMODELCODE + "',N'" + str_USEFLAG + "',N'"
                    + str_MODELGROUP + "',N'" + str_MODELDIV + "',N'" + str_MODELTYPE + "',N'" + str_MODELCHILD + "',N'"
                    + str_CUST_GB + "',N'" + str_TA + "',N'" + str_BUYER + "',N'" + str_COLORIn + "',N'" + str_MODELSPECIn + "',N'"
                    + str_MODELLENGTHIn + "',N'" + str_MODELWIDTHIn + "',N'" + str_MODELHEIGHTIn + "',N'" + str_MODELSPECOut + "',N'"
                    + str_MODELLENGTHOut + "',N'" + str_MODELWIDTHOut + "',N'" + str_MODELHEIGHTOut + "',N'" + str_MODELUNFOLDLENGTHIn + "',N'"
                    + str_MODELUNFOLDWIDTHIn + "',N'" + str_CUSTCODE + "',N'" + str_CUSTSHORTCODE + "',N'" + str_PAGECNT + "',N'" + "1" + "',N'"
                    + str_TYPE + "',N'" + str_PAPERGUBUNIn + "',N'" + str_WEIGHTIn + "',N'"
                    + str_WIDTHIn + "',N'" + str_HEIGHTIn + "',N'" + str_SIDEGUBUNIn + "',N'" + str_FRONTCCIn + "',N'" + str_BACKCCIn + "',N'"
                    + str_FRONTBCOLORIn + "',N'" + str_BACKBCOLORIn + "',N'" + str_BCOLORCODEIn + "',N'" + str_PHCOUNTIn + "',N'"
                    + str_PAPERGUBUNOut + "',N'" + str_WEIGHTOut + "',N'" + str_WIDTHOut + "',N'" + str_HEIGHTOut + "',N'"
                    + str_SIDEGUBUNOut + "',N'" + str_FRONTCCOut + "',N'" + str_BACKCCOut + "',N'" + str_FRONTBCOLOROut + "',N'"
                    + str_BACKBCOLOROut + "',N'" + str_BCOLORCODEOut + "',N'" + str_PHCOUNTOut + "',N'" + str_VERSIONUP + "',N'"
                    + str_PAPERNAMEIn + "',N'" + str_PAPERNAMEFULLIn + "',N'" + str_PAPERNAMEOut + "',N'" + str_PAPERNAMEFULLOut + "',N'"
                    + str_MA + "',N'" + str_SX + "',N'" + str_CL + "',N'" + str_RD + "',N'" + str_KD + "',N'" + str_TOTALPAGEIN + "',N'"
                    + str_FOLDINGPAGEIN + "',N'" + str_TAYIN + "',N'" + str_TAYPAGEIN + "',N'" + str_TOTALPAGEOUT + "',N'" + str_FOLDINGPAGEOUT + "',N'"
                    + str_TAYOUT + "',N'" + str_TAYPAGEOUT + "',N'" + str_QTYREQUEST + "',N'" + str_NOTE2 + "',N'" + str_NOTE1 + "',N'" + str_PROCESSWORDK + "',N'"
                    + Window_AttachFile.listAttachFile.Count.ToString() + "',N'0',N'" + str_depCreate + "',N'" + str_INSEMPCODE + "',N'" + dateInput + "',N'" + dateStart + "',N'" + dateFinish + "',N'" + str_PAGETYPEM +"',N'" + str_PAGETYPED + "')";
                    }
                    else if (str_FilterProduct == "UnitBox")
                    {
                        command = "INSERT tbSampleManual(cmpcode,bizdiv,samno,modelcode,modelname," +
                        "applydt,version,custpartcode,custpart_version," +
                        "custpartcode_ver,custmodelcode,repmodelcode,useflag,modelgroup,modeldiv,modeltype," +
                        "modelchild,cust_gb,ta,buyer,color,modelspecIn,modellengthIn,modelwidthIn," +
                        "modelheightIn,modelspecOut,modellengtOut,modelwidthOut,modelheightOut,modelunfoldlength," +
                        "modelunfoldwidth,custcode,custshortcode,pagecnt,seq,type,papergubunIn,weightIn," +
                        "widthIn,heightIn,sidegubunIn,frontccIn,backccIn,frontbcolorIn,backbcolorIn,bcolorcodeIn," +
                        "phcountIn,papergubunOut,weightOut,widthOut,heightOut,sidegubunOut,frontccOut,backccOut," +
                        "frontbcolorOut,backbcolorOut,bcolorcodeOut,phcountOut,versionup,papernameIn,papernameFullIn," +
                        "papernameOut,papernameFullOut,dep_mar,dep_pro,dep_qc,dep_rnd,dep_pur,totalpageIn," +
                        "foldingpageIn,tayIn,taypageIn,totalpageOut,foldingpageOut,tayOut,taypageOut,qtyRequest,Remark," +
                        "RemarkDif,etc1,etc2,Insempcode,Insdt,Sadt,Fadt,pagetypem,pagetyped)" +
                       " VALUES (N'" + str_CMPCODE + "',N'" + str_BIZDIV + "',N'" + str_SAMNO + "',N'" + txt_FilterBox.Text.Trim() + "',N'" + txt_FilterBox.Text.Trim() + "',N'"
                    + str_APPLYDT + "',N'" + str_VERSION + "',N'" + txt_FilterBox.Text + "',N'" + str_CUSTPART_VERSION + "',N'"
                    + txt_FilterBox.Text.Trim() + "',N'" + str_CUSTMODELCODE + "',N'" + str_REPMODELCODE + "',N'" + str_USEFLAG + "',N'"
                    + str_MODELGROUP + "',N'" + str_MODELDIV + "',N'" + str_MODELTYPE + "',N'" + str_MODELCHILD + "',N'"
                    + str_CUST_GB + "',N'" + str_TA + "',N'" + str_BUYER + "',N'" + str_COLORIn + "',N'" + str_MODELSPECIn + "',N'"
                    + str_MODELLENGTHIn + "',N'" + str_MODELWIDTHIn + "',N'" + str_MODELHEIGHTIn + "',N'" + str_MODELSPECOut + "',N'"
                    + str_MODELLENGTHOut + "',N'" + str_MODELWIDTHOut + "',N'" + str_MODELHEIGHTOut + "',N'" + str_MODELUNFOLDLENGTHIn + "',N'"
                    + str_MODELUNFOLDWIDTHIn + "',N'" + str_CUSTCODE + "',N'" + str_CUSTSHORTCODE + "',N'" + str_PAGECNT + "',N'" + "1" + "',N'"
                    + str_TYPE + "',N'" + str_PAPERGUBUNIn + "',N'" + str_WEIGHTIn + "',N'"
                    + str_WIDTHIn + "',N'" + str_HEIGHTIn + "',N'" + str_SIDEGUBUNIn + "',N'" + str_FRONTCCIn + "',N'" + str_BACKCCIn + "',N'"
                    + str_FRONTBCOLORIn + "',N'" + str_BACKBCOLORIn + "',N'" + str_BCOLORCODEIn + "',N'" + str_PHCOUNTIn + "',N'"
                    + str_PAPERGUBUNOut + "',N'" + str_WEIGHTOut + "',N'" + str_WIDTHOut + "',N'" + str_HEIGHTOut + "',N'"
                    + str_SIDEGUBUNOut + "',N'" + str_FRONTCCOut + "',N'" + str_BACKCCOut + "',N'" + str_FRONTBCOLOROut + "',N'"
                    + str_BACKBCOLOROut + "',N'" + str_BCOLORCODEOut + "',N'" + str_PHCOUNTOut + "',N'" + str_VERSIONUP + "',N'"
                    + str_PAPERNAMEIn + "',N'" + str_PAPERNAMEFULLIn + "',N'" + str_PAPERNAMEOut + "',N'" + str_PAPERNAMEFULLOut + "',N'"
                    + str_MA + "',N'" + str_SX + "',N'" + str_CL + "',N'" + str_RD + "',N'" + str_KD + "',N'" + str_TOTALPAGEIN + "',N'"
                    + str_FOLDINGPAGEIN + "',N'" + str_TAYIN + "',N'" + str_TAYPAGEIN + "',N'" + str_TOTALPAGEOUT + "',N'" + str_FOLDINGPAGEOUT + "',N'"
                    + str_TAYOUT + "',N'" + str_TAYPAGEOUT + "',N'" + str_QTYREQUEST + "',N'" + str_NOTE2 + "',N'" + str_NOTE1 + "',N'" + str_PROCESSWORDK + "',N'"
                    + str_FilterProduct + "',N'" + str_INSEMPCODE + "',N'" + dateInput + "',N'" + dateStart + "',N'" + dateFinish + "',N'" + str_PAGETYPEM + "',N'" + str_PAGETYPED + "')";
                    }

                    //if (str_FilterProduct == "Manual")
                    //{
                    //    command = "INSERT tbSampleManual(cmpcode,bizdiv,samno,modelcode,modelname," +
                    //    "applydt,version,custpartcode,custpart_version," +
                    //    "custpartcode_ver,custmodelcode,repmodelcode,useflag,modelgroup,modeldiv,modeltype," +
                    //    "modelchild,cust_gb,ta,buyer,color,modelspecIn,modellengthIn,modelwidthIn," +
                    //    "modelheightIn,modelspecOut,modellengtOut,modelwidthOut,modelheightOut,modelunfoldlength," +
                    //    "modelunfoldwidth,custcode,custshortcode,pagecnt,seq,type,papergubunIn,weightIn," +
                    //    "widthIn,heightIn,sidegubunIn,frontccIn,backccIn,frontbcolorIn,backbcolorIn,bcolorcodeIn," +
                    //    "phcountIn,papergubunOut,weightOut,widthOut,heightOut,sidegubunOut,frontccOut,backccOut," +
                    //    "frontbcolorOut,backbcolorOut,bcolorcodeOut,phcountOut,versionup,papernameIn,papernameFullIn," +
                    //    "papernameOut,papernameFullOut,dep_mar,dep_pro,dep_qc,dep_rnd,dep_pur,totalpageIn," +
                    //    "foldingpageIn,tayIn,taypageIn,totalpageOut,foldingpageOut,tayOut,taypageOut,qtyRequest,Remark," +
                    //    "RemarkDif,process,qtyAttach,reject,depCreate,Insempcode,Insdt,Sadt,Fadt)" +
                    //   " VALUES ('" + str_CMPCODE + "','" + str_BIZDIV + "','" + str_SAMNO + "','" + str_MODELCODE + "','" + str_MODELNAME + "','"
                    //+ str_APPLYDT + "','" + str_VERSION + "','" + str_CUSTPARTCODE + "','" + str_CUSTPART_VERSION + "','"
                    //+ str_CUSTPARTCODE_VER + "','" + str_CUSTMODELCODE + "','" + str_REPMODELCODE + "','" + str_USEFLAG + "','"
                    //+ str_MODELGROUP + "','" + str_MODELDIV + "','" + str_MODELTYPE + "','" + str_MODELCHILD + "','"
                    //+ str_CUST_GB + "','" + str_TA + "','" + str_BUYER + "','" + str_COLORIn + "','" + str_MODELSPECIn + "','"
                    //+ str_MODELLENGTHIn + "','" + str_MODELWIDTHIn + "','" + str_MODELHEIGHTIn + "','" + str_MODELSPECOut + "','"
                    //+ str_MODELLENGTHOut + "','" + str_MODELWIDTHOut + "','" + str_MODELHEIGHTOut + "','" + str_MODELUNFOLDLENGTHIn + "','"
                    //+ str_MODELUNFOLDWIDTHIn + "','" + str_CUSTCODE + "','" + str_CUSTSHORTCODE + "','" + str_PAGECNT + "','" + "1" + "','"
                    //+ str_TYPE + "','" + str_PAPERGUBUNIn + "','" + str_WEIGHTIn + "','"
                    //+ str_WIDTHIn + "','" + str_HEIGHTIn + "','" + str_SIDEGUBUNIn + "','" + str_FRONTCCIn + "','" + str_BACKCCIn + "','"
                    //+ str_FRONTBCOLORIn + "','" + str_BACKBCOLORIn + "','" + str_BCOLORCODEIn + "','" + str_PHCOUNTIn + "','"
                    //+ str_PAPERGUBUNOut + "','" + str_WEIGHTOut + "','" + str_WIDTHOut + "','" + str_HEIGHTOut + "','"
                    //+ str_SIDEGUBUNOut + "','" + str_FRONTCCOut + "','" + str_BACKCCOut + "','" + str_FRONTBCOLOROut + "','"
                    //+ str_BACKBCOLOROut + "','" + str_BCOLORCODEOut + "','" + str_PHCOUNTOut + "',N'" + str_VERSIONUP + "','"
                    //+ str_PAPERNAMEIn + "','" + str_PAPERNAMEFULLIn + "','" + str_PAPERNAMEOut + "','" + str_PAPERNAMEFULLOut + "','"
                    //+ str_MA + "','" + str_SX + "','" + str_CL + "','" + str_RD + "','" + str_KD + "','" + str_TOTALPAGEIN + "','"
                    //+ str_FOLDINGPAGEIN + "','" + str_TAYIN + "','" + str_TAYPAGEIN + "','" + str_TOTALPAGEOUT + "','" + str_FOLDINGPAGEOUT + "','"
                    //+ str_TAYOUT + "','" + str_TAYPAGEOUT + "','" + str_QTYREQUEST + "',N'" + str_NOTE2 + "',N'" + str_NOTE1 + "',N'" + str_PROCESSWORDK + "','"
                    //+ Window_AttachFile.listAttachFile.Count.ToString()+"','0','"+str_depCreate+"','" + str_INSEMPCODE + "','" + dateInput + "','" + dateStart + "','" + dateFinish + "')";
                    //}
                    //else if (str_FilterProduct == "UnitBox")
                    //{
                    //    command = "INSERT tbSampleManual(cmpcode,bizdiv,samno,modelcode,modelname," +
                    //    "applydt,version,custpartcode,custpart_version," +
                    //    "custpartcode_ver,custmodelcode,repmodelcode,useflag,modelgroup,modeldiv,modeltype," +
                    //    "modelchild,cust_gb,ta,buyer,color,modelspecIn,modellengthIn,modelwidthIn," +
                    //    "modelheightIn,modelspecOut,modellengtOut,modelwidthOut,modelheightOut,modelunfoldlength," +
                    //    "modelunfoldwidth,custcode,custshortcode,pagecnt,seq,type,papergubunIn,weightIn," +
                    //    "widthIn,heightIn,sidegubunIn,frontccIn,backccIn,frontbcolorIn,backbcolorIn,bcolorcodeIn," +
                    //    "phcountIn,papergubunOut,weightOut,widthOut,heightOut,sidegubunOut,frontccOut,backccOut," +
                    //    "frontbcolorOut,backbcolorOut,bcolorcodeOut,phcountOut,versionup,papernameIn,papernameFullIn," +
                    //    "papernameOut,papernameFullOut,dep_mar,dep_pro,dep_qc,dep_rnd,dep_pur,totalpageIn," +
                    //    "foldingpageIn,tayIn,taypageIn,totalpageOut,foldingpageOut,tayOut,taypageOut,qtyRequest,Remark," +
                    //    "RemarkDif,etc1,etc2,Insempcode,Insdt,Sadt,Fadt)" +
                    //   " VALUES ('" + str_CMPCODE + "','" + str_BIZDIV + "','" + str_SAMNO + "','" + txt_FilterBox.Text.Trim() + "','" + txt_FilterBox.Text.Trim() + "','"
                    //+ str_APPLYDT + "','" + str_VERSION + "','" + txt_FilterBox.Text + "','" + str_CUSTPART_VERSION + "','"
                    //+ txt_FilterBox.Text.Trim() + "','" + str_CUSTMODELCODE + "','" + str_REPMODELCODE + "','" + str_USEFLAG + "','"
                    //+ str_MODELGROUP + "','" + str_MODELDIV + "','" + str_MODELTYPE + "','" + str_MODELCHILD + "','"
                    //+ str_CUST_GB + "','" + str_TA + "','" + str_BUYER + "','" + str_COLORIn + "','" + str_MODELSPECIn + "','"
                    //+ str_MODELLENGTHIn + "','" + str_MODELWIDTHIn + "','" + str_MODELHEIGHTIn + "','" + str_MODELSPECOut + "','"
                    //+ str_MODELLENGTHOut + "','" + str_MODELWIDTHOut + "','" + str_MODELHEIGHTOut + "','" + str_MODELUNFOLDLENGTHIn + "','"
                    //+ str_MODELUNFOLDWIDTHIn + "','" + str_CUSTCODE + "','" + str_CUSTSHORTCODE + "','" + str_PAGECNT + "','" + "1" + "','"
                    //+ str_TYPE + "','" + str_PAPERGUBUNIn + "','" + str_WEIGHTIn + "','"
                    //+ str_WIDTHIn + "','" + str_HEIGHTIn + "','" + str_SIDEGUBUNIn + "','" + str_FRONTCCIn + "','" + str_BACKCCIn + "','"
                    //+ str_FRONTBCOLORIn + "','" + str_BACKBCOLORIn + "','" + str_BCOLORCODEIn + "','" + str_PHCOUNTIn + "','"
                    //+ str_PAPERGUBUNOut + "','" + str_WEIGHTOut + "','" + str_WIDTHOut + "','" + str_HEIGHTOut + "','"
                    //+ str_SIDEGUBUNOut + "','" + str_FRONTCCOut + "','" + str_BACKCCOut + "','" + str_FRONTBCOLOROut + "','"
                    //+ str_BACKBCOLOROut + "','" + str_BCOLORCODEOut + "','" + str_PHCOUNTOut + "',N'" + str_VERSIONUP + "','"
                    //+ str_PAPERNAMEIn + "','" + str_PAPERNAMEFULLIn + "','" + str_PAPERNAMEOut + "','" + str_PAPERNAMEFULLOut + "','"
                    //+ str_MA + "','" + str_SX + "','" + str_CL + "','" + str_RD + "','" + str_KD + "','" + str_TOTALPAGEIN + "','"
                    //+ str_FOLDINGPAGEIN + "','" + str_TAYIN + "','" + str_TAYPAGEIN + "','" + str_TOTALPAGEOUT + "','" + str_FOLDINGPAGEOUT + "','"
                    //+ str_TAYOUT + "','" + str_TAYPAGEOUT + "','" + str_QTYREQUEST + "',N'" + str_NOTE2 + "',N'" + str_NOTE1 + "',N'" + str_PROCESSWORDK + "','"
                    //+ str_FilterProduct + "','" + str_INSEMPCODE + "','" + dateInput + "','" + dateStart + "','" + dateFinish + "')";
                    //}

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        cmd.ExecuteNonQuery();
                    }
                    ColorRowListView(Filter_Sample_All());
                    conn.Close();
                    Window_AttachFile.listAttachFile.Clear();
                }
                ProcessSampleHistory(str_SAMNO,"Save");                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ProcessButtonEdit_Save", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        public void ProcessButtonEdit_Printer()
        {
            MainWindow.pl_Print = "Manual";
            MainWindow.print.Show();
            MainWindow.checkPrint = true;
        }

        bool checkRun = false;
        public void ProcessButtonEdit_Run()
        {
            try
            {
                if (MessageBox.Show("Bạn có muốn xác nhận Sample này không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                {
                    if (checkRun == true)
                    {
                        using (SqlConnection conn = new SqlConnection(path_sql))
                        {
                            conn.Open();

                            var command = "Update tbSampleManual SET printed='Y' where samno='" + SamplePaper.PaperOut.IDNumber + "' ";
                            using (SqlCommand cmd = new SqlCommand(command, conn))
                            {
                                cmd.CommandTimeout = 100;
                                cmd.ExecuteNonQuery();
                            }
                            ColorRowListView(Filter_Sample_All());
                            ProcessSampleHistory(SamplePaper.PaperOut.IDNumber, "Printed");
                            conn.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Bạn chưa được phê duyệt chức năng này!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    if (checkRun == true)
                    {
                        using (SqlConnection conn = new SqlConnection(path_sql))
                        {
                            conn.Open();

                            var command = "Update tbSampleManual SET printed='' where samno='" + SamplePaper.PaperOut.IDNumber + "' ";
                            using (SqlCommand cmd = new SqlCommand(command, conn))
                            {
                                cmd.CommandTimeout = 100;
                                cmd.ExecuteNonQuery();
                            }
                            ColorRowListView(Filter_Sample_All());
                            ProcessSampleHistory(SamplePaper.PaperOut.IDNumber, "Printed");
                            conn.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Bạn chưa được phê duyệt chức năng này!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/ProcessButtonEdit_Run", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void ProcessUploadFile()
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += Worker_DoWork;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
            worker.RunWorkerAsync(10000);
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //MessageBox.Show("Tạo mới thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        public async Task Process_AttachFile()
        {
            if (accessProcess == "Creat" && checkDowload == false)
            {
                if (Window_AttachFile.listAttachFile.Count > 0)
                {                   
                    foreach (var item in Window_AttachFile.listAttachFile)
                    {
                        if (item.Check == "New")
                        {
                            await UploadAttachFile(item.FileName, item.Path, str_SAMNO, item.Stt);
                        }
                    }                   
                }
            }
            if (accessProcess == "Approve" || checkDowload == true)
            {

                using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                    if (dialog.SelectedPath != "")
                    {
                        await DowLoadAttachFile(dialog.SelectedPath);                        
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng chọn vị trí lưu tập tin", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                }
            }
        }


        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {          

            Application.Current.Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            new Action(() =>
            {
                
            }));


        }

        public void ColorRowListView(List<Helper_TaixinDB_Model> listAllData_Input)
        {
            try
            {
                if (lvApproveSample != null)
                {
                    lvApproveSample.Items.Clear();
                    int index = 0;
                    if (listAllData_Input != null)
                    {
                        
                        foreach (var item in listAllData_Input)
                        {
                            item.ID = index;
                            if (item.reject == "1" && str_depCreate=="MARKETING")
                            {
                                item.MA = "Purple";
                            }
                            if (item.reject == "1" && str_depCreate == "RND")
                            {
                                item.RD = "Purple";
                            }
                            index++;
                        }

                        foreach (var item in listAllData_Input)
                        {
                            lvApproveSample.Items.Add(item);
                            //frameLoadingData.Visibility = Visibility.Visible;
                            //frameLoadingData.Navigate(loadingData);
                        }
                        //frameLoadingData.Visibility = Visibility.Hidden;
                        if (checkView == true)
                            lvApproveSample.SelectedIndex = 0;
                    }                    
                    
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ColorRowListView", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void ManagerApproval()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "";
                    var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                    var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                    string dateInput = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                    dateApproval = dateInput;
                    checkView = true;
                    MainWindow.timeApproveManual = dateInput;
                    switch (department)
                    {
                        case "SX":
                            {
                                command = "UPDATE tbSampleManual set dep_PRO = 'Y',INSDT = '" + dateInput + "' where SAMNO = " + " '" + ApprovalClickItem.IDNumber + "'";
                                break;
                            }
                        case "CL":
                            {
                                command = "UPDATE tbSampleManual set dep_QC = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.IDNumber + "'";
                                break;
                            }
                        case "RD":
                            {
                                command = "UPDATE tbSampleManual set dep_RND = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.IDNumber + "'";
                                break;
                            }
                        case "KD":
                            {
                                command = "UPDATE tbSampleManual set dep_PUR = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.IDNumber + "'";
                                break;
                            }
                        case "MA":
                            {
                                command = "UPDATE tbSampleManual set dep_MAR = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.IDNumber + "'";
                                break;
                            }
                    }
                    ProcessSampleHistory(ApprovalClickItem.IDNumber,"Arv-"+department);
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 100;
                        cmd.ExecuteNonQuery();
                    }
                    ColorRowListView(Filter_Sample_All());
                    MessageBox.Show("Mẫu được phê duyệt Thành Công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    conn.Close();

                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ManagerApproval", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }        

        public void AccessManage()
        {
            try
            {
                Helper_AccessManger access = new Helper_AccessManger();
                stackApprove.Visibility = Visibility.Hidden;
                access.CreatApprove = "Y";
                access.ApproveApprove = "Y";
                if (access.CreatApprove == "Y" && access.ApproveApprove == "Y")
                {
                    ckbTest.Visibility = Visibility.Visible;
                    CreatData();
                }
                else
                {
                    if (access.CreatApprove == "Y")
                    {
                        CreatData();
                    }
                    if (access.ApproveApprove == "Y")
                    {
                        ApproveData();
                    }
                }
            }
            catch (Exception ex)     
            {
                MessageBox.Show(ex.Message, "Manual/AccessManage", MessageBoxButton.OK, MessageBoxImage.Error);
            }       
            
        }

        public void CreatData()
        {
            CreatAllButtonEdit();
            ckbTest.Content = "Chuyển đổi";
            accessProcess = "Creat";
            stackManul.Visibility = Visibility.Visible;
            lvButtonTop.Visibility = Visibility.Visible;
            stackApprove.Visibility = Visibility.Visible;           
            grid_ButtonEditor.Visibility = Visibility.Visible;
            btnApprove_Ma.Visibility = Visibility.Hidden;
            btnApprove_SX.Visibility = Visibility.Hidden;
            btnApprove_CL.Visibility = Visibility.Hidden;
            btnApprove_RD.Visibility = Visibility.Hidden;
            btnApprove_KD.Visibility = Visibility.Hidden;
            ckbApprove_Ma.Visibility = Visibility.Visible;
            ckbApprove_SX.Visibility = Visibility.Visible;
            ckbApprove_CL.Visibility = Visibility.Visible;
            ckbApprove_KD.Visibility = Visibility.Visible;
            ckbApprove_RD.Visibility = Visibility.Visible;
            txt_FilterBox.IsEnabled = true;
            rb_Manual.Visibility = Visibility.Visible;
            rb_UnitBox.Visibility = Visibility.Visible;
            ckb_DowloadFile.Visibility = Visibility.Visible;

        }

        public void ApproveData()
        {
            lvButtonTop.Items.Clear(); 
            if(listButtonTop.Count<6)
            listButtonTop.Add(new Helper_DataButton
            {
                ID = 6,
                ContentButton = "Run",
                ImageSource = "Image/Edit/check.png",
                BackGroundColor = PinValue.OFF
            });
            foreach (var button in listButtonTop)
            {
                if (button.ID > 4)
                    lvButtonTop.Items.Add(button);

            }
            ckbTest.Content = "Chuyển đổi";
            accessProcess = "Approve";
            stackApprove.Visibility = Visibility.Visible;
            //grid_ButtonEditor.Visibility = Visibility.Hidden;           
            btnApprove_Ma.Visibility = Visibility.Visible;
            btnApprove_SX.Visibility = Visibility.Visible;
            btnApprove_CL.Visibility = Visibility.Visible;
            btnApprove_RD.Visibility = Visibility.Visible;
            btnApprove_KD.Visibility = Visibility.Visible;
            ckbApprove_Ma.Visibility = Visibility.Hidden;
            ckbApprove_SX.Visibility = Visibility.Hidden;
            ckbApprove_CL.Visibility = Visibility.Hidden;
            ckbApprove_KD.Visibility = Visibility.Hidden;
            ckbApprove_RD.Visibility = Visibility.Hidden;
            stackManul.Visibility = Visibility.Hidden;
            stackBox.Visibility = Visibility.Visible;
            rb_Manual.Visibility = Visibility.Hidden;
            rb_UnitBox.Visibility = Visibility.Hidden;
            txt_FilterBox.IsEnabled = false;
            if (str_FilterProduct == "Manual")
            {
                stackBox.Visibility = Visibility.Hidden;
            }
            else
            {
                stackBox.Visibility = Visibility.Visible;
            }
            ckb_DowloadFile.Visibility = Visibility.Hidden;

        }

        public void DataItemClick(Helper_TaixinDB_Model clickItem)
        {
            try
            {                
                SamplePaper.PaperOut.CMPCODE = clickItem.CMPCODE;
                SamplePaper.PaperOut.BIZDIV = clickItem.BIZDIV;
                SamplePaper.PaperOut.IDNumber = clickItem.IDNumber;
                MainWindow.at_Samno = clickItem.IDNumber;
                SamplePaper.PaperOut.MODELCODE = clickItem.MODELCODE;                
                SamplePaper.PaperOut.MODELNAME = clickItem.MODELNAME;
                SamplePaper.PaperOut.APPLYDT = clickItem.APPLYDT;
                SamplePaper.PaperOut.VERSION = clickItem.VERSION;
                SamplePaper.PaperOut.CUSTPARTCODE = clickItem.CUSTPARTCODE;
                MainWindow.at_ModelCode = clickItem.CUSTPARTCODE;
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
                SamplePaper.PaperOut.MODELWIDTHOut = clickItem.MODELLENGTHOut;
                SamplePaper.PaperOut.MODELHEIGHTOut = clickItem.MODELHEIGHTOut;
                SamplePaper.PaperOut.MODELUNFOLDLENGTH = clickItem.MODELUNFOLDLENGTH;
                SamplePaper.PaperOut.MODELUNFOLDWIDTH = clickItem.MODELUNFOLDWIDTH;
                SamplePaper.PaperOut.CUSTCODE = clickItem.CUSTCODE;
                SamplePaper.PaperOut.CUSTSHORTCODE = clickItem.CUSTSHORTCODE;
                SamplePaper.PaperOut.PAGECNT = clickItem.PAGECNT;
                SamplePaper.PaperOut.PAGETYPEM = clickItem.PAGETYPEM;
                SamplePaper.PaperOut.PAGETYPED = clickItem.PAGETYPED;   
                SamplePaper.PaperOut.SEQ = clickItem.SEQ;
                SamplePaper.PaperOut.TYPE = clickItem.TYPE;
                SamplePaper.PaperOut.PAPERGUBUNOut = clickItem.PAPERGUBUNOut;
                SamplePaper.PaperOut.WEIGHTOut = clickItem.WEIGHTOut;
                SamplePaper.PaperOut.WIDTHOut = clickItem.WIDTHOut;
                SamplePaper.PaperOut.HEIGHTOut = clickItem.HEIGHTOut;
                SamplePaper.PaperOut.SIDEGUBUNOut = clickItem.SIDEGUBUNOut;
                SamplePaper.PaperOut.FRONTCCOut = clickItem.FRONTCCOut;
                SamplePaper.PaperOut.BACKCCOut = clickItem.BACKCCOut;
                SamplePaper.PaperOut.FRONTBCOLOROut = clickItem.FRONTBCOLOROut;
                SamplePaper.PaperOut.BACKBCOLOROut = clickItem.BACKBCOLOROut;
                SamplePaper.PaperOut.BCOLORCODEOut = clickItem.BCOLORCODEOut;
                SamplePaper.PaperOut.PHCOUNTOut = clickItem.PHCOUNTOut;
                SamplePaper.PaperOut.TOTALPAGEOUT = clickItem.TOTALPAGEOUT;
                SamplePaper.PaperOut.FOLDINGPAGEOUT = clickItem.FOLDINGPAGEOUT;
                SamplePaper.PaperOut.TAYOUT = clickItem.TAYOUT;
                SamplePaper.PaperOut.TAYPAGEOUT = clickItem.TAYPAGEOUT;
                SamplePaper.PaperOut.VERSIONUP = clickItem.VERSIONUP;
                SamplePaper.PaperOut.PAPERNAMEOut = clickItem.PAPERNAMEOut;
                SamplePaper.PaperOut.INSEMPCODE = clickItem.INSEMPCODE;
                SamplePaper.PaperOut.PAPERNAME_FullOut = clickItem.PAPERNAME_FullOut;
                SamplePaper.PaperOut.MA = clickItem.MA;
                SamplePaper.PaperOut.SX = clickItem.SX;
                SamplePaper.PaperOut.CL = clickItem.CL;
                SamplePaper.PaperOut.RD = clickItem.RD;
                SamplePaper.PaperOut.KD = clickItem.KD;
                SamplePaper.PaperOut.NOTE1 = clickItem.NOTE1;
                SamplePaper.PaperOut.NOTE2 = clickItem.NOTE2;
                SamplePaper.PaperOut.ETC1 = clickItem.ETC1;
                SamplePaper.PaperOut.ETC3 = clickItem.ETC3 + " File";
                SamplePaper.PaperOut.QTYREQUEST = clickItem.QTYREQUEST;
                SamplePaper.PaperOut.DATESTARTAPPROVE = clickItem.DATESTARTAPPROVE;
                SamplePaper.PaperOut.DATEFINISHAPPROVE = clickItem.DATEFINISHAPPROVE;
                SamplePaper.PaperOut.printed = clickItem.printed;

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
                SamplePaper.PaperIn.PAGETYPEM = clickItem.PAGETYPEM;
                SamplePaper.PaperIn.PAGETYPED = clickItem.PAGETYPED;
                SamplePaper.PaperIn.SEQ = clickItem.SEQ;
                SamplePaper.PaperIn.TYPE = clickItem.TYPE;
                SamplePaper.PaperIn.PAPERGUBUNIn = clickItem.PAPERGUBUNIn;
                SamplePaper.PaperIn.WEIGHTIn = clickItem.WEIGHTIn;
                SamplePaper.PaperIn.WIDTHIn = clickItem.WIDTHIn;
                SamplePaper.PaperIn.HEIGHTIn = clickItem.HEIGHTIn;
                SamplePaper.PaperIn.SIDEGUBUNIn = clickItem.SIDEGUBUNIn;
                SamplePaper.PaperIn.FRONTCCIn = clickItem.FRONTCCIn;
                SamplePaper.PaperIn.BACKCCIn = clickItem.BACKCCIn;
                SamplePaper.PaperIn.FRONTBCOLORIn = clickItem.FRONTBCOLORIn;
                SamplePaper.PaperIn.BACKBCOLORIn = clickItem.BACKBCOLORIn;
                SamplePaper.PaperIn.BCOLORCODEIn = clickItem.BCOLORCODEIn;
                SamplePaper.PaperIn.PHCOUNTIn = clickItem.PHCOUNTIn;
                SamplePaper.PaperIn.TOTALPAGEIN = clickItem.TOTALPAGEIN;
                SamplePaper.PaperIn.FOLDINGPAGEIN = clickItem.FOLDINGPAGEIN;
                SamplePaper.PaperIn.TAYIN = clickItem.TAYIN;
                SamplePaper.PaperIn.TAYPAGEIN = clickItem.TAYPAGEIN;
                SamplePaper.PaperIn.VERSIONUP = clickItem.VERSIONUP;
                SamplePaper.PaperIn.PAPERNAMEIn = clickItem.PAPERNAMEIn;
                SamplePaper.PaperIn.INSEMPCODE = clickItem.INSEMPCODE;
                SamplePaper.PaperIn.PAPERNAME_FullIn = clickItem.PAPERNAME_FullIn;                
                Page_RejectSample.samno = clickItem.IDNumber;
                Page_RejectSample.modelcode = clickItem.CUSTPARTCODE;
                Page_RejectSample.typeSample = "Manual";               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/DataItemClick", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        public void FilterPaperInBox(string strModelCode)
        {
            try
            {
                DataBaseHelper db = new DataBaseHelper();
                ListPaper = db.Read_TaxinDb_ModelCode(path_sql, "TMSTMODEL", "ModelCode", strModelCode);
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
                SamplePaper.PaperIn.PAGETYPEM = "";
                SamplePaper.PaperIn.PAGETYPED = "";
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
                SamplePaper.PaperOut.PAGETYPEM = "";
                SamplePaper.PaperOut.PAGETYPED = "";
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

                    if (clickItem.SEQ == "1" || clickItem.SEQ == "")
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
                        SamplePaper.PaperOut.PAGETYPEM = clickItem.PAGETYPEM;
                        SamplePaper.PaperOut.PAGETYPED = clickItem.PAGETYPED;
                        SamplePaper.PaperOut.SEQ = clickItem.SEQ;
                        SamplePaper.PaperOut.TYPE = clickItem.TYPE;
                        SamplePaper.PaperOut.PAPERGUBUNOut = clickItem.PAPERGUBUNOut;
                        SamplePaper.PaperOut.WEIGHTOut = clickItem.WEIGHTOut;
                        SamplePaper.PaperOut.WIDTHOut = clickItem.WIDTHOut;
                        SamplePaper.PaperOut.HEIGHTOut = clickItem.HEIGHTOut;
                        SamplePaper.PaperOut.SIDEGUBUNOut = clickItem.SIDEGUBUNOut;
                        SamplePaper.PaperOut.FRONTCCOut = clickItem.FRONTCCOut;
                        SamplePaper.PaperOut.BACKCCOut = clickItem.BACKCCOut;
                        SamplePaper.PaperOut.FRONTBCOLOROut = clickItem.FRONTBCOLOROut;
                        SamplePaper.PaperOut.BACKBCOLOROut = clickItem.BACKBCOLOROut;
                        SamplePaper.PaperOut.BCOLORCODEOut = clickItem.BCOLORCODEOut;
                        if (clickItem.PHCOUNTOut != "")
                            SamplePaper.PaperOut.PHCOUNTOut = "1*" + clickItem.PHCOUNTOut;
                        string process = ReadProcessWork(clickItem.MODELCODE).ToUpper();
                        if (process != "")
                            SamplePaper.PaperOut.ETC1 = process.Substring(0, process.Length - 2);
                        if (clickItem.FOLDINGPAGEOUT != "")
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
                        SamplePaper.PaperIn.PAGETYPEM = clickItem.PAGETYPEM;
                        SamplePaper.PaperIn.PAGETYPED = clickItem.PAGETYPED;
                        SamplePaper.PaperIn.SEQ = clickItem.SEQ;
                        SamplePaper.PaperIn.TYPE = clickItem.TYPE;
                        SamplePaper.PaperIn.PAPERGUBUNIn = clickItem.PAPERGUBUNIn;
                        SamplePaper.PaperIn.WEIGHTIn = clickItem.WEIGHTIn;
                        SamplePaper.PaperIn.WIDTHIn = clickItem.WIDTHIn;
                        SamplePaper.PaperIn.HEIGHTIn = clickItem.HEIGHTIn;
                        SamplePaper.PaperIn.SIDEGUBUNIn = clickItem.SIDEGUBUNIn;
                        SamplePaper.PaperIn.FRONTCCIn = clickItem.FRONTCCIn;
                        SamplePaper.PaperIn.BACKCCIn = clickItem.BACKCCIn;
                        SamplePaper.PaperIn.FRONTBCOLORIn = clickItem.FRONTBCOLORIn;
                        SamplePaper.PaperIn.BACKBCOLORIn = clickItem.BACKBCOLORIn;
                        SamplePaper.PaperIn.BCOLORCODEIn = clickItem.BCOLORCODEIn;
                        SamplePaper.PaperIn.PHCOUNTIn = "1*" + clickItem.PHCOUNTIn;
                        if (clickItem.FOLDINGPAGEOUT != "")
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
                        if (clickItem.PAPERNAMEIn != "")
                            SamplePaper.PaperIn.PAPERNAMEIn = clickItem.PAPERNAMEIn + " " + clickItem.WEIGHTIn + "g";
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
                    SamplePaper.PaperIn.PAGETYPEM = "";
                    SamplePaper.PaperIn.PAGETYPED = "";
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
                MessageBox.Show(ex.Message, "Manual/FilterPaperInBox", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public List<Helper_TaixinDB_Model> ProcessBoxUnit(string parentCode)
        {
            try
            {
                cbb_BoxUnit.Items.Clear();
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "SELECT * FROM tmstBOM a LEFT OUTER JOIN tmstmodel b on a.ChildCode = b.modelcode LEFT OUTER JOIN tmstmodelpaper c on a.ChildCode = c.modelcode LEFT OUTER JOIN tmstpaper d on c.papergubun = d.papercode where a.ParentCode = '" + parentCode + "' order by SeqNo asc";
                    int index = 0;
                    List<Helper_TaixinDB_Model> listModel = new List<Helper_TaixinDB_Model>();
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                Helper_TaixinDB_Model model = new Helper_TaixinDB_Model();
                                if (dr[0] != null)
                                {
                                    index++;
                                    model.MODELCODE = dr[29].ToString();
                                    model.MODELNAME = dr[30].ToString();
                                    listModel.Add(model);
                                }
                            }
                        }
                    }
                    conn.Close();
                    if (index == 0)
                    {
                        MessageBox.Show("Mã vừa nhập không tồn tại. \r\nVui lòng kiểm tra lại", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    foreach (var item in listModel)
                    {
                        cbb_BoxUnit.Items.Add(item);
                    }
                    cbb_BoxUnit.SelectedIndex = 0;
                    return listModel;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ProcessBoxUnit", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public List<Helper_TaixinDB_Model> ListPaper = new List<Helper_TaixinDB_Model>();

        public async Task UploadAttachFile(string namefile, string path_File, string samno, string qty)
        {
            try
            {
                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        MainWindow.txbloading = "Waitting save data";
                        Page_LoadingData page_Loading = new Page_LoadingData();
                        frameLoading.Navigate(page_Loading);
                        frameLoading.Visibility = Visibility.Visible;
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);

                });
                await Task.Run(() =>
                {
                    using (SqlConnection conn = new SqlConnection(path_sql_attach))                 
                    {
                        conn.Open();
                        string typeSample = "Manual";
                        string modelcode = str_CUSTPARTCODE;
                        byte[] buffer = File.ReadAllBytes(path_File);
                        string base64Encoded = Convert.ToBase64String(buffer);
                        var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                        var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                        string seq = db.Rejetc_MaxSeq(path_sql_attach, "tbSampleAttach", samno, typeSample);
                        string imsempcode = MainWindow.UserLogin;
                        string insdt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                        string updempcode = MainWindow.UserLogin;
                        string upddt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                        //string query = ("INSERT tbSampleAttach(cmpcode,bizdiv,samno,seq,typeSample,modelcode,filename,filedata,qty,imsempcode,insdt,updempcode,upddt) VALUES('02','300','" + samno + "','" + seq + "','" + typeSample + "','" + modelcode + "',N'" + namefile + "','" + base64Encoded + "','" + qty + "','" + imsempcode + "','" + insdt + "','" + updempcode + "','" + upddt + "')");
                        string query = ("INSERT tbSampleAttach(cmpcode,bizdiv,samno,seq,typeSample,modelcode,filename,filedata,qty,imsempcode,insdt,updempcode,upddt) VALUES('02',N'300',N'" + samno + "',N'" + seq + "',N'" + typeSample + "',N'" + modelcode + "',N'" + namefile + "',N'" + base64Encoded + "',N'" + qty + "',N'" + imsempcode + "',N'" + insdt + "',N'" + updempcode + "',N'" + upddt + "')");
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            int count = (int)cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }
                });
                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        frameLoading.Visibility = Visibility.Hidden;
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);
                });
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/UploadAttachFile", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        public string FileType(string fileName)
        {
            string text = fileName;
            int lengText = text.Length;
            int vitri = text.IndexOf(".");
            string fileType = text.Substring(vitri, lengText - vitri);
            return fileType;
        }

        public async Task DowLoadAttachFile(string pathFolder)
        {
            try
            {
                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        MainWindow.txbloading = "Waitting dowload";
                        Page_LoadingData page_Loading = new Page_LoadingData();
                        frameLoading.Navigate(page_Loading);
                        frameLoading.Visibility = Visibility.Visible;
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);

                });
                await Task.Run(() =>
                {
                    string bufferExe;
                    using (SqlConnection conn = new SqlConnection(path_sql_attach))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand("select * from tbSampleAttach where typeSample = 'Manual' and samno ='" + ApprovalClickItem.IDNumber + "'", conn))
                        {
                            using (IDataReader dr = cmd.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    if (dr[0] != null)
                                    {
                                        string fileType = FileType(dr[6].ToString());
                                        string name = dr[6].ToString();
                                        bufferExe = dr[7].ToString();
                                        byte[] buffer = Convert.FromBase64String(bufferExe);
                                        string fileSave = pathFolder + "\\" + name + fileType;
                                        File.WriteAllBytes(fileSave, buffer);
                                    }
                                }
                            }
                        }
                        conn.Close();
                    }
                });
                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        frameLoading.Visibility = Visibility.Hidden;
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/DowLoadAttachFile", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public string Read_QtyAttachFile()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql_attach))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("select * from tbSampleAttach where samno ='" + Page_RejectSample.samno + "'", conn))
                    {
                        using (IDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                if (dr[0].ToString() != "")
                                {
                                    qtyAttachFile = dr[8].ToString();
                                }
                                else
                                {
                                    qtyAttachFile = "0";
                                }
                            }
                        }
                    }
                    conn.Close();
                    return qtyAttachFile;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/DowLoadAttachFile", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public async void SelectFileAttach()
        {
            if (accessProcess == "Creat" && checkDowload == false)
            {
                try
                {
                    ofd = new OpenFileDialog();
                    ofd.Filter = "All files(*.*)| *.*| Exe Files(*.exe) | *.exe*| Text File(*.txt) |*.txt";
                    ofd.FilterIndex = 0;
                    ofd.Multiselect = true;
                    qtyFileUpload = 0;
                    if (ofd.ShowDialog() == true)
                    {
                        foreach (var item in ofd.FileNames)
                        {
                            var onlyFileName = System.IO.Path.GetFileName(item);
                            long length = new System.IO.FileInfo(item).Length / 1000;
                            if (length > 80000)
                            {
                                MessageBox.Show("File có dung lượng quá lớn.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                            qtyFileUpload = ofd.FileNames.Count();
                        }
                        //txb_NameFileUpload.Text = qtyFileUpload + " File";
                        //MessageBox.Show("Upload file thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng chọn File cần Upload", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Manual/SelectFileAttach", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else if (accessProcess == "Approve" || checkDowload == true)
            {
                try
                {
                    using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                    {
                        System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                        if (dialog.SelectedPath != "")
                        {
                           await DowLoadAttachFile(dialog.SelectedPath);                          
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng chọn vị trí lưu tập tin", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Manual/SelectFileAttach", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }


        }

        public string ReadProcessWork(string modelcode)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "SELECT b.workcodename_loc FROM tmstmodelwork a left outer JOIN tmstprocwork b on a.workcode = b.workcode WHERE a.modelcode = '" + modelcode + "' ORDER by seq ASC";
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
                MessageBox.Show(ex.Message, "Manual/ReadProcessWork", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private void lvApproveSample_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var click = sender as ListView;
                var clickItem = click.SelectedItem as Helper_TaixinDB_Model;


                if (clickItem != null)
                {
                    qtyAttachFile = "0";
                    //FilterPaperInBox(clickItem.MODELCODE);
                    DataItemClick(clickItem);
                    dpkStartApprove.SelectedDate = DateTime.Parse(clickItem.DATESTARTAPPROVE.ToString());
                    dpkFinishApprove.SelectedDate = DateTime.Parse(clickItem.DATEFINISHAPPROVE.ToString());
                    cbbTypeCertification.Text = clickItem.PAGETYPEM;
                    cbbTypeDetail.Text = clickItem.PAGETYPED;
                    ApprovalClickItem = clickItem;
                    //ma
                    if (ApprovalClickItem.MA == "LightGray")
                    {
                        ckbApprove_Ma.IsChecked = false;
                    }
                    else
                    {
                        ckbApprove_Ma.IsChecked = true;
                    }
                    //sx
                    if (ApprovalClickItem.SX == "LightGray")
                    {
                        ckbApprove_SX.IsChecked = false;
                    }
                    else
                    {
                        ckbApprove_SX.IsChecked = true;
                    }
                    //cl
                    if (ApprovalClickItem.CL == "LightGray")
                    {
                        ckbApprove_CL.IsChecked = false;
                    }
                    else
                    {
                        ckbApprove_CL.IsChecked = true;
                    }
                    //rd
                    if (ApprovalClickItem.RD == "LightGray")
                    {
                        ckbApprove_RD.IsChecked = false;
                    }
                    else
                    {
                        ckbApprove_RD.IsChecked = true;
                    }
                    //kd
                    if (ApprovalClickItem.KD == "LightGray")
                    {
                        ckbApprove_KD.IsChecked = false;
                    }
                    else
                    {
                        ckbApprove_KD.IsChecked = true;
                    }
                    if (str_FilterProduct == "UnitBox")
                    {
                        ProcessBoxUnit(clickItem.MODELCODE);
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Manual/lvApproveSample_SelectionChanged", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        private void btnApprove_SX_Click(object sender, RoutedEventArgs e)
        {
            if(str_depCreate == "SX" && ApprovalClickItem.SX == "Red" && (ApprovalClickItem.MA == "DodgerBlue" || ApprovalClickItem.RD == "DodgerBlue"))
            {
                department = "SX";
                ManagerApproval();
            }
            else
            {
                MessageBox.Show("Bạn không có quyền phê duyệt. \nHoặc bộ phận tạo chưa phê duyệt", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            }
        }

        private void btnApprove_CL_Click(object sender, RoutedEventArgs e)
        {
            if (str_depCreate == "QC" && ApprovalClickItem.CL == "Red" && (ApprovalClickItem.MA == "DodgerBlue"||ApprovalClickItem.RD== "DodgerBlue"))
            {
                department = "CL";
                ManagerApproval();
            }
            else
            {
                MessageBox.Show("Bạn không có quyền phê duyệt. \nHoặc bộ phận tạo chưa phê duyệtt", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            }

        }

        private void btnApprove_RD_Click(object sender, RoutedEventArgs e)
        {
            if (str_depCreate == "RND" && ((ApprovalClickItem.depCreat == "RND" && ApprovalClickItem.RD == "Red")||(ApprovalClickItem.depCreat == "MARKETING" && ApprovalClickItem.MA == "DodgerBlue")))
            {
                department = "RD";
                ManagerApproval();
            }
            else
            {
                MessageBox.Show("Bạn không có quyền phê duyệt. \nHoặc bộ phận tạo chưa phê duyệt", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            }

        }

        private void btnApprove_KD_Click(object sender, RoutedEventArgs e)
        {
            if (str_depCreate == "PUR" && ApprovalClickItem.KD == "Red" && (ApprovalClickItem.MA == "DodgerBlue" || ApprovalClickItem.RD == "DodgerBlue"))
            {
                department = "KD";
                ManagerApproval();
            }
            else
            {
                MessageBox.Show("Bạn không có quyền phê duyệt. \nHoặc bộ phận tạo chưa phê duyệt", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            }
        }

        private void btnApprove_Ma_Click(object sender, RoutedEventArgs e)
        {
            if (str_depCreate == "MARKETING" && ((ApprovalClickItem.depCreat == "MARKETING" && ApprovalClickItem.MA == "Red")||(ApprovalClickItem.depCreat == "RND" && ApprovalClickItem.RD == "DodgerBlue")))
            {
                department = "MA";
                ManagerApproval();
            }
            else
            {
                MessageBox.Show("Bạn không có quyền phê duyệt. \nHoặc bộ phận tạo chưa phê duyệt", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            }
        }

        private void ckbApprove_Ma_Checked(object sender, RoutedEventArgs e)
        {
            
        }

        private void ckbApprove_Ma_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_SX_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_SX_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_CL_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_CL_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_RD_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_RD_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_KD_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ckbApprove_KD_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void cbbFillterApprove_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var click = sender as ComboBox;
                var clickItem = click.SelectedItem as ComboBoxItem;
                if (clickItem != null)
                {
                    if (clickItem.Content.ToString() == "Approve NG")
                    {
                        str_cbbFilterApprove = "Approve NG";
                        ColorRowListView(Filter_Sample_NG());
                    }
                    if (clickItem.Content.ToString() == "Approve OK")
                    {
                        str_cbbFilterApprove = "Approve OK";
                        ColorRowListView(Filter_Sample_OK());
                    }
                    if (clickItem.Content.ToString() == "Approve RE")
                    {
                        str_cbbFilterApprove = "Approve RE";
                        ColorRowListView(Filter_Sample_RE());
                    }
                    if (clickItem.Content.ToString() == "Tìm kiếm All")
                    {
                        str_cbbFilterApprove = "Tìm kiếm All";
                        ColorRowListView(Filter_Sample_All());
                    }
                    if (clickItem.Content.ToString() == "Run OK")
                    {
                        str_cbbFilterApprove = "Run OK";
                        ColorRowListView(Filter_Run_OK());
                    }
                    if (clickItem.Content.ToString() == "Run NG")
                    {
                        str_cbbFilterApprove = "Run NG";
                        ColorRowListView(Filter_Run_NG());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/cbbFillterApprove_SelectionChanged", MessageBoxButton.OK, MessageBoxImage.Error);
                
            }
            
        }

        private void txt_CustomerCode_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        public void btnSearchCustomerSample_Click(object sender, RoutedEventArgs e)
        {
            int checklenfind = txt_CustomerCode.Text.Length;
            if (checklenfind < 6)
            {
                MessageBox.Show("Mời bạn nhập ít nhất 6 ký tự để tìm kiếm", "Thông báo", MessageBoxButton.OK);
                return;
            }
            CustomerCode = txt_CustomerCode.Text;
            Page_SortCustomer page_SortCustomer = new Page_SortCustomer();
            page_SortCustomer.Show();


        }

        private void dpkStartApprove_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            str_DATESTARTAPPROVE = dpkStartApprove.SelectedDate.ToString();
        }

        private void dpkFinishApprove_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            str_DATEFINISHAPPROVE = dpkFinishApprove.SelectedDate.ToString();
        }

        private void ckbTest_Checked(object sender, RoutedEventArgs e)
        {
            CreatData();
        }

        private void ckbTest_Unchecked(object sender, RoutedEventArgs e)
        {
            ApproveData();
        }

        private void txt_FilterSample_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var command = "SELECT * from tbSampleManual WHERE (custpartcode LIKE '%" + txt_FilterSample.Text + "%' or custmodelcode LIKE '%" + txt_FilterSample.Text + "%')  AND etc2='" + str_FilterProduct + "' ORDER by INSDT desc";
                ColorRowListView(ReadDataBase(command));
            }
        }        

        private void cbb_BoxUnit_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

               
            var click = sender as ComboBox;
            var clickItem = click.SelectedItem as Helper_TaixinDB_Model;
            if (clickItem != null)
            {
                string code = clickItem.MODELCODE;
                if (code != null)
                {
                    code.Trim();
                    if (code.Length > 0)
                    {
                        FilterPaperInBox(code);
                    }
                }
            }

        }               

        private void txt_FilterBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                ProcessButtonEdit_Add();
                ProcessBoxUnit(txt_FilterBox.Text);
            }
        }

        private void rb_Manual_Checked(object sender, RoutedEventArgs e)
        {
            stackBox.Visibility = Visibility.Hidden;
            stackManul.Visibility = Visibility.Visible;
            str_FilterProduct = "Manual";
            var command = "SELECT * from tbSampleManual WHERE etc2='" + str_FilterProduct + "'  ORDER by Insdt desc";
            ColorRowListView(ReadDataBase(command));
        }

        private void rb_UnitBox_Checked(object sender, RoutedEventArgs e)
        {
            stackBox.Visibility = Visibility.Visible;
            stackManul.Visibility = Visibility.Hidden;
            str_FilterProduct = "UnitBox";
            var command = "SELECT * from tbSampleManual WHERE etc2='" + str_FilterProduct + "'  ORDER by Insdt desc";
            ColorRowListView(ReadDataBase(command));
        }

        private void rb_UnitBox_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        private void rb_Manual_Unchecked(object sender, RoutedEventArgs e)
        {

        }        

        private void cbbFilterProduc_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var click = sender as ComboBox;
            var clickItem = click.SelectedItem as ComboBoxItem;
            if (clickItem != null)
            {
                str_FilterProduct = clickItem.Content.ToString();
                var command = "SELECT * from tbSampleManual WHERE etc2='" + str_FilterProduct + "'  ORDER by Insdt desc";
                ColorRowListView(ReadDataBase(command));
            }
        }

        private void btn_FilterSample_Click(object sender, RoutedEventArgs e)
        {
            var command = "";
            var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
            var jsonDateFilterStart = JsonConvert.SerializeObject(DateTime.Parse(dateFilterStart), settings);
            string FilterStart = jsonDateFilterStart.Substring(1, jsonDateFilterStart.Length - 2);
            var jsonDateFilterFinish = JsonConvert.SerializeObject(DateTime.Parse(dateFilterFinish), settings);
            string FilterFinish = jsonDateFilterFinish.Substring(1, jsonDateFilterFinish.Length - 2);
            switch (str_cbbFilterApprove)
            {
                case "Approve NG":
                    {
                        command = "SELECT * from tbSampleManual WHERE (dep_Mar = 'N' or dep_PRO='N' or dep_QC='N' or dep_RnD='N' or dep_Pur='N') AND " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Approve OK":
                    {
                        command = "SELECT * from tbSampleManual WHERE (dep_Mar = 'Y' or dep_Mar='O') and ( dep_PRO='Y' or dep_PRO = 'O') and (dep_QC='Y' or dep_QC ='O') and " +
                            "(dep_RnD='Y' or dep_RnD ='O') and (dep_Pur = 'Y'  or dep_Pur ='O') AND " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Approve RE":
                    {
                        command = "SELECT * from tbSampleManual WHERE reject='1' AND " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Tìm kiếm All":
                    {
                        command = "SELECT * from tbSampleManual WHERE " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Run OK":
                    {
                        command = "SELECT * from tbSampleManual WHERE " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND printed = 'Y' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Run NG":
                    {
                        command = "SELECT * from tbSampleManual WHERE " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND printed is null AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }

            }
            if (str_FilterProduct == "Manual")
            {
                stackBox.Visibility = Visibility.Hidden;
            }
            else
            {
                stackBox.Visibility = Visibility.Visible;
            }
            ColorRowListView(ReadDataBase(command));
        }

        private void Dp_DateStart_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            string date = DateTime.Parse(dp_DateStart.SelectedDate.ToString()).ToString("yyyy-MM-dd");
            dateFilterStart = date + " 00:00:00";
        }

        private void Dp_DateFinish_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            string date = DateTime.Parse(dp_DateFinish.SelectedDate.ToString()).ToString("yyyy-MM-dd");
            dateFilterFinish = date + " 23:59:59";
        }

        private void btn_Reject_Click(object sender, RoutedEventArgs e)
        {
            Page_RejectSample page_RejectSample = new Page_RejectSample();
            page_RejectSample.Show();
        }        

        private void btnAttachFile_Click(object sender, RoutedEventArgs e)
        {
            if (accessProcess == "Creat" && checkDowload == false)
            {
                MainWindow.pl_Print = "Manual";
                if (checkDowload == false)
                {
                    Window_AttachFile attachFile = new Window_AttachFile();
                    attachFile.ShowDialog();
                }
                else
                {
                    Process_AttachFile();
                }
            }
            else if (accessProcess == "Approve" || checkDowload == true)
            {
                Process_AttachFile();
            }
        }

        private void btnApprove_Ma_Click_1(object sender, RoutedEventArgs e)
        {

        }        

        private void ckb_DowloadFile_Checked(object sender, RoutedEventArgs e)
        {
            checkDowload = true;
        }

        private void ckb_DowloadFile_Unchecked(object sender, RoutedEventArgs e)
        {
            checkDowload = false;
        }

        public void CheckXLS()
        {
            foreach (var item in listSampleManual)
            {
                item.checkXLS = "True";
            }
            ColorRowListView(listSampleManual);
        }
        public void UncheckXLS()
        {
            foreach (var item in listSampleManual)
            {
                item.checkXLS = "False";
            }
            ColorRowListView(listSampleManual);
        }
        private void ckb_CheckXLS_Checked(object sender, RoutedEventArgs e)
        {
            CheckXLS();
        }

        private void ckb_CheckXLS_Unchecked(object sender, RoutedEventArgs e)
        {
            UncheckXLS();
        }

        private void checkListSample_Checked(object sender, RoutedEventArgs e)
        {
            var click = sender as CheckBox;
            var clickItem = click.DataContext as Helper_TaixinDB_SampleBox;
            if (clickItem != null)
            {
                clickItem.checkXLS = "True";
            }
        }

        private void checkListSample_UnChecked(object sender, RoutedEventArgs e)
        {
            var click = sender as CheckBox;
            var clickItem = click.DataContext as Helper_TaixinDB_SampleBox;
            if (clickItem != null)
            {
                clickItem.checkXLS = "False";
            }
        }
        List<Helper_Combobox> _tempRoom = new List<Helper_Combobox>();
        List<Helper_Combobox> _tempTeam = new List<Helper_Combobox>();
        string depatment = "";
        private void cbbTypeCertification_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _tempRoom.Clear();
            _tempTeam.Clear();
            var click = sender as ComboBox;
            var clickItem = click.SelectedItem as ComboBoxItem;
            if (clickItem != null)
            {
                depatment = clickItem.Content.ToString();
                string code = "";
                switch (depatment)
                {
                    case "":
                        {
                            code = "000";
                            break;
                        }
                    case "FSC":
                        {
                            code = "002";
                            break;
                        }
                    case "PEFC":
                        {
                            code = "003";
                            break;
                        }

                }

                //if (depatment != "")
                //{
                    cbbTypeDetail.ClearValue(ComboBox.ItemsSourceProperty);

                    _tempRoom = MainWindow._ListTypeD.Where(X => X.code == code).ToList();
                    //_tempRoom.Add(new Helper_Combobox { Name_loc = "", code = "000" });
                    cbbTypeDetail.ItemsSource = _tempRoom.OrderBy(x => x.code).ToList();
                    cbbTypeDetail.SelectedIndex = 0;
                //}
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            Process_GetDataExcel();
            Process_ExportExcel();
        }

        public void Process_GetDataExcel()
        {
            string query = "SPGetDataExcelManul @date";
            string date = "";
            var listdata = DataProvider.Instance.executeQuery(path_sql, query, new object[] { date });
            
            foreach(DataRow row in listdata.Rows)
            {
                Helper_DataExcel Model = new Helper_DataExcel();
                Model.ModelCode = row["ModelCode"].ToString();
                Model.ModelName = row["ModelName"].ToString();
                Model.Ver = row["Ver"].ToString();
                Model.ItemNo = row["ItemNo"].ToString();
                Model.CustNm = row["CustNm"].ToString();
                Model.Status = row["Status"].ToString();
                Model.PurcharNm = row["PurcharNm"].ToString();
                Model.Datetime = row["Date"].ToString();
                list_Excel.Add(Model);
            }    
        }

        public async void Process_ExportExcel()
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        sfd.ShowDialog();
                        if (sfd.FileName == "")
                        {
                                MessageBox.Show("Vui lòng chọn vị trí lưu tập tin", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);
                });
                await Task.Run(() =>
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        if (sfd.FileName != "")
                        {
                            Page_LoadingData page_Loading = new Page_LoadingData();
                            stackLoading.Visibility = Visibility.Visible;
                            frameLoading.Navigate(page_Loading);
                        }
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);
                });
                await Task.Run(() =>
                {
                    Thread.Sleep(500);
                    this.Dispatcher.Invoke(() =>
                    {
                        if (sfd.FileName != "")
                        {
                            CreatListExcel();
                            File.Copy(pathFileExcel, sfd.FileName + ".xlsx");
                        }
                        stackLoading.Visibility = Visibility.Hidden;
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);
                });
            }
            catch (Exception)
            {
                    MessageBox.Show("Tên file trùng với một file có sẵn.\nVui lòng nhập một tên mới", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void CreatListExcel()
        {
            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    int numberRow = 0;
                    foreach (var item in list_Excel)
                    {
                        if (item.ModelCode != "")
                        {
                            numberRow++;
                        }
                    }

                    numberRow = numberRow + 5;
                    p.Workbook.Properties.Author = DateTime.Now.ToShortDateString();
                    p.Workbook.Properties.Title = "Danh sách Modelcode";
                    p.Workbook.Worksheets.Add("Sheet1");
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];
                    ws.Name = "Sheet1";

                    //Cột 1 
                    ws.Column(1).Width = 5;//stt
                    ws.Column(2).Width = 15;//Model Code
                    ws.Column(3).Width = 30;//Model Name
                    ws.Column(4).Width = 10;//Ver
                    ws.Column(5).Width = 20;//Mã hàng
                    ws.Column(6).Width = 30;//Khách hàng
                    ws.Column(7).Width = 15;//Tình trạng
                    ws.Column(8).Width = 15;//Người yêu cầu
                    ws.Column(9).Width = 30;

                    ws.Row(1).Height = 10;
                    ws.Row(2).Height = 40;
                    ws.Row(3).Height = 20;
                    ws.Row(4).Height = 25;


                    //căn hàng và cột cho tất cả các ô                 


                    for (int i = 1; i < numberRow; i++)
                    {
                        string strCell = "A" + i.ToString() + ":" + "J" + i.ToString();
                        var cell = ws.Cells[strCell];
                        var border = cell.Style.Border;
                        border.Bottom.Style =
                        border.Top.Style =
                        border.Left.Style =
                        border.Right.Style = ExcelBorderStyle.Thin;
                        cell.Style.WrapText = true;
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    for (int i = 5; i < numberRow; i++)
                    {
                        string strCell = "A" + i.ToString() + ":" + "J" + i.ToString();
                        var cell = ws.Cells[strCell];
                        ws.Row(i).Height = 25;
                        cell.Style.Font.Size = 11;
                        cell.Style.Font.Bold = false;

                        string strCell1 = "A" + i.ToString() + ":" + "A" + i.ToString();
                        ws.Cells[strCell1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell2 = "B" + i.ToString() + ":" + "B" + i.ToString();
                        ws.Cells[strCell2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell3 = "C" + i.ToString() + ":" + "C" + i.ToString();
                        ws.Cells[strCell3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell4 = "D" + i.ToString() + ":" + "D" + i.ToString();
                        ws.Cells[strCell4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell5 = "E" + i.ToString() + ":" + "E" + i.ToString();
                        ws.Cells[strCell5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell6 = "F" + i.ToString() + ":" + "F" + i.ToString();
                        ws.Cells[strCell6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell7 = "G" + i.ToString() + ":" + "G" + i.ToString();
                        ws.Cells[strCell7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell8 = "H" + i.ToString() + ":" + "H" + i.ToString();
                        ws.Cells[strCell8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell9 = "I" + i.ToString() + ":" + "I" + i.ToString();
                        ws.Cells[strCell9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    for (int i = 5; i < numberRow; i++)
                    {
                        if (i % 2 == 0)
                        {
                            string strCell = "A" + i.ToString() + ":" + "I" + i.ToString();
                            var cell = ws.Cells[strCell];
                            var fill = cell.Style.Fill;
                            fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            fill.BackgroundColor.SetColor(System.Drawing.Color.AliceBlue);
                        }
                    }

                    //Bôi den backgroud
                    //

                    ws.Cells["A2:J2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Cells["A2:J2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Azure);

                    ws.Cells["A4:J4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Cells["A4:J4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Ivory);


                    ws.Cells["A1:A1"].Value = "";
                    ws.Cells["A1:J1"].Merge = true;
                    ws.Cells["A1:A1"].Style.Font.Size = 25;
                    ws.Cells["A1:A1"].Style.Font.Bold = true;


                    ws.Cells["A2:A2"].Value = "DANH SÁCH MODELCODE";
                    ws.Cells["A2:J2"].Merge = true;
                    ws.Cells["A2:A2"].Style.Font.Size = 22;
                    ws.Cells["A2:A2"].Style.Font.Bold = true;

                    //Ngày SX
                    ws.Cells["A3:A3"].Value = "Ngày : " + DateTime.Now.ToString("dd/MM/yyyy") + "  Số lượng : " + (numberRow - 5);
                    ws.Cells["A3:T3"].Merge = true;
                    ws.Cells["A3:A3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    ws.Cells["A3:A3"].Style.Font.Bold = true;

                    //Head                  
                    ws.Cells["A4:J4"].Style.Font.Size = 12;
                    ws.Cells["A4:J4"].Style.Font.Bold = true;
                    ws.Cells["A4:A4"].Value = "STT";
                    ws.Cells["B4:B4"].Value = "Model Code";
                    ws.Cells["C4:C4"].Value = "Model Name";
                    ws.Cells["D4:D4"].Value = "Ver";
                    ws.Cells["E4:E4"].Value = "Mã hàng";
                    ws.Cells["F4:F4"].Value = "Khách hàng";
                    ws.Cells["G4:G4"].Value = "Tình trạng";
                    ws.Cells["H4:H4"].Value = "Người yêu cầu";
                    ws.Cells["I4:I4"].Value = "Ngày yêu cầu";

                    int index = 4;
                    int stt = 0;

                    foreach (var item in list_Excel)
                    {
                        if (item.ModelCode != "")
                        {
                            index++;
                            stt++;
                            //--
                            string strCell1 = "A" + index.ToString() + ":" + "A" + index.ToString();
                            ws.Cells[strCell1].Value = stt;
                            //--
                            string strCell2 = "B" + index.ToString() + ":" + "B" + index.ToString();
                            ws.Cells[strCell2].Value = item.ModelCode;
                            //--
                            string strCell3 = "C" + index.ToString() + ":" + "C" + index.ToString();
                            ws.Cells[strCell3].Value = item.ModelName;
                            //--
                            string strCell4 = "D" + index.ToString() + ":" + "D" + index.ToString();
                            ws.Cells[strCell4].Value = item.Ver;
                            //--
                            string strCell5 = "E" + index.ToString() + ":" + "E" + index.ToString();
                            ws.Cells[strCell5].Value = item.ItemNo;
                            //--
                            string strCell6 = "F" + index.ToString() + ":" + "F" + index.ToString();
                            ws.Cells[strCell6].Value = item.CustNm;
                            //--
                            string strCell7 = "G" + index.ToString() + ":" + "G" + index.ToString();
                            ws.Cells[strCell7].Value = item.Status;
                            //--
                            string strCell8 = "H" + index.ToString() + ":" + "H" + index.ToString();
                            ws.Cells[strCell8].Value = item.PurcharNm;
                            //--
                            string strCell9 = "I" + index.ToString() + ":" + "I" + index.ToString();
                            ws.Cells[strCell9].Value = item.Datetime;
                        }
                    }
                    ws.PrinterSettings.PaperSize = ePaperSize.A4;
                    ws.PrinterSettings.Orientation = eOrientation.Landscape;
                    ws.PrinterSettings.FitToPage = true;
                    ws.Cells["A4:I4"].AutoFilter = true;
                    ws.PrinterSettings.TopMargin = Decimal.Parse("0");
                    ws.PrinterSettings.LeftMargin = Decimal.Parse("0.25");
                    ws.PrinterSettings.BottomMargin = Decimal.Parse("0.25");
                    ws.PrinterSettings.RightMargin = Decimal.Parse("0.25");
                    File.Delete(pathFileExcel);
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(pathFileExcel, bin);
                    //exportFileExcel = false;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "CreatListExcel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
