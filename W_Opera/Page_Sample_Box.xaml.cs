using DataHelper;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Tulpep.NotificationWindow;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Page_Sample_Box.xaml
    /// </summary>
    public partial class Page_Sample_Box : Page
    {
        #region Khai báo
        public static string path_sql_attach = "Data Source=192.168.2.10;Initial Catalog=taixin_attach;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        string path_sql= "Data Source=192.168.2.10;Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
        List<Helper_TaixinDB_Model> ListAllDataSample = new List<Helper_TaixinDB_Model>();
        List<Helper_DataButton> listButtonTop = new List<Helper_DataButton>();
        public static Helper_TaixinDB_SampleBox ApprovalClickItem = new Helper_TaixinDB_SampleBox();
        Helper_AccessManger access_db = new Helper_AccessManger();
        List<Helper_AccessManger> list_access = new List<Helper_AccessManger>();
        PopupNotifier popup = new PopupNotifier();
        DispatcherTimer dt = new DispatcherTimer();
        OpenFileDialog ofd;
        public static DataBaseHelper db = new DataBaseHelper();
        List<Helper_TaixinDB_SampleBox> listSampleBox = new List<Helper_TaixinDB_SampleBox>();
        int qtyFileUpload = 0;
        string date = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");
        string str_cbbFilterApprove = "Tìm kiếm All";
        string dateFilterStart = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyy-MM-dd") + " 00:00:00";
        string dateFilterFinish = DateTime.Parse(DateTime.Now.ToString()).ToString("yyyy-MM-dd") + " 23:59:59";       
        string qtyAttachFile = "0";
        string accessProcess = "Approve";       
        string processButton = "";
        string str_depCreate = "";
        string dateInput = "";
        bool checkPopup = false;
        bool checkRun = false;
        bool checkDowload = false;
        bool checkView = false;        
        string dateApproval = "";
        string department = "";
        string str_cmpcode;
        string str_bizdiv;
        string str_samno;
        string str_modelcode;
        string str_modelname;
        string str_applydt;
        string str_version;
        string str_custpartcode;
        string str_custpart_version;
        string str_custpartcode_ver;
        string str_custmodelcode;
        string str_repmodelcode;
        string str_useflag;
        string str_modelgroup;
        string str_modeldiv;
        string str_modeltype;
        string str_modelchild;
        string str_cust_gb;
        string str_information;
        string str_paper_name1;
        string str_paper_size1;
        string str_paper_scale1;
        string str_paper_name2;
        string str_paper_size2;
        string str_paper_scale2;
        string str_paper_name3;
        string str_paper_size3;
        string str_paper_scale3;
        string str_cover_up_name1;
        string str_cover_up_size1;
        string str_cover_up_scale1;
        string str_cover_up_name2;
        string str_cover_up_size2;
        string str_cover_up_scale2;
        string str_cover_up_name3;
        string str_cover_up_size3;
        string str_cover_up_scale3;
        string str_cover_up_name4;
        string str_cover_up_size4;
        string str_cover_up_scale4;
        string str_cover_up_name5;
        string str_cover_up_size5;
        string str_cover_up_scale5;
        string str_fullsize;
        string str_print_color1;
        string str_print_color2;
        string str_print_color3;
        string str_print_color4;
        string str_print_color5;
        string str_coating;
        string str_glossy_color;
        string str_glossy_detail;
        string str_holo_color1;
        string str_holo_detail1;
        string str_holo_color2;
        string str_holo_detail2;
        string str_holo_color3;
        string str_holo_detail3;
        string str_holo_color4;
        string str_holo_detail4;
        string str_debosing_detail;
        string str_imbosing_detail;
        string str_boi_detaiil;
        string str_boi_color;
        string str_qty;
        string str_forecast;
        string str_remark1;
        string str_remark2;
        string str_dep_mar;
        string str_dep_pro;
        string str_dep_qc;
        string str_dep_rnd;
        string str_dep_pur;
        string str_paperindex1 = "0";
        string str_paperindex2 = "1";
        string str_paperindex3 = "2";
        string str_printindex1 = "0";
        string str_printindex2 = "1";
        string str_printindex3 = "2";
        string str_reject;
        string str_qtyAttach;
        string str_etc3;
        string str_etc4;
        string str_etc5;
        string str_etc6;
        string str_etc7;
        string str_etc8;
        string str_etc9;
        string str_etc10;
        string str_sadt;
        string str_fadt;
        string str_imsempcode;
        string str_insdt;
        string str_updempcode;
        string str_upddt;
        public List<string> ListSpecPaper { get; set; }
        public List<string> ListSpecColor { get; set; }



        #endregion
        public Page_Sample_Box()
        {
            InitializeComponent();
            CreatAllButtonEdit();
            Loaded += Page_Sample_Box_Loaded;
        }
        
        private void Page_Sample_Box_Loaded(object sender, RoutedEventArgs e)
        {
            path_sql = "Data Source=" + MainWindow.ip + ";Initial Catalog=taixin_db;Persist Security Info=True;User ID=sa;Password= oneuser1!";
            path_sql_attach = MainWindow.path_sql_attach;
            dpkStartApprove.SelectedDate = DateTime.Now;
            dpkFinishApprove.SelectedDate = DateTime.Now;
            dp_DateStart.SelectedDate = DateTime.Now;
            dp_DateFinish.SelectedDate = DateTime.Now;
            str_applydt = DateTime.Now.Year.ToString("0000") + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00");
            AccessManage();
            GetDataDeptUser();
            Filter_Sample_All();           
            GetDataSpec();            
        }
        
        public void GetDataSpec()
        {
            ListSpecPaper = new List<string>();
            ListSpecColor = new List<string>();
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var commandPaper = "SELECT * FROM tbSampleSpecPaper where typeDep='Box' and typeDiv='Paper' ORDER by ID ASC";
                    var commandColor = "SELECT * FROM tbSampleSpecPaper where typeDep='Box' and typeDiv='Color' ORDER by ID ASC";
                    using (SqlCommand cmd = new SqlCommand(commandPaper, conn))
                    {
                        cmd.CommandTimeout = 0;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                if (dr[0].ToString() != "")
                                {
                                    ListSpecPaper.Add(dr[3].ToString());
                                }
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand(commandColor, conn))
                    {
                        cmd.CommandTimeout = 0;
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                if (dr[0].ToString() != "")
                                {
                                    ListSpecColor.Add(dr[3].ToString());
                                }
                            }
                        }
                    }
                    DataContext = this;
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/GetDataSpec", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
            //listButtonTop.Add(new Helper_DataButton
            //{
            //    ID = 6,
            //    ContentButton = "Check",
            //    ImageSource = "Image/Edit/check.png",
            //    BackGroundColor = PinValue.OFF
            //});
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
                            processButton = "Run";
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
                MessageBox.Show(ex.Message, "Box/AccessManage", MessageBoxButton.OK, MessageBoxImage.Error);                
            }            
        }

        
        public string ReadSampleNo()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "SELECT max(samno) FROM tbSampleBox Where applydt='" + date + "'";
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
                MessageBox.Show(ex.Message, "Box/ReadSampleNo", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }
       
       
        public void Add_NewData()
        {
            try
            {
                
                checkView = true;
                var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                var jsonDateStart = JsonConvert.SerializeObject(DateTime.Parse(str_sadt), settings);
                var jsonDateFinish = JsonConvert.SerializeObject(DateTime.Parse(str_fadt), settings);
                dateInput = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                string dateStart = jsonDateStart.Substring(1, jsonDateStart.Length - 2);
                string dateFinish = jsonDateFinish.Substring(1, jsonDateFinish.Length - 2);
                str_cmpcode = "02";
                str_bizdiv = "300";
                str_samno = ReadSampleNo();
                str_modelcode = txt_modelcode.Text;
                str_modelname = "";                
                str_version = txt_version.Text;
                str_custpartcode = txt_cuspartcode.Text;
                str_custpart_version = "";
                str_custpartcode_ver = "";
                str_custmodelcode = "";
                str_repmodelcode = "";
                str_useflag = "";
                str_modelgroup = "";
                str_modeldiv = "";
                str_modeltype = "";
                str_modelchild = "";
                str_cust_gb = txt_cusgb.Text;
                str_information = txt_infomation.Text;
                str_paper_name1 = txt_SpecPaper_Name1.Text;
                str_paper_size1 = txt_SpecPaper_Size1.Text;
                str_paper_scale1 = txt_SpecPaper_Scale1.Text;
                str_paper_name2 = txt_SpecPaper_Name2.Text;
                str_paper_size2 = txt_SpecPaper_Size2.Text;
                str_paper_scale2 = txt_SpecPaper_Scale2.Text;
                str_paper_name3 = txt_SpecPaper_Name3.Text;
                str_paper_size3 = txt_SpecPaper_Size3.Text;
                str_paper_scale3 = txt_SpecPaper_Scale3.Text;
                str_cover_up_name1 = txt_SpecCover_Up_Name.Text;
                str_cover_up_size1 = txt_SpecCover_Up_Size.Text;
                str_cover_up_scale1 = txt_SpecCover_Up_Scale.Text;
                str_cover_up_name2 = txt_SpecCover_Lo_Name.Text;
                str_cover_up_size2 = txt_SpecCover_Lo_Size.Text;
                str_cover_up_scale2 = txt_SpecCover_Lo_Scale.Text;
                str_cover_up_name3 = txt_SpecCover_Cover_Name.Text;
                str_cover_up_size3 = txt_SpecCover_Cover_Name.Text;
                str_cover_up_scale3 = txt_SpecCover_Cover_Scale.Text;
                str_cover_up_name4 = txt_SpecCover_Mid_Name.Text;
                str_cover_up_size4 = txt_SpecCover_Mid_Size.Text;
                str_cover_up_scale4 = txt_SpecCover_Mid_Scale.Text;
                str_cover_up_name5 = "";
                str_cover_up_size5 = "";
                str_cover_up_scale5 = "";
                str_fullsize = txt_SpecFullSize.Text;
                str_print_color1 = txt_Process_Color1.Text;
                str_print_color2 = txt_Process_Color2.Text;
                str_print_color3 = txt_Process_Color3.Text;
                str_print_color4 = "";
                str_print_color5 = "";
                str_coating = txt_Process_Coating.Text;
                str_glossy_color = txt_Process_Glosy_Name.Text;
                str_glossy_detail = txt_Process_Glosy_Detail.Text;
                str_holo_color1 = txt_Process_Holo_Color1.Text;
                str_holo_detail1 = txt_Process_Holo_Detail1.Text;
                str_holo_color2 = txt_Process_Holo_Color2.Text;
                str_holo_detail2 = txt_Process_Holo_Detail2.Text;
                str_holo_color3 = txt_Process_Holo_Color3.Text;
                str_holo_detail3 = txt_Process_Holo_Detail3.Text;
                str_holo_color4 = txt_Process_Holo_Color4.Text;
                str_holo_detail4 = txt_Process_Holo_Detail4.Text;
                str_debosing_detail = txt_Process_Debo_Detail.Text;
                str_imbosing_detail = txt_Process_Imbo_Detail.Text;
                str_boi_detaiil = txt_Process_Boi_Name.Text;
                str_boi_color = txt_Process_Boi_Color.Text;
                str_qty = txt_Note_Qty.Text;
                str_forecast = txt_Note_Forecast.Text;
                str_remark1 = txt_Note_Detail.Text;
                str_remark2 = txt_Note_Remark.Text;
                str_dep_mar = "N";
                //sx
                if (ckbApprove_SX.IsChecked == true)
                {
                    str_dep_pro = "N";
                }
                else
                {
                    str_dep_pro = "O";
                }
                //cl
                if (ckbApprove_CL.IsChecked == true)
                {
                    str_dep_qc = "N";
                }
                else
                {
                    str_dep_qc = "O";
                }
                //rd
                if (ckbApprove_RD.IsChecked == true)
                {
                    str_dep_rnd = "N";
                }
                else
                {
                    str_dep_rnd = "O";
                }
                //kd
                if (ckbApprove_KD.IsChecked == true)
                {
                    str_dep_pur = "N";
                }
                else
                {
                    str_dep_pur = "O";
                }
                str_reject = "0";
                str_qtyAttach = Window_AttachFile.listAttachFile.Count.ToString();
                str_etc3 = "";
                str_etc4 = "";
                str_etc5 = "";
                str_etc6 = "";
                str_etc7 = "";
                str_etc8 = "";
                str_etc9 = "";
                str_etc10 = "";
                str_sadt = dateStart;
                str_fadt = dateFinish;
                str_imsempcode = txt_imsempcode.Text;
                str_insdt = dateInput;
                str_updempcode = "";
                str_upddt = "";
                dateApproval = dateInput;      
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/Add_NewData", MessageBoxButton.OK, MessageBoxImage.Error);
                
            }
        }

      
        public void ProcessButtonEdit_Add()
        {
            try
            {
                SampleBox.box.modelcode = "";
                SampleBox.box.modelname = "";
                SampleBox.box.applydt = "";
                SampleBox.box.version = "";
                SampleBox.box.custpartcode = "";
                SampleBox.box.custpart_version = "";
                SampleBox.box.custpartcode_ver = "";
                SampleBox.box.custmodelcode = "";
                SampleBox.box.repmodelcode = "";
                SampleBox.box.useflag = "";
                SampleBox.box.modelgroup = "";
                SampleBox.box.modeldiv = "";
                SampleBox.box.modeltype = "";
                SampleBox.box.modelchild = "";
                SampleBox.box.cust_gb = "";
                SampleBox.box.information = "";
                SampleBox.box.paper_name1 = "";
                SampleBox.box.paper_size1 = "";
                SampleBox.box.paper_scale1 = "";
                SampleBox.box.paper_name2 = "";
                SampleBox.box.paper_size2 = "";
                SampleBox.box.paper_scale2 = "";
                SampleBox.box.paper_name3 = "";
                SampleBox.box.paper_size3 = "";
                SampleBox.box.paper_scale3 = "";
                SampleBox.box.cover_up_name1 = "";
                SampleBox.box.cover_up_size1 = "";
                SampleBox.box.cover_up_scale1 = "";
                SampleBox.box.cover_up_name2 = "";
                SampleBox.box.cover_up_size2 = "";
                SampleBox.box.cover_up_scale2 = "";
                SampleBox.box.cover_up_name3 = "";
                SampleBox.box.cover_up_size3 = "";
                SampleBox.box.cover_up_scale3 = "";
                SampleBox.box.cover_up_name4 = "";
                SampleBox.box.cover_up_size4 = "";
                SampleBox.box.cover_up_scale4 = "";
                SampleBox.box.cover_up_name5 = "";
                SampleBox.box.cover_up_size5 = "";
                SampleBox.box.cover_up_scale5 = "";
                SampleBox.box.fullsize = "";
                SampleBox.box.print_color1 = "";
                SampleBox.box.print_color2 = "";
                SampleBox.box.print_color3 = "";
                SampleBox.box.print_color4 = "";
                SampleBox.box.print_color5 = "";
                SampleBox.box.coating = "";
                SampleBox.box.glossy_color = "";
                SampleBox.box.glossy_detail = "";
                SampleBox.box.holo_color1 = "";
                SampleBox.box.holo_detail1 = "";
                SampleBox.box.holo_color2 = "";
                SampleBox.box.holo_detail2 = "";
                SampleBox.box.holo_color3 = "";
                SampleBox.box.holo_detail3 = "";
                SampleBox.box.holo_color4 = "";
                SampleBox.box.holo_detail4 = "";
                SampleBox.box.debosing_detail = "";
                SampleBox.box.imbosing_detail = "";
                SampleBox.box.boi_detaiil = "";
                SampleBox.box.boi_color = "";
                SampleBox.box.qty = "";
                SampleBox.box.forecast = "";
                SampleBox.box.remark1 = "";
                SampleBox.box.remark2 = "";
                SampleBox.box.dep_mar = "";
                SampleBox.box.dep_pro = "";
                SampleBox.box.dep_qc = "";
                SampleBox.box.dep_rnd = "";
                SampleBox.box.dep_pur = "";
                SampleBox.box.paperindex1 = "0";
                SampleBox.box.paperindex2 = "1";
                SampleBox.box.paperindex3 = "2";
                SampleBox.box.printindex1 = "0";
                SampleBox.box.printindex2 = "1";
                SampleBox.box.printindex3 = "2";
                SampleBox.box.reject = "";
                SampleBox.box.etc2 = "";
                SampleBox.box.etc3 = "";
                SampleBox.box.etc4 = "";
                SampleBox.box.etc5 = "";
                SampleBox.box.etc6 = "";
                SampleBox.box.etc7 = "";
                SampleBox.box.etc8 = "";
                SampleBox.box.etc9 = "";
                SampleBox.box.etc10 = "";
                SampleBox.box.sadt = "";
                SampleBox.box.fadt = "";
                SampleBox.box.imsempcode = "";
                SampleBox.box.insdt = "";
                SampleBox.box.updempcode = "";
                SampleBox.box.upddt = "";
                SampleBox.box.qtyAttach = "0 File";
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
                MessageBox.Show(ex.Message, "Box/ProcessButtonEdit_Add", MessageBoxButton.OK, MessageBoxImage.Error);

            }


        }      

        public void ProcessButtonEdit_Del()
        {
            if (MessageBoxResult.Yes == MessageBox.Show("Bạn muốn xóa dữ liệu?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question))
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(path_sql))
                    {
                        conn.Open();
                        var command1 = "DELETE tbSampleBox WHERE samno =" + "'" + ApprovalClickItem.samno + "'";                       
                        var command3 = "DELETE tbSampleReject WHERE typeSample ='Box' and samno =" + "'" + ApprovalClickItem.samno + "'";
                        using (SqlCommand cmd = new SqlCommand(command1, conn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.ExecuteNonQuery();
                        }                       
                        using (SqlCommand cmd = new SqlCommand(command3, conn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.ExecuteNonQuery();
                        }                        
                        conn.Close();

                    }
                    using (SqlConnection conn = new SqlConnection(path_sql_attach))
                    {
                        conn.Open();
                        var command2 = "DELETE tbSampleAttach WHERE typeSample = 'Box' and samno =" + "'" + ApprovalClickItem.samno + "'";                        
                        using (SqlCommand cmd = new SqlCommand(command2, conn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.ExecuteNonQuery();
                        }                        
                        conn.Close();
                    }
                    ProcessSampleHistory(ApprovalClickItem.samno, "Delete");
                    Filter_Sample_All();

                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, "Box/ProcessButtonEdit_Del", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        }

        public void ProcessButtonEdit_Edit()
        {
            if (MessageBoxResult.Yes == MessageBox.Show("Bạn muốn sửa dữ liệu?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question))
            {
                try
                {
                    Add_NewData();
                    using (SqlConnection conn = new SqlConnection(path_sql))
                    {
                        conn.Open();
                        //sx

                        //var command = "UPDATE SAMPLEAPPROVE SET model = '" + model + "',modelcode = '" + modelcode + "',customer = '" + customer + "',ver = '" + ver + "'" +
                        //    ",sx = '" + sx + "',cl = '" + cl + "',rd = '" + rd + "',kd = '" + kd + "' WHERE ID = '"+ApproverClickItem.ID+"'";

                        var command = "UPDATE tbSampleBox SET modelcode = '" + str_modelcode + "',modelname = '" + str_modelname + "',applydt = '" + str_applydt + "',version = '" + str_version + "',custpartcode = '" + str_custpartcode + "',custpart_version = '" + str_custpart_version + "',custpartcode_ver = '" + str_custpartcode_ver + "',custmodelcode = '" + str_custmodelcode + "',repmodelcode = '" + str_repmodelcode + "',useflag = '" + str_useflag + "',modelgroup = '" + str_modelgroup + "',modeldiv = '" + str_modeldiv + "',modeltype = '" + str_modeltype + "',modelchild = '" + str_modelchild + "',cust_gb = '" + str_cust_gb + "',information = '" + str_information + "',paper_name1 = '" + str_paper_name1 + "',paper_size1 = '" + str_paper_size1 + "',paper_scale1 = '" + str_paper_scale1 + "',paper_name2 = '" + str_paper_name2 + "',paper_size2 = '" + str_paper_size2 + "',paper_scale2 = '" + str_paper_scale2 + "',paper_name3 = '" + str_paper_name3 + "',paper_size3 = '" + str_paper_size3 + "',paper_scale3 = '" + str_paper_scale3 + "',cover_up_name1 = '" + str_cover_up_name1 + "',cover_up_size1 = '" + str_cover_up_size1 + "',cover_up_scale1 = '" + str_cover_up_scale1 + "',cover_up_name2 = '" + str_cover_up_name2 + "',cover_up_size2 = '" + str_cover_up_size2 + "',cover_up_scale2 = '" + str_cover_up_scale2 + "',cover_up_name3 = '" + str_cover_up_name3 + "',cover_up_size3 = '" + str_cover_up_size3 + "',cover_up_scale3 = '" + str_cover_up_scale3 + "',cover_up_name4 = '" + str_cover_up_name4 + "',cover_up_size4 = '" + str_cover_up_size4 + "',cover_up_scale4 = '" + str_cover_up_scale4 + "',cover_up_name5 = '" + str_cover_up_name5 + "',cover_up_size5 = '" + str_cover_up_size5 + "',cover_up_scale5 = '" + str_cover_up_scale5 + "',fullsize = '" + str_fullsize + "',print_color1 = '" + str_print_color1 + "',print_color2 = '" + str_print_color2 + "',print_color3 = '" + str_print_color3 + "',print_color4 = '" + str_print_color4 + "',print_color5 = '" + str_print_color5 + "',coating = '" + str_coating + "',glossy_color = '" + str_glossy_color + "',glossy_detail = '" + str_glossy_detail + "',holo_color1 = '" + str_holo_color1 + "',holo_detail1 = '" + str_holo_detail1 + "',holo_color2 = '" + str_holo_color2 + "',holo_detail2 = '" + str_holo_detail2 + "',holo_color3 = '" + str_holo_color3 + "',holo_detail3 = '" + str_holo_detail3 + "',debosing_detail = '" + str_debosing_detail + "',imbosing_detail = '" + str_imbosing_detail + "',boi_detaiil = '" + str_boi_detaiil + "',boi_color = '" + str_boi_color + "',qty = '" + str_qty + "',forecast = '" + str_forecast + "',remark1 = N'" + str_remark1 + "',remark2 = N'" + str_remark2 + "',dep_mar = '" + str_dep_mar + "',dep_pro = '" + str_dep_pro + "',dep_qc = '" + str_dep_qc + "',dep_rnd = '" + str_dep_rnd + "',dep_pur = '" + str_dep_pur + "',paperindex1 = '" + str_paperindex1 + "',paperindex2 = '" + str_paperindex2 + "',paperindex3 = '" + str_paperindex3 + "',printindex1 = '" + str_printindex1 + "',printindex2 = '" + str_printindex2 + "',printindex3 = '" + str_printindex3 + "',reject = '" + str_reject + "',etc2 = '" + str_qtyAttach + "',etc3 = '" + str_etc3 + "',etc4 = '" + str_holo_color4 + "',etc5 = '" + str_holo_detail4 + "',etc6 = '" + str_etc6 + "',etc7 = '" + str_etc7 + "',etc8 = '" + str_etc8 + "',etc9 = '" + str_etc9 + "',etc10 = '" + str_etc10 + "',sadt = '" + str_sadt + "',fadt = '" + str_fadt + "',updempcode = '" + txt_imsempcode.Text + "',upddt = '" + dateInput + "' where samno = " + "'" + ApprovalClickItem.samno + "'";
                        using (SqlCommand cmd = new SqlCommand(command, conn))
                        {
                            cmd.CommandTimeout = 0;
                            cmd.ExecuteNonQuery();
                        }
                        conn.Close();
                        Filter_Sample_All();
                        MessageBox.Show("Dữ liệu sửa Thành Công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    ProcessSampleHistory(ApprovalClickItem.samno, "Edit");

                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, "Box/ProcessButtonEdit_Edit", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        
        public async Task ProcessButtonEdit_Save()
        {
            try
            {
                frameLoading.Visibility = Visibility.Visible;
                Add_NewData();
                await Process_AttachFile();
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();
                    var command = "INSERT tbSampleBox(cmpcode,bizdiv,samno,modelcode,modelname,applydt,version,custpartcode,custpart_version,custpartcode_ver,custmodelcode,repmodelcode,useflag,modelgroup,modeldiv,modeltype,modelchild,cust_gb,information,paper_name1,paper_size1,paper_scale1,paper_name2,paper_size2,paper_scale2,paper_name3,paper_size3,paper_scale3,cover_up_name1,cover_up_size1,cover_up_scale1,cover_up_name2,cover_up_size2,cover_up_scale2,cover_up_name3,cover_up_size3,cover_up_scale3,cover_up_name4,cover_up_size4,cover_up_scale4,cover_up_name5,cover_up_size5,cover_up_scale5,fullsize,print_color1,print_color2,print_color3,print_color4,print_color5,coating,glossy_color,glossy_detail,holo_color1,holo_detail1,holo_color2,holo_detail2,holo_color3,holo_detail3,debosing_detail,imbosing_detail,boi_detaiil,boi_color,qty,forecast,remark1,remark2,dep_mar,dep_pro,dep_qc,dep_rnd,dep_pur,paperindex1,paperindex2,paperindex3,printindex1,printindex2,printindex3,reject,qtyAttach,depCreate,etc3,etc4,etc5,etc6,etc7,etc8,etc9,etc10,sadt,fadt,imsempcode,insdt,updempcode,upddt)" +
                       " VALUES(N'" + str_cmpcode + "',N'" + str_bizdiv + "',N'" + str_samno + "',N'" + str_modelcode + "',N'" + str_modelname + "',N'" + str_applydt + "',N'" + str_version + "',N'" + str_custpartcode + "',N'" + str_custpart_version + "',N'" + str_custpartcode_ver + "',N'" + str_custmodelcode + "',N'" + str_repmodelcode + "',N'" + str_useflag + "',N'" + str_modelgroup + "',N'" + str_modeldiv + "',N'" + str_modeltype + "',N'" + str_modelchild + "',N'" + str_cust_gb + "',N'" + str_information + "',N'" + str_paper_name1 + "',N'" + str_paper_size1 + "',N'" + str_paper_scale1 + "',N'" + str_paper_name2 + "',N'" + str_paper_size2 + "',N'" + str_paper_scale2 + "',N'" + str_paper_name3 + "',N'" + str_paper_size3 + "',N'" + str_paper_scale3 + "',N'" + str_cover_up_name1 + "',N'" + str_cover_up_size1 + "',N'" + str_cover_up_scale1 + "',N'" + str_cover_up_name2 + "',N'" + str_cover_up_size2 + "',N'" + str_cover_up_scale2 + "',N'" + str_cover_up_name3 + "',N'" + str_cover_up_size3 + "',N'" + str_cover_up_scale3 + "',N'" + str_cover_up_name4 + "',N'" + str_cover_up_size4 + "',N'" + str_cover_up_scale4 + "',N'" + str_cover_up_name5 + "',N'" + str_cover_up_size5 + "',N'" + str_cover_up_scale5 + "',N'" + str_fullsize + "',N'" + str_print_color1 + "',N'" + str_print_color2 + "',N'" + str_print_color3 + "',N'" + str_print_color4 + "',N'" + str_print_color5 + "',N'" + str_coating + "',N'" + str_glossy_color + "',N'" + str_glossy_detail + "',N'" + str_holo_color1 + "',N'" + str_holo_detail1 + "',N'" + str_holo_color2 + "',N'" + str_holo_detail2 + "',N'" + str_holo_color3 + "',N'" + str_holo_detail3 + "',N'" + str_debosing_detail + "',N'" + str_imbosing_detail + "',N'" + str_boi_detaiil + "',N'" + str_boi_color + "',N'" + str_qty + "',N'" + str_forecast + "',N'" + str_remark1 + "',N'" + str_remark2 + "',N'" + str_dep_mar + "',N'" + str_dep_pro + "',N'" + str_dep_qc + "',N'" + str_dep_rnd + "',N'" + str_dep_pur + "',N'" + str_paperindex1 + "',N'" + str_paperindex2 + "',N'" + str_paperindex3 + "',N'" + str_printindex1 + "',N'" + str_printindex2 + "',N'" + str_printindex3 + "',N'" + str_reject + "',N'" + str_qtyAttach + "',N'"+str_depCreate+"',N'" + str_etc3 + "',N'" + str_holo_color4 + "',N'" + str_holo_detail4 + "',N'" + str_etc6 + "',N'" + str_etc7 + "',N'" + str_etc8 + "',N'" + str_etc9 + "',N'" + str_etc10 + "',N'" + str_sadt + "',N'" + str_fadt + "',N'" + str_imsempcode + "',N'" + str_insdt + "',N'" + str_updempcode + "',N'" + str_upddt + "')";

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.ExecuteNonQuery();
                    }
                    Window_AttachFile.listAttachFile.Clear();
                    Filter_Sample_All();            
                    conn.Close();
                }
                ProcessSampleHistory(ApprovalClickItem.samno, "Save");
               
                //MessageBox.Show("Lưu mẫu thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/ProcessButtonEdit_Save", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void ProcessButtonEdit_Printer()
        {
            MainWindow.pl_Print = "Box";
            MainWindow.print.Show();
            MainWindow.checkPrint = true;
            SampleBox.listSampleExportExcel = listSampleBox;
        }

        public void ProcessButtonEdit_Run()
        {
            try
            { 
                if(MessageBox.Show("Bạn có muốn xác nhận Sample này không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Warning,MessageBoxResult.Yes)==MessageBoxResult.Yes)
                {
                    if (checkRun == true)
                    {
                        using (SqlConnection conn = new SqlConnection(path_sql))
                        {
                            conn.Open();

                            var command = "Update tbSampleBox SET printed='Y' where samno='" + SampleBox.box.samno + "' ";
                            using (SqlCommand cmd = new SqlCommand(command, conn))
                            {
                                cmd.CommandTimeout = 0;
                                cmd.ExecuteNonQuery();
                            }
                            Filter_Sample_All();
                            ProcessSampleHistory(SampleBox.box.samno, "Printed");
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

                            var command = "Update tbSampleBox SET printed='' where samno='" + SampleBox.box.samno + "' ";
                            using (SqlCommand cmd = new SqlCommand(command, conn))
                            {
                                cmd.CommandTimeout = 0;
                                cmd.ExecuteNonQuery();
                            }
                            Filter_Sample_All();
                            ProcessSampleHistory(SampleBox.box.samno, "Printed");
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

        public void ProcessSampleHistory(string samno, string ope)
        {
            using (SqlConnection conn = new SqlConnection(path_sql))
            {
                conn.Open();
                var command = "Insert tbSampleHistory(cmpcode,bizdiv,samno,modelcode,typeSample,applydt,operator,imsempcode,insdt) values( '02','300','" + samno + "','" + txt_modelcode.Text + "','Box','" + str_applydt + "','" + ope + "','" + MainWindow.UserLogin + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                using (SqlCommand cmd = new SqlCommand(command, conn))
                {
                    cmd.CommandTimeout = 0;
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }
        }

        public void ProcessUploadFile(){
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

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
                       
            Application.Current.Dispatcher.BeginInvoke(
            DispatcherPriority.Background,
            new Action(async () =>
            {                
               
            }));
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
                           await  UploadAttachFile(item.FileName, item.Path, str_samno, item.Stt);
                        }
                    }
                    str_qtyAttach = Window_AttachFile.listAttachFile.Count.ToString();
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


       
        public void ColorRowListView(List<Helper_TaixinDB_SampleBox> listAllData_Input)
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
                            if (item.reject == "1" && str_depCreate == "MARKETING")
                            {
                                item.dep_mar = "Purple";
                            }
                            if (item.reject == "1" && str_depCreate == "RND")
                            {
                                item.dep_rnd = "Purple";
                            }
                            index++;
                        }

                        foreach (var item in listAllData_Input)
                        {
                            lvApproveSample.Items.Add(item);
                        }
                        if (checkView == true)
                            lvApproveSample.SelectedIndex = 0;

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/ColorRowListView", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    switch (department)
                    {
                        case "SX":
                            {
                                command = "UPDATE tbSampleBox set dep_PRO = 'Y',INSDT = '" + dateInput + "' where SAMNO = " + " '" + ApprovalClickItem.samno + "'";
                                break;
                            }
                        case "CL":
                            {
                                command = "UPDATE tbSampleBox set dep_QC = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.samno + "'";
                                break;
                            }
                        case "RD":
                            {
                                command = "UPDATE tbSampleBox set dep_RND = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.samno + "'";
                                break;
                            }
                        case "KD":
                            {
                                command = "UPDATE tbSampleBox set dep_PUR = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.samno + "'";
                                break;
                            }
                        case "MA":
                            {
                                command = "UPDATE tbSampleBox set dep_MAR = 'Y',INSDT = '" + dateInput + "'  where SAMNO =" + " '" + ApprovalClickItem.samno + "'";
                                break;
                            }
                    }

                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.ExecuteNonQuery();
                    }
                    Filter_Sample_All();
                    MessageBox.Show("Mẫu được phê duyệt Thành Công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    conn.Close();
                }
                ProcessSampleHistory(ApprovalClickItem.samno, "Arv-"+department);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/ManagerApproval", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        
        public void Filter_Sample_All()
        {
            var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
            var jsonDateFilterStart = JsonConvert.SerializeObject(DateTime.Parse(dateFilterStart), settings);
            string FilterStart = jsonDateFilterStart.Substring(1, jsonDateFilterStart.Length - 2);
            var jsonDateFilterFinish = JsonConvert.SerializeObject(DateTime.Parse(dateFilterFinish), settings);
            string FilterFinish = jsonDateFilterFinish.Substring(1, jsonDateFilterFinish.Length - 2);
            var command = "SELECT * FROM tbSampleBox where INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "' ORDER BY Insdt DESC";
            listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
            ColorRowListView(listSampleBox);
        }
       
        public void Filter_Sample_OK()
        {
            var command = "SELECT * from tbSampleBox where (dep_MAR = 'Y' or dep_MAR='O') and ( dep_PRO='Y' or dep_PRO = 'O') and " +
                "(dep_QC='Y' or dep_QC ='O') and (dep_RND='Y' or dep_RND ='O') and (dep_PUR = 'Y'  or dep_PUR ='O') ORDER BY Insdt DESC";
            listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
            ColorRowListView(listSampleBox);
        }
        
        
        public void Filter_Sample_NG()
        {
            var command = "SELECT * from tbSampleBox where (dep_MAR = 'N' or dep_PRO='N' or dep_QC='N' or dep_RND='N' or dep_PUR='N') ORDER BY Insdt DESC";
            listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
            ColorRowListView(listSampleBox);
        }
        
        
        public void Filter_Sample_RE()
        {
            var command = "SELECT * from tbSampleBox where reject='1' ORDER BY Insdt DESC";
            listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
            ColorRowListView(listSampleBox);
        }
        
        
        public void Filter_Run_OK()
        {
            var command = "SELECT * from tbSampleBox where printed='Y' ORDER BY Insdt DESC";
            listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
            ColorRowListView(listSampleBox);
        }
        
        
        public void Filter_Run_NG()
        {
            var command = "SELECT * from tbSampleBox where printed is null ORDER BY Insdt DESC";
            listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
            ColorRowListView(listSampleBox);
        }

        
        public string FileType(string fileName)
        {
            string text = fileName;
            int lengText = text.Length;
            int vitri = text.IndexOf(".");
            string fileType = text.Substring(vitri, lengText - vitri);
            return fileType;
        }        

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
                        string typeSample = "Box";
                        string modelcode = ApprovalClickItem.custpartcode;
                        byte[] buffer = File.ReadAllBytes(path_File);
                        string base64Encoded = Convert.ToBase64String(buffer);
                        var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                        var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                        string seq = db.Rejetc_MaxSeq(path_sql_attach, "tbSampleAttach", samno, typeSample);
                        string imsempcode = MainWindow.UserLogin;
                        string insdt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                        string updempcode = MainWindow.UserLogin;
                        string upddt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                        Helper_TaixinDB_SampleBox taixinDB = new Helper_TaixinDB_SampleBox();
                        string query = ("INSERT tbSampleAttach(cmpcode,bizdiv,samno,seq,typeSample,modelcode,filename,filedata,qty,imsempcode,insdt,updempcode,upddt) VALUES('02','300','" + samno + "','" + seq + "','" + typeSample + "','" + modelcode + "',N'" + namefile + "','" + base64Encoded + "','" + qty + "','" + imsempcode + "','" + insdt + "','" + updempcode + "','" + upddt + "')");
                        SqlCommand cmd = new SqlCommand(query, conn);
                        int count = (int)cmd.ExecuteNonQuery();
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
                MessageBox.Show(ex.Message, "Box/UploadAttachFile", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }        

        
        public async Task DowLoadAttachFile(string pathFolder)
        {           
            try
            {
                string bufferExe;
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
                    using (SqlConnection conn = new SqlConnection(path_sql_attach))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand("select * from tbSampleAttach where typeSample = 'Box' and samno ='" + ApprovalClickItem.samno + "'", conn))
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
                MessageBox.Show(ex.Message, "Box/DowLoadAttachFile", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }        

        
        public void CreatData()
        {
            CreatAllButtonEdit();            
            ckbTest.Content = "Chuyển đổi";
            accessProcess = "Creat";
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
            ckb_DowloadFile.Visibility = Visibility.Visible;
        }
        
       
        public void ApproveData()
        {           
            lvButtonTop.Items.Clear();          
            if(listButtonTop.Count<6)
            {
                listButtonTop.Add(new Helper_DataButton
                {
                    ID = 6,
                    ContentButton = "Run",
                    ImageSource = "Image/Edit/check.png",
                    BackGroundColor = PinValue.OFF
                });
            }            
            foreach (var button in listButtonTop)
            {                
                if(button.ID>4)
                lvButtonTop.Items.Add(button);
                
            }           
            ckbTest.Content = "Chuyển đổi";
            accessProcess = "Approve";
            stackApprove.Visibility = Visibility.Visible;
            grid_ButtonEditor.Visibility = Visibility.Visible;         
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
            ckb_DowloadFile.Visibility = Visibility.Hidden;

        }

       
        public void DataItemClick(Helper_TaixinDB_SampleBox sample)
        {
            try
            {               
                SampleBox.box.cmpcode = sample.cmpcode;
                SampleBox.box.bizdiv = sample.bizdiv;
                SampleBox.box.samno = sample.samno;
                MainWindow.at_Samno = sample.samno;
                SampleBox.box.modelcode = sample.modelcode;               
                SampleBox.box.modelname = sample.modelname;
                SampleBox.box.applydt = sample.applydt;
                SampleBox.box.version = sample.version;
                SampleBox.box.custpartcode = sample.custpartcode;
                MainWindow.at_ModelCode = sample.custpartcode;
                SampleBox.box.custpart_version = sample.custpart_version;
                SampleBox.box.custpartcode_ver = sample.custpartcode_ver;
                SampleBox.box.custmodelcode = sample.custmodelcode;
                SampleBox.box.repmodelcode = sample.repmodelcode;
                SampleBox.box.useflag = sample.useflag;
                SampleBox.box.modelgroup = sample.modelgroup;
                SampleBox.box.modeldiv = sample.modeldiv;
                SampleBox.box.modeltype = sample.modeltype;
                SampleBox.box.modelchild = sample.modelchild;
                SampleBox.box.cust_gb = sample.cust_gb;
                SampleBox.box.information = sample.information;
                SampleBox.box.paper_name1 = sample.paper_name1;
                SampleBox.box.paper_size1 = sample.paper_size1;
                SampleBox.box.paper_scale1 = sample.paper_scale1;
                SampleBox.box.paper_name2 = sample.paper_name2;
                SampleBox.box.paper_size2 = sample.paper_size2;
                SampleBox.box.paper_scale2 = sample.paper_scale2;
                SampleBox.box.paper_name3 = sample.paper_name3;
                SampleBox.box.paper_size3 = sample.paper_size3;
                SampleBox.box.paper_scale3 = sample.paper_scale3;
                SampleBox.box.cover_up_name1 = sample.cover_up_name1;
                SampleBox.box.cover_up_size1 = sample.cover_up_size1;
                SampleBox.box.cover_up_scale1 = sample.cover_up_scale1;
                SampleBox.box.cover_up_name2 = sample.cover_up_name2;
                SampleBox.box.cover_up_size2 = sample.cover_up_size2;
                SampleBox.box.cover_up_scale2 = sample.cover_up_scale2;
                SampleBox.box.cover_up_name3 = sample.cover_up_name3;
                SampleBox.box.cover_up_size3 = sample.cover_up_size3;
                SampleBox.box.cover_up_scale3 = sample.cover_up_scale3;
                SampleBox.box.cover_up_name4 = sample.cover_up_name4;
                SampleBox.box.cover_up_size4 = sample.cover_up_size4;
                SampleBox.box.cover_up_scale4 = sample.cover_up_scale4;
                SampleBox.box.cover_up_name5 = sample.cover_up_name5;
                SampleBox.box.cover_up_size5 = sample.cover_up_size5;
                SampleBox.box.cover_up_scale5 = sample.cover_up_scale5;
                SampleBox.box.fullsize = sample.fullsize;
                SampleBox.box.print_color1 = sample.print_color1;
                SampleBox.box.print_color2 = sample.print_color2;
                SampleBox.box.print_color3 = sample.print_color3;
                SampleBox.box.print_color4 = sample.print_color4;
                SampleBox.box.print_color5 = sample.print_color5;
                SampleBox.box.coating = sample.coating;
                SampleBox.box.glossy_color = sample.glossy_color;
                SampleBox.box.glossy_detail = sample.glossy_detail;
                SampleBox.box.holo_color1 = sample.holo_color1;
                SampleBox.box.holo_detail1 = sample.holo_detail1;
                SampleBox.box.holo_color2 = sample.holo_color2;
                SampleBox.box.holo_detail2 = sample.holo_detail2;
                SampleBox.box.holo_color3 = sample.holo_color3;
                SampleBox.box.holo_detail3 = sample.holo_detail3;
                SampleBox.box.debosing_detail = sample.debosing_detail;
                SampleBox.box.imbosing_detail = sample.imbosing_detail;
                SampleBox.box.boi_detaiil = sample.boi_detaiil;
                SampleBox.box.boi_color = sample.boi_color;
                SampleBox.box.qty = sample.qty;
                SampleBox.box.forecast = sample.forecast;
                SampleBox.box.remark1 = sample.remark1;
                SampleBox.box.remark2 = sample.remark2;
                SampleBox.box.dep_mar = sample.dep_mar;
                SampleBox.box.dep_pro = sample.dep_pro;
                SampleBox.box.dep_qc = sample.dep_qc;
                SampleBox.box.dep_rnd = sample.dep_rnd;
                SampleBox.box.dep_pur = sample.dep_pur;
                SampleBox.box.paperindex1 = sample.paperindex1;
                SampleBox.box.paperindex2 = sample.paperindex2;
                SampleBox.box.paperindex3 = sample.paperindex3;
                SampleBox.box.printindex1 = sample.printindex1;
                SampleBox.box.printindex2 = sample.printindex2;
                SampleBox.box.printindex3 = sample.printindex3;
                SampleBox.box.reject = sample.reject;
                SampleBox.box.qtyAttach = sample.qtyAttach + " File";
                SampleBox.box.depCreate = sample.depCreate;
                SampleBox.box.printed = sample.printed;
                SampleBox.box.etc2 = sample.etc2;
                SampleBox.box.etc3 = sample.etc3;
                SampleBox.box.holo_color4 = sample.etc4;
                SampleBox.box.holo_detail4 = sample.etc5;
                SampleBox.box.etc6 = sample.etc6;
                SampleBox.box.etc7 = sample.etc7;
                SampleBox.box.etc8 = sample.etc8;
                SampleBox.box.etc9 = sample.etc9;
                SampleBox.box.etc10 = sample.etc10;
                SampleBox.box.sadt = sample.sadt;
                SampleBox.box.fadt = sample.fadt;
                SampleBox.box.imsempcode = sample.imsempcode;
                SampleBox.box.insdt = sample.insdt;
                SampleBox.box.updempcode = sample.updempcode;
                SampleBox.box.upddt = sample.upddt;
                //sx
                if (sample.dep_pro == "LightGray")
                {
                    ckbApprove_SX.IsChecked = false;
                }
                else
                {
                    ckbApprove_SX.IsChecked = true;
                }
                //ma
                if (sample.dep_mar == "LightGray")
                {
                    ckbApprove_Ma.IsChecked = false;
                }
                else
                {
                    ckbApprove_Ma.IsChecked = true;
                }
                //qc
                if (sample.dep_qc == "LightGray")
                {
                    ckbApprove_CL.IsChecked = false;
                }
                else
                {
                    ckbApprove_CL.IsChecked = true;
                }
                //rnd
                if (sample.dep_rnd == "LightGray")
                {
                    ckbApprove_RD.IsChecked = false;
                }
                else
                {
                    ckbApprove_RD.IsChecked = true;
                }
                //kd
                if (sample.dep_pur == "LightGray")
                {
                    ckbApprove_KD.IsChecked = false;
                }
                else
                {
                    ckbApprove_KD.IsChecked = true;
                }
                ApprovalClickItem = sample;
                dpkStartApprove.SelectedDate = DateTime.Parse(sample.sadt);
                dpkFinishApprove.SelectedDate = DateTime.Parse(sample.fadt);               
                Page_RejectSample.samno = sample.samno;
                Page_RejectSample.modelcode = sample.custpartcode;
                Page_RejectSample.typeSample = "Box";
                ApprovalClickItem.samno = sample.samno;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/DataItemClick", MessageBoxButton.OK, MessageBoxImage.Error);
            }            
        }

        
        public void CheckXLS()
        {
            foreach (var item in listSampleBox)
            {
                item.checkXLS = "True";
            }
            ColorRowListView(listSampleBox);
        }
        public void UncheckXLS()
        {
            foreach (var item in listSampleBox)
            {
                item.checkXLS = "False";
            }
            ColorRowListView(listSampleBox);
        }

        private void btnApprove_SX_Click(object sender, RoutedEventArgs e)
        {
            if (str_depCreate=="SX" && ApprovalClickItem.dep_pro == "Red" && (ApprovalClickItem.dep_mar == "DodgerBlue" || ApprovalClickItem.dep_rnd == "DodgerBlue"))
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
            if (str_depCreate == "QC" && ApprovalClickItem.dep_qc == "Red"  && (ApprovalClickItem.dep_mar == "DodgerBlue" || ApprovalClickItem.dep_rnd == "DodgerBlue"))
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
            if (str_depCreate == "RND" && ((ApprovalClickItem.depCreate == "RND" && ApprovalClickItem.dep_rnd == "Red")||(ApprovalClickItem.depCreate == "MARKETING" && ApprovalClickItem.dep_mar == "DodgerBlue")))
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
            if (str_depCreate == "PUR" && ApprovalClickItem.dep_qc == "Red" && (ApprovalClickItem.dep_mar == "DodgerBlue" || ApprovalClickItem.dep_rnd == "DodgerBlue"))
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
            if (str_depCreate == "MARKETING" && ((ApprovalClickItem.depCreate == "MARKETING" && ApprovalClickItem.dep_mar == "Red")||(ApprovalClickItem.depCreate == "RND" && ApprovalClickItem.dep_rnd == "DodgerBlue")))
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

        private void lvApproveSample_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var click = sender as ListView;
            var clickItem = click.SelectedItem as Helper_TaixinDB_SampleBox;
            if (clickItem != null)
            {                
                DataItemClick(clickItem);
            }
        }

        private void dpkStartApprove_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            str_sadt = dpkStartApprove.SelectedDate.ToString();
        }

        private void dpkFinishApprove_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            str_fadt = dpkFinishApprove.SelectedDate.ToString();
        }

        private void cbb_SpecPaper1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            str_paperindex1 = cbb_SpecPaper1.SelectedIndex.ToString();
        }

        private void cbb_SpecPaper2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            str_paperindex2 = cbb_SpecPaper2.SelectedIndex.ToString();
        }

        private void cbb_SpecPaper3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            str_paperindex3 = cbb_SpecPaper3.SelectedIndex.ToString();
        }

        private void cbb_Process_Color1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            str_printindex1 = cbb_Process_Color1.SelectedIndex.ToString();
        }

        private void cbb_Process_Color2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            str_printindex2 = cbb_Process_Color2.SelectedIndex.ToString();
        }

        private void cbb_Process_Color3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            str_printindex3 = cbb_Process_Color3.SelectedIndex.ToString();
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
                        Filter_Sample_NG();
                    }
                    if (clickItem.Content.ToString() == "Approve OK")
                    {
                        str_cbbFilterApprove = "Approve OK";
                        Filter_Sample_OK();
                    }
                    if (clickItem.Content.ToString() == "Approve RE")
                    {
                        str_cbbFilterApprove = "Approve RE";
                        Filter_Sample_RE();
                    }
                    if (clickItem.Content.ToString() == "Tìm kiếm All")
                    {
                        str_cbbFilterApprove = "Tìm kiếm All";
                        Filter_Sample_All();
                    }
                    if (clickItem.Content.ToString() == "Run OK")
                    {
                        str_cbbFilterApprove = "Run OK";
                        Filter_Run_OK();
                    }
                    if (clickItem.Content.ToString() == "Run NG")
                    {
                        str_cbbFilterApprove = "Run NG";
                        Filter_Run_NG();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/cbbFillterApprove_SelectionChanged", MessageBoxButton.YesNo, MessageBoxImage.Warning); throw;
            }            
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
                var command = "SELECT * from tbSampleBox WHERE (custpartcode LIKE '%" + txt_FilterSample.Text + "%'or modelcode LIKE '%" + txt_FilterSample.Text + "%') ORDER by INSDT desc";
                listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
                ColorRowListView(listSampleBox);
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
                        command = "SELECT * from tbSampleBox WHERE (dep_Mar = 'N' or dep_PRO='N' or dep_QC='N' or dep_RnD='N' or dep_Pur='N') AND " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Approve OK":
                    {
                        command = "SELECT * from tbSampleBox WHERE (dep_Mar = 'Y' or dep_Mar='O') and ( dep_PRO='Y' or dep_PRO = 'O') and (dep_QC='Y' or dep_QC ='O') and " +
                            "(dep_RnD='Y' or dep_RnD ='O') and (dep_Pur = 'Y'  or dep_Pur ='O') AND " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Approve RE":
                    {
                        command = "SELECT * from tbSampleBox WHERE reject ='1' AND " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Tìm kiếm All":
                    {
                        command = "SELECT * from tbSampleBox WHERE " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Run OK":
                    {
                        command = "SELECT * from tbSampleBox WHERE " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND printed = 'Y' AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }
                case "Run NG":
                    {
                        command = "SELECT * from tbSampleBox WHERE " +
                            "custpartcode LIKE '%" + txt_FilterSample.Text + "%' AND printed is null AND  INSDT BETWEEN '" + FilterStart + "' AND '" + FilterFinish + "'  ORDER by INSDT desc";
                        break;
                    }

            }
            listSampleBox = db.Read_TaxinDb_SampleBox(path_sql, command);
            ColorRowListView(listSampleBox);

        }
       
        private void btnAttachFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (accessProcess == "Creat" && checkDowload == false)
                {
                    MainWindow.pl_Print = "Box";
                    if (checkDowload == false)
                    {
                        Window_AttachFile attachFile = new Window_AttachFile();
                        attachFile.ShowDialog();
                    }
                    else
                    {
                        //ProcessUploadFile();
                        Process_AttachFile();
                    }
                }
                else if (accessProcess == "Approve" || checkDowload == true)
                {
                    Process_AttachFile();
                }

            }
            catch (Exception)
            {

                throw;
            }
             
        }

        private void btn_Reject_Click(object sender, RoutedEventArgs e)
        {
            Page_RejectSample page_RejectSample = new Page_RejectSample();
            page_RejectSample.Show();
        }

        private void btnApprove_Ma_Click_1(object sender, RoutedEventArgs e)
        {

        }
        
        private void ckb_DowloadFile_Unchecked(object sender, RoutedEventArgs e)
        {
            checkDowload = false;
        }

        private void ckb_DowloadFile_Checked(object sender, RoutedEventArgs e)
        {
            checkDowload = true;
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
    }
    public static class SampleBox
    {
        public static Helper_TaixinDB_SampleBox box = new Helper_TaixinDB_SampleBox();
        public static List<Helper_TaixinDB_SampleBox> listSampleExportExcel = new List<Helper_TaixinDB_SampleBox>();
    }
}
