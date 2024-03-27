using DataHelper;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
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
using System.Windows.Threading;
using W_Opera.Properties;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Window_AttachFile.xaml
    /// </summary>
    public partial class Window_AttachFile : Window
    {

        string path_sql_attach = "Data Source=192.168.2.10;Initial Catalog=taixin_attach;Persist Security Info=True;User ID=sa;Password= oneuser1!";       
        List<Helper_DataButton> listButtonTop = new List<Helper_DataButton>();
        string processButton = "";
        string _typeSample = "";
        string _modelcode = "";
        string _samno = "";
        int qtyFileUpload = 0;
        public static int qtyAttachFile = 0;
        public static OpenFileDialog ofd_AttachFile;
        FileAttach fileRomove = new FileAttach();
        public static List<FileAttach> listAttachFile = new List<FileAttach>();
        int index = 0;
        int indexHis = 0;
        string check = "";
        public string str_depCreate = "";
        public string str_AddMan = "";
        public string str_DelMan = "";
        public string str_EditMan = "";
        public string str_SaveMan = "";
        public string str_AddBox = "";
        public string str_DelBox = "";
        public string str_EditBox = "";
        public string str_SaveBox = "";
        public Window_AttachFile()
        {
            InitializeComponent();
            Loaded += Window_AttachFile_Loaded;           
        }

        private void Window_AttachFile_Loaded(object sender, RoutedEventArgs e)
        {
            GetDataDeptUser();
            CreatAllButtonEdit();          
            listAttachFile.Clear();
            if (MainWindow.at_Samno != "")
            {
                Db_Read_MaxSeq();
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
                        var command = "SELECT Department, AddMan, DelMan, EditMan, SaveMan, AddBox, DelBox, EditBox, SaveBox FROM tbSampleAccess where UserLogin = '" + MainWindow.UserLogin + "'";
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
                                        str_AddMan = dr[1].ToString();
                                        str_DelMan = dr[2].ToString();
                                        str_EditMan = dr[3].ToString();
                                        str_SaveMan = dr[4].ToString();
                                        str_AddBox = dr[5].ToString();
                                        str_DelBox = dr[6].ToString();
                                        str_EditBox = dr[7].ToString();
                                        str_SaveBox = dr[8].ToString();
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
            //listButtonTop.Add(new Helper_DataButton
            //{
            //    ID = 3,
            //    ContentButton = "Edit",
            //    ImageSource = "Image/Edit/edit.png",
            //    BackGroundColor = PinValue.OFF
            //});
            listButtonTop.Add(new Helper_DataButton
            {
                ID = 4,
                ContentButton = "Save",
                ImageSource = "Image/Edit/save.png",
                BackGroundColor = PinValue.OFF
            });
            //listButtonTop.Add(new Helper_DataButton
            //{
            //    ID = 5,
            //    ContentButton = "Print",
            //    ImageSource = "Image/Edit/printer.png",
            //    BackGroundColor = PinValue.OFF
            //});
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
                    //case "Edit":
                    //    {
                    //        processButton = "Edit";
                    //        ProcessButtonEdit_Edit();
                    //        break;
                    //    }
                    case "Save":
                        {
                            processButton = "Save";
                            ProcessButtonEdit_Save();
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

        private void ProcessButtonEdit_Add()
        {            
            SelectFileAttach();
        }

        private void ProcessButtonEdit_Edit()
        {
            if(MessageBoxResult.Yes == MessageBox.Show("Bạn muốn sửa File đính kèm?","Thông báo",MessageBoxButton.YesNo,MessageBoxImage.Question))
            {
                //Read_DataAttachFile();               
            }    
            
        }

        private void ProcessButtonEdit_Del()
        {
            if(check=="New" && (str_EditMan == "Y" || str_EditBox == "Y" || str_SaveMan == "Y" || str_SaveBox == "Y") )
            {
                if (MessageBoxResult.Yes == MessageBox.Show("Bạn muốn Xóa File đính kèm?\r\nHãy chắc chắn điều này", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question))
                {
                    lvAttachFile.ClearValue(ListView.ItemsSourceProperty);
                    listAttachFile.Remove(fileRomove);
                    lvAttachFile.ItemsSource = listAttachFile;
                }
            } 
            else if(check=="Edit" && (str_EditMan == "Y" || str_EditBox == "Y" || str_SaveMan == "Y" || str_SaveBox == "Y"))
            {
                if (MessageBoxResult.Yes == MessageBox.Show("Bạn muốn xóa file đính kèm?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question))
                {
                    try
                    {
                        if(fileRomove.Samno!=null)
                        {

                            
                            var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                            var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                            //int seq = int.Parse(db.Rejetc_MaxSeq(path_sql_attach, "tbSampleAttach", _samno, _typeSample));
                            //string seqMax = seq.ToString("0000");
                            string imsempcode = MainWindow.UserLogin;
                            string insdt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);

                            using (SqlConnection conn = new SqlConnection(path_sql_attach))
                            {
                                conn.Open();
                                var command = "INSERT INTO tbSampleAttachHistory SELECT *,'D','" + imsempcode + "','" + insdt + "' from tbSampleAttach where samno='" + fileRomove.Samno + "' and seq = '" + fileRomove.Stt + "' and typeSample = '" + fileRomove.Type + "'";
                                using (SqlCommand cmd = new SqlCommand(command, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                    conn.Close();
                                }
                            }


                            using (SqlConnection conn = new SqlConnection(path_sql_attach))
                            {
                                conn.Open();
                                var command = "Delete tbSampleAttach where samno='" + fileRomove.Samno + "' and seq = '" + fileRomove.Stt + "' and typeSample = '" + fileRomove.Type + "'";
                                using (SqlCommand cmd = new SqlCommand(command, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                    conn.Close();
                                    Db_Read_AttachFile();
                                }
                            }
                            using (SqlConnection conn = new SqlConnection(MainWindow.path_sql))
                            {
                                conn.Open();
                                string query = "";
                                if (MainWindow.pl_Print == "Box")
                                {
                                    query = "UPDATE tbSampleBox SET qtyAttach='" + listAttachFile.Count.ToString() + "' where samno='" + fileRomove.Samno + "'";
                                }
                                else if (MainWindow.pl_Print == "Manual")
                                {
                                    query = "UPDATE tbSampleManual SET qtyAttach='" + listAttachFile.Count.ToString() + "' where samno='" + fileRomove.Samno + "'";
                                }
                                using (SqlCommand cmd = new SqlCommand(query, conn))
                                {
                                    cmd.ExecuteNonQuery();
                                    conn.Close();
                                }
                            }
                        }
                        else
                        {
                            listAttachFile.Remove(fileRomove);
                        }
                        lvAttachFile.ClearValue(ListView.ItemsSourceProperty);
                        lvAttachFile.ItemsSource = listAttachFile;
                        MessageBox.Show("Xóa file thành công", "Thông báo", MessageBoxButton.OK,MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Box/UploadAttachFile", MessageBoxButton.OK, MessageBoxImage.Error);
                    }                   
                }
            }     
        }

        private async void ProcessButtonEdit_Save()
        {
            if (MessageBoxResult.Yes == MessageBox.Show("Bạn muốn thêm file đính kèm?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question))
            {
                if (check == "Edit" && (str_EditMan == "Y" || str_EditBox == "Y" || str_SaveMan == "Y" || str_SaveBox == "Y"))
                {
                    foreach (var item in listAttachFile)
                    {
                        if (item.Model == null)
                        {
                            await UploadAttachFile(item.FileName, item.Path, item.Stt);
                        }
                    }
                    using (SqlConnection conn = new SqlConnection(MainWindow.path_sql))
                    {
                        conn.Open();
                        string query = "";
                        if (MainWindow.pl_Print == "Box")
                        {
                            query = "UPDATE tbSampleBox SET qtyAttach='" + listAttachFile.Count.ToString() + "' where samno='" + MainWindow.at_Samno + "'";
                        }
                        else if (MainWindow.pl_Print == "Manual")
                        {
                            query = "UPDATE tbSampleManual SET qtyAttach='" + listAttachFile.Count.ToString() + "' where samno='" + MainWindow.at_Samno + "'";
                        }
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                        listAttachFile.Clear();
                        Db_Read_AttachFile();
                    }
                                  
                }

            }               
        }     

        public void Db_Read_AttachFile()
        {
            try
            {
                lvAttachFile.ClearValue(ListView.ItemsSourceProperty);
                listAttachFile.Clear();
                using (SqlConnection conn = new SqlConnection(path_sql_attach))
                {
                    conn.Open();                    
                    var command = "SELECT samno,seq,modelcode,filename,typeSample FROM tbSampleAttach where samno='" + MainWindow.at_Samno + "'  and typeSample='" + MainWindow.pl_Print + "'";                    
                    using (SqlCommand cmd = new SqlCommand(command, conn))
                    {
                        using (IDataReader dr = cmd.ExecuteReader())
                        {
                            while(dr.Read())
                            {
                                listAttachFile.Add(new FileAttach { Samno=dr[0].ToString(), Stt = dr[1].ToString(), Model = dr[2].ToString(),FileName = dr[3].ToString(),Type =dr[4].ToString() });
                            }    
                        }    
                    }
                    conn.Close();
                }
                index= listAttachFile.Count;
                lvAttachFile.ItemsSource = listAttachFile;              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Helper DataBase/Read_TaxinDb_SampleBox", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }      

        public void Db_Read_MaxSeq()
        {
            using (SqlConnection conn = new SqlConnection(path_sql_attach))
            {
                conn.Open();
                var command = "SELECT max(seq) from tbSampleAttach WHERE typeSample='"+MainWindow.pl_Print+"' and samno='"+MainWindow.at_Samno+"'";
                using (SqlCommand cmd = new SqlCommand(command, conn))
                {
                    if(cmd.ExecuteScalar().ToString()!="")
                    index = int.Parse(cmd.ExecuteScalar().ToString());
                }
                conn.Close();
            }
        }
       
        public void SelectFileAttach()
        {
            try
            {
                        
                lvAttachFile.ClearValue(ListView.ItemsSourceProperty);
                List<FileAttach> listAttachTemp = new List<FileAttach>();
                ofd_AttachFile = new OpenFileDialog();
                ofd_AttachFile.Filter = "All files(*.*)| *.*| Exe Files(*.exe) | *.exe*| Text File(*.txt) |*.txt";
                ofd_AttachFile.FilterIndex = 0;
                ofd_AttachFile.Multiselect = true;
                qtyFileUpload = 0;
                if (ofd_AttachFile.ShowDialog() == true)
                {
                    foreach (var pathFile in ofd_AttachFile.FileNames)
                    {
                        index++;
                        var onlyFileName = System.IO.Path.GetFileName(pathFile);
                        long length = new System.IO.FileInfo(pathFile).Length / 1000;
                        if (length > 80000)
                        {
                            MessageBox.Show("File có dung lượng quá lớn. Vui lòng thay đổi file.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        if(check=="New")
                        {
                            listAttachTemp.Add(new FileAttach { Stt = index.ToString("0000"), FileName = onlyFileName, Path = pathFile,Check="New" });
                        }
                        else
                        {
                            listAttachTemp.Add(new FileAttach { Stt = index.ToString("0000"), FileName = onlyFileName, Path = pathFile,Check="Edit" });
                        }
                    }
                    listAttachFile.AddRange(listAttachTemp);
                    lvAttachFile.ItemsSource = listAttachFile;
                    qtyFileUpload = ofd_AttachFile.FileNames.Count();                    
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn File cần Upload", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Box/SelectFileAttach", MessageBoxButton.OK, MessageBoxImage.Error);
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
       
        public async Task UploadAttachFile(string namefile, string path_File,string stt)
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
                        byte[] buffer = File.ReadAllBytes(path_File);
                        string base64Encoded = Convert.ToBase64String(buffer);
                        var settings = new JsonSerializerSettings { DateFormatString = "yyyy-MM-dd HH:mm:ss" };
                        var jsonDateInput = JsonConvert.SerializeObject(DateTime.Now, settings);
                        //int seq = int.Parse(db.Rejetc_MaxSeq(path_sql_attach, "tbSampleAttach", _samno, _typeSample));
                        //string seqMax = seq.ToString("0000");
                        string imsempcode = MainWindow.UserLogin;
                        string insdt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                        string updempcode = MainWindow.UserLogin;
                        string upddt = jsonDateInput.Substring(1, jsonDateInput.Length - 2);
                        string query = ("INSERT tbSampleAttach(cmpcode,bizdiv,samno,seq,typeSample,modelcode,filename,filedata,qty,imsempcode,insdt,updempcode,upddt) VALUES('02','300','" + MainWindow.at_Samno + "','" + stt + "','" + MainWindow.pl_Print + "','" + MainWindow.at_ModelCode + "',N'" + namefile + "','" + base64Encoded + "','" + stt + "','" + imsempcode + "','" + insdt + "','" + updempcode + "','" + upddt + "')");
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }

                        string queryHis = ("INSERT tbSampleAttachHistory(cmpcode,bizdiv,samno,seq,typeSample,modelcode,filename,filedata,qty,imsempcode,insdt,updempcode,upddt,typeAUD) VALUES('02','300','" + MainWindow.at_Samno + "','" + stt + "','" + MainWindow.pl_Print + "','" + MainWindow.at_ModelCode + "',N'" + namefile + "','" + base64Encoded + "','" + stt + "','" + imsempcode + "','" + insdt + "','" + updempcode + "','" + upddt + "','A')");
                        using (SqlCommand cmd = new SqlCommand(queryHis, conn))
                        {
                            cmd.ExecuteNonQuery();
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
                MessageBox.Show(ex.Message, "Window_AttachFile/UploadAttachFile", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void lvAttachFile_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var click = sender as ListView;
            var clickItem = click.SelectedItem as FileAttach;
            if(clickItem!=null)
            {
                fileRomove = clickItem;
            }    
        }
       
        private void rb_NewAttach_Checked(object sender, RoutedEventArgs e)
        {
            check = "New";
        }

        private void rb_EditAttach_Checked(object sender, RoutedEventArgs e)
        {
            check = "Edit";
            Db_Read_AttachFile();
        }
    }
    public class FileAttach
    {
        public string Samno { get; set; }
        public string Model { get; set; }
        public string Stt { get; set; }
        public string Path { get; set; }
        public string Type { get; set; }
        public string FileName { get; set; }
        public string Check { get; set; }
    }
}
