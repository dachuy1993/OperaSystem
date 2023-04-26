using DataHelper;
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
namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Page_Setting.xaml
    /// </summary>
    public partial class Page_Setting : Page
    {
        List<Helper_DataButton> listButtonTop = new List<Helper_DataButton>();
        string path_sql = MainWindow.path_sql;
        string nameTableApprove = "tbSampleAccess";
        Helper_AccessManger db = new Helper_AccessManger();
        List<Helper_AccessManger> listAccess = new List<Helper_AccessManger>();
        string str_id = "";
        string str_name = "";
        string str_dep = "";
        string str_create = "N";
        string str_approve = "N";
        string str_run = "N";
        string str_add = "N";
        string str_del = "N";
        string str_edit = "N";
        string str_save = "N";
        string str_date = "";        
        string IdNumber = "";
        public Page_Setting()
        {
            InitializeComponent();
            CreatAllButtonEdit();
            Loaded += Page_Setting_Loaded;
        }

        private void Page_Setting_Loaded(object sender, RoutedEventArgs e)
        {
            Read_SettingApprove();
        }
        
        public void CreatAllButtonEdit()
        {
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
            foreach (var button in listButtonTop)
            {
                lvButtonTop.Items.Add(button);
            }

        }

        private void ButtonTop_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var click = sender as Button;
                var clickItem = click.DataContext as Helper_DataButton;
                if (clickItem != null)
                {
                    switch (clickItem.ContentButton)
                    {
                        case "Add":
                            {
                                ProcessButtonEdit_Add();
                                break;
                            }
                        case "Del":
                            {
                                ProcessButtonEdit_Del();
                                break;
                            }
                        case "Edit":
                            {
                                ProcessButtonEdit_Edit();
                                break;
                            }
                        case "Save":
                            {
                                ProcessButtonEdit_Save();
                                break;
                            }
                            //case "Print":
                            //    {
                            //        ProcessButtonEdit_Scan();
                            //        break;
                            //    }
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Page Setting/ButtonTop_Click", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }
       
        public void Read_SettingApprove()
        {
            try
            {
                var command = "SELECT * from " + nameTableApprove + " order by department asc";
                listAccess = db.Read_AccessApprove(path_sql, nameTableApprove, command);
                lvApproveSample.Items.Clear();
                if (listAccess.Count > 0)
                    foreach (var item in listAccess)
                    {
                        if (item.CreatApprove == "Y")
                        {
                            item.CreatApprove = "DodgerBlue";
                        }
                        else
                        {
                            item.CreatApprove = "Red";
                        }
                        //
                        if (item.ApproveApprove == "Y")
                        {
                            item.ApproveApprove = "DodgerBlue";
                        }
                        else
                        {
                            item.ApproveApprove = "Red";
                        }
                        //
                        if (item.ProRun == "Y")
                        {
                            item.ProRun = "DodgerBlue";
                        }
                        else
                        {
                            item.ProRun = "Red";
                        }
                        lvApproveSample.Items.Add(item);
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Page Setting/Read_SettingApprove", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }  

        public void ProcessButtonEdit_Add()
        {
            txt_UserApprove.Text = "";
            txt_DepApprove.Text = "";
            txt_NameApprove.Text = "";
            ck_Creat.IsChecked = false;
            ck_Approve.IsChecked = false;
            ck_Add.IsChecked = false;
            ck_Del.IsChecked = false;
            ck_Edit.IsChecked = false;
            ck_Save.IsChecked = false;           
        }

        public void ProcessButtonEdit_Del()
        {
            db.Delete_AccessApprove(path_sql, nameTableApprove,IdNumber);
            Read_SettingApprove();
            MessageBox.Show("Xóa dữ liệu Thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public void ProcessButtonEdit_Edit()
        {
            str_id = txt_UserApprove.Text.ToUpper();
            str_name = txt_NameApprove.Text;
            str_dep = txt_DepApprove.Text;
            str_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            db.Update_AccessApprove(path_sql,nameTableApprove,IdNumber, str_id, str_name, str_dep, str_create, str_approve,str_run,str_date);
            Read_SettingApprove();
            MessageBox.Show("Cập nhật dữ liệu Thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        public void ProcessButtonEdit_Save()
        {            
            str_id = txt_UserApprove.Text.ToUpper();
            str_name = txt_NameApprove.Text;
            str_dep = txt_DepApprove.Text;
            str_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string maxSamno = db.Read_MaxSamno(path_sql, nameTableApprove);
            db.Insert_AccessApprove(path_sql, nameTableApprove,maxSamno,str_id, str_name,str_dep,str_create, str_approve,str_run,str_date);
            Read_SettingApprove();
            MessageBox.Show("Thêm dữ liệu Thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Dp_DateStart_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Dp_DateFinish_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void txt_FilterSample_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void btn_FilterSample_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var command = "SELECT * from " + nameTableApprove + " where UserLogin like '%" + txt_FilterSample.Text + "%'";
                listAccess = db.Read_AccessApprove(path_sql, nameTableApprove, command);
                lvApproveSample.Items.Clear();
                if (listAccess.Count > 0)
                    foreach (var item in listAccess)
                    {
                        if (item.CreatApprove == "Y")
                        {
                            item.CreatApprove = "DodgerBlue";
                        }
                        else
                        {
                            item.CreatApprove = "Red";
                        }
                        //
                        if (item.ApproveApprove == "Y")
                        {
                            item.ApproveApprove = "DodgerBlue";
                        }
                        else
                        {
                            item.ApproveApprove = "Red";
                        }
                        lvApproveSample.Items.Add(item);

                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Page Setting/btn_FilterSample_Click", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }       

        private void lvApproveSample_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var click = sender as ListView;
                var clickItem = click.SelectedItem as Helper_AccessManger;
                if (clickItem != null)
                {
                    ProcessButtonEdit_Add();
                    txt_UserApprove.Text = clickItem.UserApprove;
                    txt_NameApprove.Text = clickItem.NameApprove;
                    txt_DepApprove.Text = clickItem.DepApprove;
                    IdNumber = clickItem.SamNo.ToString();
                    if (clickItem.CreatApprove == "DodgerBlue")
                    {
                        ck_Creat.IsChecked = true;
                    }
                    if (clickItem.ApproveApprove == "DodgerBlue")
                    {
                        ck_Approve.IsChecked = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Page Setting/lvApproveSample_SelectionChanged", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        private void ck_Creat_Checked(object sender, RoutedEventArgs e)
        {
            str_create = "Y";
        }

        private void ck_Creat_Unchecked(object sender, RoutedEventArgs e)
        {
            str_create = "N";
        }

        private void ck_Approve_Checked(object sender, RoutedEventArgs e)
        {
            str_approve = "Y";
        }

        private void ck_Approve_Unchecked(object sender, RoutedEventArgs e)
        {
            str_approve = "N";

        }

        private void ck_Add_Checked(object sender, RoutedEventArgs e)
        {
            str_add = "Y";
        }

        private void ck_Add_Unchecked(object sender, RoutedEventArgs e)
        {
            str_add = "N";

        }

        private void ck_Del_Checked(object sender, RoutedEventArgs e)
        {
            str_del = "Y";
        }

        private void ck_Del_Unchecked(object sender, RoutedEventArgs e)
        {
            str_del = "N";
        }

        private void ck_Edit_Checked(object sender, RoutedEventArgs e)
        {
            str_edit = "Y";
        }

        private void ck_Edit_Unchecked(object sender, RoutedEventArgs e)
        {
            str_edit = "N";
        }

        private void ck_Save_Checked(object sender, RoutedEventArgs e)
        {
            str_save = "Y";
        }

        private void ck_Save_Unchecked(object sender, RoutedEventArgs e)
        {
            str_save = "N";
        }

        private void ck_Run_Checked(object sender, RoutedEventArgs e)
        {
            str_run = "Y";
        }

        private void ck_Run_Unchecked(object sender, RoutedEventArgs e)
        {
            str_run = "N";
        }
    }

    public class AccessApprove
    {
        public static Helper_AccessManger access = new Helper_AccessManger();
    }
        
}
