using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
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
using System.Windows.Xps.Packaging;
using W_Opera.DAO;
using ZXing;
using ZXing.QrCode;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Print.xaml
    /// </summary>
    public partial class Print : Window
    {
        string DateNow;

        public Print()
        {
            InitializeComponent();
            Loaded += Print_Loaded;
        }

        private void Print_Loaded(object sender, RoutedEventArgs e)
        {
           
            if (!Directory.Exists(@"TempFile"))
                Directory.CreateDirectory(@"TempFile");           
            TimeRun();           
        }


        DispatcherTimer dt = new DispatcherTimer();
        public void TimeRun()
        {
            dt.Interval = TimeSpan.FromMilliseconds(1);
            dt.Tick += Dt_Tick;
            dt.Start();
        }

        private void Dt_Tick(object sender, EventArgs e)
        {
            if(MainWindow.checkPrint==true)
            {
                List<string> list = new List<string>();
                string query1 = "SELECT RIGHT(LEFT(CONVERT(VARCHAR(8),getdate(),112),6),2)+'/'+RIGHT(CONVERT(VARCHAR(8),getdate(),112),2) + '/' + LEFT(CONVERT(VARCHAR(8),getdate(),112),4)";
                string Pra1 = "122";
                //list = db.Read_TaxinDb_SampleBox(MainWindow.path_sql, query1);
                var DateNowList = DataProvider.Instance.executeQuery(MainWindow.path_sql, query1, new object[] { Pra1 });
                foreach (DataRow rowA in DateNowList.Rows)
                {
                    DateNow = rowA[0].ToString();
                }

               

                CreatFileExcel("Box","", "", "", "");                 
                ViewExcelFile(pathFileExcel);               
            }
        }

        string pathFileImage1 = @"TempFile//QrCode1.png";
        string pathFileImage2 = @"TempFile//QrCode2.png";
        string pathFileImage3 = @"TempFile//QrCode3.png";
        string pathFileImage4 = @"TempFile//QrCode4.png";
        string pathFileExcel  = @"TempFile//ExcelFile.xlsx";
        bool exportFileExcel  = false;       

        public void WriteBarcode(string strCode, string pathFileSave)
        {
            try
            {
                IBarcodeWriter writer = new ZXing.BarcodeWriter();
                QrCodeEncodingOptions options = new QrCodeEncodingOptions();
                options = new QrCodeEncodingOptions
                {
                    DisableECI = true,
                    CharacterSet = "UTF-8",
                    Width = 75,
                    Height = 75,
                };
                var qr = new ZXing.BarcodeWriter();
                qr.Options = options;
                qr.Format = ZXing.BarcodeFormat.QR_CODE;
                if (strCode.Length > 0)
                {
                    //string pathSaveImage = @"D:\Drive\3.Visual Studio\9.WPF\25.Excel_QrCode\Excel_QrCode_Master\Excel_QrCode_Master\bin\Debug\";
                    //SaveFileDialog sfd = new SaveFileDialog();
                    //sfd.Filter = " Image file (*PNG)|*.png|(*JPG)|*.jpg|(*JPEG)|*.jpeg|(*GIF)|*.gif|All file(*.*)|*.*";
                    //sfd.ShowDialog();
                    //pathSaveImage = sfd.FileName;                  
                    var result = new System.Drawing.Bitmap(qr.Write(strCode));
                    result.Save(pathFileSave);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Printed(string samno,string typeSample,string imsemplecode)
        {         
            

                try
            {
                using (SqlConnection conn = new SqlConnection(MainWindow.path_sql))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("Insert tbSamplePrint(samno,typeSample,imsempcode) values('"+samno+"','"+typeSample+"','"+imsemplecode+"')", conn))
                    {
                        cmd.ExecuteNonQuery();                        
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Print/Printed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        public void CreatFileExcel(string name,string _LotNo, string _BopNo, string _TaixinNo, string _CusNo)
        {

            switch(MainWindow.pl_Print)
            {
                case "input":
                    {
                        _LotNo = MainWindow.print_DB.LotNo;
                        _BopNo = MainWindow.print_DB.WorkedNo;
                        _TaixinNo = MainWindow.print_DB.ModelCodel;
                        _CusNo = MainWindow.print_DB.CustomerCode;
                        WriteBarcode(_LotNo, pathFileImage1);
                        WriteBarcode(_BopNo, pathFileImage2);
                        WriteBarcode(_TaixinNo, pathFileImage3);
                        WriteBarcode(_CusNo, pathFileImage4);

                        //txt_From_Name.Text = db_Input.PositionFromName.ToString();
                        //txt_To_Name.Text = db_Input.PositionToName.ToString();
                        //txb_Hour.Text = db_Input.HourInput.ToString();
                        //txb_Minute.Text = db_Input.MinuteInput.ToString();
                        //txt_LotNo.Text = db_Input.LotNo.ToString();
                        //txt_WorkNo.Text = db_Input.WorkedNo.ToString();
                        //txt_CustomerCode.Text = db_Input.CustomerCode.ToString();
                        //txt_ModelCode.Text = db_Input.ModelCodel.ToString();
                        ////txt_ModelName.Text = db_Input.TMSTMODEL_ModelName.ToString();
                        //txt_QtyInput.Text = db_Input.QtyInput.ToString();
                        //txbInput_UserInput.Text = db_Input.UserInput.ToString();
                        //txt_Note.Text = db_Input.Note;
                        //dp_DateInput.SelectedDate = DateTime.Parse(db_Input.AllTimeInput.ToString());
                        //strIDnumber = db_Input.InputNo;
                        if (exportFileExcel == true)
                        {
                            // tạo SaveFileDialog để lưu file excel
                            SaveFileDialog dialog = new SaveFileDialog();

                            // chỉ lọc ra các file có định dạng Excel
                            dialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";

                            // Nếu mở file và chọn nơi lưu file thành công sẽ lưu đường dẫn lại dùng
                            if (dialog.ShowDialog() == true)
                            {
                                pathFileExcel = dialog.FileName;
                            }

                            // nếu đường dẫn null hoặc rỗng thì báo không hợp lệ và return hàm
                            if (string.IsNullOrEmpty(pathFileExcel))
                            {
                                MessageBox.Show("Đường dẫn báo cáo không hợp lệ");
                                return;
                            }
                        }

                        try
                        {
                            using (ExcelPackage p = new ExcelPackage())
                            {
                                p.Workbook.Properties.Author = "Hoang Minh";
                                p.Workbook.Properties.Title = "ID TAG";
                                p.Workbook.Worksheets.Add("Sheet1");
                                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                                ws.Name = "Sheet1";
                                ws.Cells.Style.Font.Size = 11;

                                for (int i = 1; i <= 6; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    ws.Column(i).Width = 12;
                                }

                                for (int i = 1; i <= 26; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "F" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var border = cell.Style.Border;
                                    border.Bottom.Style =
                                        border.Top.Style =
                                        border.Left.Style =
                                        border.Right.Style = ExcelBorderStyle.Thin;
                                    cell.Style.WrapText = true;
                                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                for (int i = 4; i <= 12; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "F" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    cell.Style.Font.Size = 16;
                                    cell.Style.Font.Bold = true;
                                }

                                for (int i = 3; i < 26; i++)
                                {

                                    string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Size = 9;
                                    cell.Style.Font.Bold = true;
                                }

                                for (int i = 3; i < 26; i++)
                                {
                                    if (i == 6 || i == 7 || i == 10 || i == 11)
                                    {
                                        string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                        var cell = ws.Cells[strCell];
                                        var fill = cell.Style.Fill;
                                        fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                                        cell.Style.Font.Size = 9;
                                        cell.Style.Font.Bold = true;
                                    }
                                    if (i == 6 || i == 7 || i == 10 || i == 11)
                                    {
                                        string strCell = "F" + i.ToString() + ":" + "F" + i.ToString();
                                        var cell = ws.Cells[strCell];
                                        var fill = cell.Style.Fill;
                                        fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                        cell.Style.Font.Size = 16;
                                        cell.Style.Font.Bold = true;
                                    }

                                }


                                for (int i = 1; i <= 26; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    if (i <= 12 && i > 4)
                                    {
                                        ws.Row(i).Height = 30;
                                        cell.Style.Font.Size = 16;
                                    }
                                    else if (i > 24)
                                    {
                                        ws.Row(i).Height = 30;
                                    }
                                    else
                                    {
                                        ws.Row(i).Height = 17;
                                    }
                                }


                                ws.Cells["A1:A1"].Value = "▲";
                                ws.Cells["A1:F1"].Merge = true;
                                ws.Cells["A1:F1"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["A1:F1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells["A1:A1"].Style.Font.Size = 30;
                                ws.Cells["A1:A1"].Style.Font.Bold = true;
                                ws.Row(1).Height = 35;


                                ws.Cells["A2:A2"].Value = "ID TAG";
                                ws.Cells["A2:F2"].Merge = true;
                                ws.Cells["A2:F2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["A2:F2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                ws.Cells["A2:A2"].Style.Font.Size = 30;
                                ws.Cells["A2:A2"].Style.Font.Bold = true;
                                ws.Row(2).Height = 35;

                                //Ngày SX
                                ws.Cells["A3:A3"].Value = "Ngày SX";
                                ws.Cells["A3:A4"].Merge = true;

                                ws.Cells["B3:B3"].Value = MainWindow.print_DB.DateInput;
                                ws.Cells["B3:C4"].Merge = true;

                                ws.Cells["D3:D3"].Value = "Assy Manual";
                                ws.Cells["D3:F4"].Merge = true;

                                //LOT NO

                                ws.Cells["A5:A5"].Value = "LOTNO";
                                ws.Cells["A5:A6"].Merge = true;

                                ws.Cells["B5:B5"].Value = _LotNo;
                                ws.Cells["B5:B5"].Style.Font.Size = 22;
                                //ws.Cells["B5:B5"].Style.Font.SetFromFont(new System.Drawing.Font("Segoe UI Black", 18));             
                                ws.Cells["B5:E6"].Merge = true;

                                ws.Cells["F5:F5"].Value = "";
                                ws.Cells["F5:F6"].Merge = true;

                                //BOP

                                ws.Cells["F7:F7"].Value = "BOP";
                                ws.Cells["A7:A8"].Merge = true;

                                ws.Cells["B7:D7"].Value = _BopNo;
                                ws.Cells["B7:D7"].Style.Font.Size = 22;
                                ws.Cells["B7:E8"].Merge = true;
                                ws.Cells["F7:F8"].Merge = true;

                                //P/NO

                                ws.Cells["A9:A9"].Value = "P/NO";
                                ws.Cells["A9:A10"].Merge = true;

                                ws.Cells["B9:B9"].Value = _TaixinNo;
                                ws.Cells["B9:B9"].Style.Font.Size = 22;
                                ws.Cells["B9:E10"].Merge = true;

                                ws.Cells["F9:F9"].Value = "";
                                ws.Cells["F9:F10"].Merge = true;

                                //CuspartCode

                                ws.Cells["F11:F11"].Value = "CUST PART CODE";
                                ws.Cells["A11:A12"].Merge = true;

                                ws.Cells["B11:B11"].Value = _CusNo;
                                ws.Cells["B11:B11"].Style.Font.Size = 22;
                                ws.Cells["B11:E12"].Merge = true;
                                ws.Cells["F11:F12"].Merge = true;

                                //CuspartCode

                                ws.Cells["A13:A13"].Value = "WONO";
                                ws.Cells["A13:A14"].Merge = true;

                                ws.Cells["B13:B13"].Value = "12";
                                ws.Cells["B13:E14"].Merge = true;

                                ws.Cells["F13:F13"].Value = "TA";
                                ws.Cells["F13:F14"].Merge = true;
                                ws.Cells["F13:F14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                ws.Cells["F13:F14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                //QTy

                                ws.Cells["A15:A15"].Value = "Số Lượng";
                                ws.Cells["A15:A16"].Merge = true;

                                ws.Cells["B15:B15"].Value = MainWindow.print_DB.QtyInput;
                                ws.Cells["B15:C16"].Merge = true;
                                ws.Cells["B15:B15"].Style.Font.Size = 20;
                                ws.Cells["B15:B15"].Style.Font.Bold = true;

                                ws.Cells["D15:D15"].Value = "PLT";
                                ws.Cells["D15:D15"].Style.Font.Bold = true;
                                ws.Cells["D16:D16"].Value = MainWindow.print_DB.UnitInput;
                                ws.Cells["D16:D16"].Style.Font.Bold = true;

                                ws.Cells["E15:E15"].Value = "";
                                ws.Cells["E15:E16"].Merge = true;

                                ws.Cells["F15:F15"].Value = "";
                                ws.Cells["F15:F16"].Merge = true;

                                //QTy

                                ws.Cells["A17:A17"].Value = "TIẾP THEO";
                                ws.Cells["A17:A18"].Merge = true;

                                ws.Cells["B17:B17"].Value = "";
                                ws.Cells["B17:B18"].Merge = true;

                                ws.Cells["C17:C17"].Value = "PACK THỨ";
                                ws.Cells["C17:C18"].Merge = true;

                                ws.Cells["D17:D17"].Value = "1 C";
                                ws.Cells["D17:D18"].Merge = true;

                                ws.Cells["E17:E17"].Value = "BÌNH THƯỜNG";
                                ws.Cells["E17:F18"].Merge = true;

                                //CÔNG ĐOẠN

                                ws.Cells["A19:A19"].Value = "CÔNG ĐOẠN";
                                ws.Cells["A19:A20"].Merge = true;

                                ws.Cells["B19:B19"].Value = MainWindow.print_DB.PositionChangeDate;
                                ws.Cells["B19:F20"].Merge = true;

                                //CÔNG NHÂN

                                ws.Cells["A21:A21"].Value = "CÔNG NHÂN";
                                ws.Cells["A21:A22"].Merge = true;

                                ws.Cells["B21:B21"].Value = "L07099";
                                ws.Cells["B21:C22"].Merge = true;

                                ws.Cells["D21:D21"].Value = "TAG TEMP";
                                ws.Cells["D21:D22"].Merge = true;

                                ws.Cells["E21:E21"].Value = "ĐẠT";
                                ws.Cells["E21:F22"].Merge = true;

                                //GHI CHÚ

                                ws.Cells["A23:A23"].Value = "GHI CHÚ";
                                ws.Cells["A23:A24"].Merge = true;

                                ws.Cells["B23:B23"].Value = "GHI CHÚ NỘI DUNG";
                                ws.Cells["B23:D24"].Merge = true;

                                ws.Cells["E23:E23"].Value = "NONE FSC";
                                ws.Cells["E23:F24"].Merge = true;

                                //CONG ĐOẠN SAU

                                ws.Cells["A25:A25"].Value = "CÔNG ĐOẠN SAU";
                                ws.Cells["A25:A26"].Merge = true;

                                ws.Cells["B25:B25"].Value = "PHỤ TRÁCH";
                                ws.Cells["B25:D25"].Merge = true;

                                ws.Cells["B26:B26"].Value = "NGÀY";
                                ws.Cells["B26:D26"].Merge = true;

                                ws.Cells["E25:E25"].Value = "XÁC NHẬN";
                                ws.Cells["E25:F26"].Merge = true;
                                //Taixin

                                ws.Cells["A27:A27"].Value = "Page 1 of 1";
                                ws.Cells["A27:B27"].Merge = true;
                                ws.Cells["A27:B27"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["A27:B27"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                                ws.Cells["C27:C27"].Value = "Taixin Printing Vina";
                                ws.Cells["C27:D27"].Merge = true;
                                ws.Cells["C27:D27"].Style.Font.Bold = true;
                                ws.Cells["C27:D27"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["C27:D27"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                //ws.Cells["E27:E27"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                ws.Cells["E27:E27"].Value = DateNow;
                                ws.Cells["E27:F27"].Merge = true;
                                ws.Cells["E27:F27"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["E27:F27"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                ws.PrinterSettings.PaperSize = ePaperSize.A5;
                                ws.PrinterSettings.Orientation = eOrientation.Portrait;
                                ws.PrinterSettings.FitToPage = true;
                                //ws.PrinterSettings.FitToWidth = 0;
                                //ws.PrinterSettings.FitToHeight = 0;
                                ws.PrinterSettings.TopMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.LeftMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.BottomMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.RightMargin = Decimal.Parse("0.25");
                                //ws.PrinterSettings.HorizontalCentered = true;
                                //ws.PrinterSettings.VerticalCentered = true;

                                FileInfo fileImage1 = new FileInfo(pathFileImage1);
                                FileInfo fileImage2 = new FileInfo(pathFileImage2);
                                FileInfo fileImage3 = new FileInfo(pathFileImage3);
                                FileInfo fileImage4 = new FileInfo(pathFileImage4);
                                if (fileImage1.Exists)
                                {
                                    var image1 = ws.Drawings.AddPicture("image1", fileImage1);
                                    image1.SetPosition(4, 2, 5, 2);
                                }
                                if (fileImage2.Exists)
                                {
                                    var image2 = ws.Drawings.AddPicture("image2", fileImage2);
                                    image2.SetPosition(6, 2, 0, 2);
                                }
                                if (fileImage3.Exists)
                                {
                                    var image3 = ws.Drawings.AddPicture("image3", fileImage3);
                                    image3.SetPosition(8, 2, 5, 2);
                                }
                                if (fileImage4.Exists)
                                {
                                    var image4 = ws.Drawings.AddPicture("image4", fileImage4);
                                    image4.SetPosition(10, 2, 0, 2);
                                }

                                File.Delete(pathFileExcel);
                                Byte[] bin = p.GetAsByteArray();
                                File.WriteAllBytes(pathFileExcel, bin);
                                if (fileImage1.Exists)
                                {
                                    File.Delete(pathFileImage1);
                                }
                                if (fileImage2.Exists)
                                {
                                    File.Delete(pathFileImage2);
                                }
                                if (fileImage3.Exists)
                                {
                                    File.Delete(pathFileImage3);
                                }
                                if (fileImage4.Exists)
                                {
                                    File.Delete(pathFileImage4);
                                }
                                exportFileExcel = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Creat Excel File", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        break;
                    }
                case "Manual":
                    {                        
                        if (exportFileExcel == true)
                        {
                            // tạo SaveFileDialog để lưu file excel
                            SaveFileDialog dialog = new SaveFileDialog();

                            // chỉ lọc ra các file có định dạng Excel
                            dialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";

                            // Nếu mở file và chọn nơi lưu file thành công sẽ lưu đường dẫn lại dùng
                            if (dialog.ShowDialog() == true)
                            {
                                pathFileExcel = dialog.FileName;
                            }

                            // nếu đường dẫn null hoặc rỗng thì báo không hợp lệ và return hàm
                            if (string.IsNullOrEmpty(pathFileExcel))
                            {
                                MessageBox.Show("Đường dẫn báo cáo không hợp lệ");
                                return;
                            }
                        }
                        try
                        {
                            using (ExcelPackage p = new ExcelPackage())
                            {
                                //p.Workbook.Properties.Author = DateTime.Now.ToShortDateString();
                                p.Workbook.Properties.Author = DateNow;
                                p.Workbook.Properties.Title = "Sample";
                                p.Workbook.Worksheets.Add("Sheet1");
                                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                                ws.Name = "Sheet1";
                                //ws.Cells.Style.Font.Size = 13;

                                //Setting font cho print
                                //for (int i = 1; i <= 100; i++)
                                //{
                                //    string strCell = "A" + i.ToString() + ":" + "Z" + i.ToString();
                                //    var cell = ws.Cells[strCell].Style.Font;
                                //    cell.SetFromFont(new Font("Times New Roman", 12));                                                               
                                //}                               

                                for (int i = 1; i <= 32; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    ws.Cells.Style.Font.Size = 8;
                                    ws.Column(i).Width = 3.7;
                                    ws.Row(i).Height = 21;                                   
                                }
                                //căn hàng và cột cho tất cả các ô
                                for (int i = 1; i < 32; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "X" + i.ToString();
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

                                //Bôi den backgroud
                                for (int i = 5; i < 7; i++)
                                {

                                    string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);                                   
                                    cell.Style.Font.Bold = true;
                                }
                                for (int i = 5; i < 7; i++)
                                {

                                    string strCell = "I" + i.ToString() + ":" + "I" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Bold = true;
                                }
                                for (int i = 5; i < 7; i++)
                                {

                                    string strCell = "Q" + i.ToString() + ":" + "Q" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Bold = true;
                                }
                                //Bôi đen cột số 2 phần Qui trình
                                for (int i = 8; i < 32; i++)
                                {

                                    string strCell = "D" + i.ToString() + ":" + "K" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Bold = true;
                                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                }
                                
                                //can le ve ben trai 

                                for (int i = 8; i < 32; i++)
                                {

                                    string strCell = "E" + i.ToString() + ":" + "X" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;                                    
                                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                }                                

                                ws.Cells["A1:A1"].Value = "▲";
                                ws.Cells["A1:X1"].Merge = true;                              
                                ws.Cells["A1:A1"].Style.Font.Size = 25;
                                ws.Cells["A1:A1"].Style.Font.Bold = true;
                                ws.Row(1).Height = 40;


                                ws.Cells["A2:A2"].Value = "LỆNH SẢN XUẤT MẪU";
                                ws.Cells["A2:X2"].Merge = true;                               
                                ws.Cells["A2:A2"].Style.Font.Size = 22;
                                ws.Cells["A2:A2"].Style.Font.Bold = true;
                                ws.Row(2).Height = 40;

                                //Ngày SX
                                ws.Cells["A3:B3"].Value = "Ngày : ";
                                ws.Cells["A3:B3"].Merge = true;
                                ws.Cells["A3:B3"].Style.Font.Size = 12;

                                //ws.Cells["C3:X3"].Value = DateTime.Now.ToString("dd/MM/yyyy");
                                ws.Cells["C3:X3"].Value = DateNow;
                                ws.Cells["C3:X3"].Merge = true;
                                ws.Cells["C3:X3"].Style.Font.Size = 12;
                                ws.Cells["C3:X3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                               

                                //CÁCH                            

                                ws.Cells["A4:X4"].Merge = true;                              
                                ws.Row(4).Height = 10;

                                //MODEL
                                ws.Cells["A5:B5"].Value = "Model";
                                ws.Cells["A5:B5"].Merge = true;                               
                                ws.Cells["A5:B5"].Style.Font.Bold = true;
                                ws.Cells["C5:H5"].Value = SamplePaper.PaperOut.CUSTMODELCODE;
                                ws.Cells["C5:H5"].Merge = true;                        

                                //VER
                                ws.Cells["A6:B6"].Value = "Ver";
                                ws.Cells["A6:B6"].Merge = true;                               
                                ws.Cells["A6:B6"].Style.Font.Bold = true;
                                ws.Cells["C6:H6"].Value = SamplePaper.PaperOut.VERSION;
                                ws.Cells["C6:H6"].Merge = true;

                                //Mã hàng
                                ws.Cells["I5:J5"].Value = "Mã Hàng";
                                ws.Cells["I5:J5"].Merge = true;
                                ws.Cells["I5:J5"].Style.Font.Bold = true;
                                ws.Cells["K5:P5"].Value = SamplePaper.PaperOut.CUSTPARTCODE;
                                ws.Cells["K5:P5"].Merge = true;

                                //Khách hàng
                                ws.Cells["I6:J6"].Value = "Khách Hàng";
                                ws.Cells["I6:J6"].Merge = true;
                                ws.Cells["I6:J6"].Style.Font.Bold = true;
                                ws.Cells["K6:P6"].Value = SamplePaper.PaperOut.CUST_GB;
                                ws.Cells["K6:P6"].Merge = true;

                                //Tình trạng
                                ws.Cells["Q5:R5"].Value = "Tình Trạng";
                                ws.Cells["Q5:R5"].Merge = true;
                                ws.Cells["Q5:R5"].Style.Font.Bold = true;
                                ws.Cells["S5:X5"].Value = SamplePaper.PaperOut.VERSIONUP;
                                ws.Cells["S5:X5"].Merge = true;

                                //Yêu cầu
                                ws.Cells["Q6:R6"].Value = "Yêu cầu";
                                ws.Cells["Q6:R6"].Merge = true;
                                ws.Cells["Q6:R6"].Style.Font.Bold = true;
                                ws.Cells["S6:X6"].Value = SamplePaper.PaperOut.INSEMPCODE;
                                ws.Cells["S6:X6"].Merge = true;

                                //Khoảng trống
                                ws.Cells["A7:X7"].Merge = true;
                                ws.Row(7).Height = 10;

                                //Nguyên liệu
                                ws.Cells["A8:C13"].Value = "SPEC";
                                ws.Cells["A8:C13"].Merge = true;
                                ws.Cells["A8:C13"].Style.Font.Bold = true;
                                ws.Cells["A8:C13"].Style.Font.Size = 16;
                                ws.Cells["A8:C13"].Style.WrapText = true;

                                //GL
                                ws.Cells["D8:K8"].Value = "Giấy in bìa";
                                ws.Cells["D8:K8"].Merge = true;
                                ws.Cells["D8:K8"].Style.Font.Bold = true;
                                //GL
                                ws.Cells["D9:K9"].Value = "Kích thước in";
                                ws.Cells["D9:K9"].Merge = true;
                                ws.Cells["D9:K9"].Style.Font.Bold = true;
                                //GL
                                ws.Cells["D10:K10"].Value = "Giấy in ruột";
                                ws.Cells["D10:K10"].Merge = true;
                                ws.Cells["D10:K10"].Style.Font.Bold = true;
                                //GL
                                ws.Cells["D11:K11"].Value = "Kích thước in";
                                ws.Cells["D11:K11"].Merge = true;
                                ws.Cells["D11:K11"].Style.Font.Bold = true;
                                //GL
                                ws.Cells["D12:K12"].Value = "Kích thước thành phẩm (Rộng/Cao)mm";
                                ws.Cells["D12:K12"].Merge = true;
                                ws.Cells["D12:K12"].Style.Font.Bold = true;
                                //GL
                                ws.Cells["D13:K13"].Value = "Số trang";
                                ws.Cells["D13:K13"].Merge = true;
                                ws.Cells["D13:K13"].Style.Font.Bold = true;

                                //Qui Trình
                                ws.Cells["A14:C22"].Value = "Qui Trình";
                                ws.Cells["A14:C22"].Merge = true;
                                ws.Cells["A14:C22"].Style.Font.Bold = true;
                                ws.Cells["A14:C22"].Style.Font.Size = 16;
                                ws.Cells["A14:C22"].Style.WrapText = true;

                                //QT
                                ws.Cells["D14:K14"].Value = "Công đoạn";
                                ws.Cells["D14:K14"].Merge = true;
                                ws.Cells["D14:K14"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["D15:G16"].Value = "IN";
                                ws.Cells["D15:G16"].Merge = true;
                                ws.Cells["D15:G16"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["H15:K15"].Value = "Màu Bìa";
                                ws.Cells["H15:K15"].Merge = true;
                                ws.Cells["H15:K15"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["H16:K16"].Value = "Màu Ruột";
                                ws.Cells["H16:K16"].Merge = true;
                                ws.Cells["H16:K16"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["D17:K17"].Value = "Coating(Lớp phủ)";
                                ws.Cells["D17:K17"].Merge = true;
                                ws.Cells["D17:K17"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["D18:K18"].Value = "Cắt sau in";
                                ws.Cells["D18:K18"].Merge = true;
                                ws.Cells["D18:K18"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["D19:K19"].Value = "Gấp";
                                ws.Cells["D19:K19"].Merge = true;
                                ws.Cells["D19:K19"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["D20:K20"].Value = "Keo";
                                ws.Cells["D20:K20"].Merge = true;
                                ws.Cells["D20:K20"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["D21:K21"].Value = "Ghim";
                                ws.Cells["D21:K21"].Merge = true;
                                ws.Cells["D21:K21"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["D22:K22"].Value = "Cắt TP";
                                ws.Cells["D22:K22"].Merge = true;
                                ws.Cells["D22:K22"].Style.Font.Bold = true;

                                //Ghi Chú
                                ws.Cells["A23:C26"].Value = "Ghi Chú";
                                ws.Cells["A23:C26"].Merge = true;
                                ws.Cells["A23:C26"].Style.Font.Bold = true;
                                ws.Cells["A23:C26"].Style.Font.Size = 16;
                                ws.Cells["A23:C26"].Style.WrapText = true;

                                //GC
                                ws.Cells["D23:G23"].Value = "Ngày Yêu Cầu";
                                ws.Cells["D23:G23"].Merge = true;
                                ws.Cells["D23:G23"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["H23:K23"].Value = "Ngày Hoàn Thành";
                                ws.Cells["H23:K23"].Merge = true;
                                ws.Cells["H23:K23"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["D24:K24"].Value = "Số lượng yêu cầu";
                                ws.Cells["D24:K24"].Merge = true;
                                ws.Cells["D24:K24"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["D25:K25"].Value = "Những Thông tin khác";
                                ws.Cells["D25:K25"].Merge = true;
                                ws.Cells["D25:K25"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["D26:K26"].Value = "Ghi chú";
                                ws.Cells["D26:K26"].Merge = true;
                                ws.Cells["D26:K26"].Style.Font.Bold = true;



                                //input giay bia
                                ws.Cells["L8:X8"].Value = SamplePaper.PaperOut.PAPERNAMEOut;
                                ws.Cells["L8:X8"].Merge = true;
                                ws.Cells["L8:X8"].Style.Font.Bold = true;

                                //input GIẤY BÌA
                                ws.Cells["L9:O9"].Value = SamplePaper.PaperOut.HEIGHTOut;
                                ws.Cells["L9:O9"].Merge = true;
                                ws.Cells["L9:O9"].Style.Font.Bold = true;
                                //TY LE
                                ws.Cells["P9:Q9"].Value = "Tỷ lệ";
                                ws.Cells["P9:Q9"].Merge = true;
                                ws.Cells["P9:Q9"].Style.Font.Bold = true;
                                ws.Cells["P9:Q9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input TY LỆ
                                ws.Cells["R9:S9"].Value = SamplePaper.PaperOut.PHCOUNTOut;
                                ws.Cells["R9:S9"].Merge = true;
                                ws.Cells["R9:S9"].Style.Font.Bold = true;
                                //input NVL
                                ws.Cells["T9:U9"].Value = "Mã NVL";
                                ws.Cells["T9:U9"].Merge = true;
                                ws.Cells["T9:U9"].Style.Font.Bold = true;
                                ws.Cells["T9:U9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input NVL
                                ws.Cells["V9:x9"].Value = SamplePaper.PaperOut.PAPERNAME_FullOut;
                                ws.Cells["V9:x9"].Merge = true;
                                ws.Cells["V9:x9"].Style.Font.Bold = true;

                                //ruột
                                ws.Cells["L10:X10"].Value = SamplePaper.PaperIn.PAPERNAMEIn;
                                ws.Cells["L10:X10"].Merge = true;
                                ws.Cells["L10:X10"].Style.Font.Bold = true;
                                //input ruột
                                ws.Cells["L11:O11"].Value = SamplePaper.PaperIn.HEIGHTIn;
                                ws.Cells["L11:O11"].Merge = true;
                                ws.Cells["L11:O11"].Style.Font.Bold = true;
                                //Ty le
                                ws.Cells["P11:Q11"].Value = "Tỷ lệ";
                                ws.Cells["P11:Q11"].Merge = true;
                                ws.Cells["P11:Q11"].Style.Font.Bold = true;
                                ws.Cells["P11:Q11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input ty le
                                ws.Cells["R11:S11"].Value = SamplePaper.PaperIn.PHCOUNTIn;
                                ws.Cells["R11:S11"].Merge = true;
                                ws.Cells["R11:S11"].Style.Font.Bold = true;
                                //NVL
                                ws.Cells["T11:U11"].Value = "Mã NVL";
                                ws.Cells["T11:U11"].Merge = true;
                                ws.Cells["T11:U11"].Style.Font.Bold = true;
                                //input NVL
                                ws.Cells["V11:x11"].Value = SamplePaper.PaperIn.PAPERNAME_FullIn;
                                ws.Cells["V11:x11"].Merge = true;
                                ws.Cells["V11:x11"].Style.Font.Bold = true;
                                //input kt
                                ws.Cells["L12:X12"].Value = SamplePaper.PaperOut.MODELSPECOut;
                                ws.Cells["L12:X12"].Merge = true;
                                ws.Cells["L12:X12"].Style.Font.Bold = true;
                                //input so trang
                                ws.Cells["L13:X13"].Value = SamplePaper.PaperOut.PAGECNT;
                                ws.Cells["L13:X13"].Merge = true;
                                ws.Cells["L13:X13"].Style.Font.Bold = true;
                                //input cong doan
                                ws.Cells["L14:X14"].Value = SamplePaper.PaperOut.ETC1;
                                ws.Cells["L14:X14"].Merge = true;
                                ws.Cells["L14:X14"].Style.Font.Bold = true;
                                //input mau bia
                                ws.Cells["L15:X15"].Value = SamplePaper.PaperOut.BCOLORCODEOut;
                                ws.Cells["L15:X15"].Merge = true;
                                ws.Cells["L15:X15"].Style.Font.Bold = true;

                                //input mau ruot
                                ws.Cells["L16:X16"].Value = SamplePaper.PaperIn.BCOLORCODEIn;
                                ws.Cells["L16:X16"].Merge = true;
                                ws.Cells["L16:X16"].Style.Font.Bold = true;
                                //input coating
                                ws.Cells["L17:X17"].Value = "";
                                ws.Cells["L17:X17"].Merge = true;
                                ws.Cells["L17:X17"].Style.Font.Bold = true;
                                //input cat sau in
                                ws.Cells["L18:X18"].Value = "";
                                ws.Cells["L18:X18"].Merge = true;
                                ws.Cells["L18:X18"].Style.Font.Bold = true;
                                //Bìa
                                ws.Cells["L19:M19"].Value = "Bìa";
                                ws.Cells["L19:M19"].Merge = true;
                                ws.Cells["L19:M19"].Style.Font.Bold = true;
                                ws.Cells["L19:M19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input bìa
                                ws.Cells["N19:p19"].Value = SamplePaper.PaperOut.TOTALPAGEOUT;
                                ws.Cells["N19:p19"].Merge = true;
                                ws.Cells["N19:p19"].Style.Font.Bold = true;
                                //ruot
                                ws.Cells["Q19:R19"].Value = "Ruột";
                                ws.Cells["Q19:R19"].Merge = true;
                                ws.Cells["Q19:R19"].Style.Font.Bold = true;
                                ws.Cells["Q19:R19"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input ruot 
                                ws.Cells["S19:X19"].Value = SamplePaper.PaperIn.TOTALPAGEIN;
                                ws.Cells["S19:X19"].Merge = true;
                                ws.Cells["S19:X19"].Style.Font.Bold = true;
                                //input keo
                                ws.Cells["L20:X20"].Value = "";
                                ws.Cells["L20:X20"].Merge = true;
                                ws.Cells["L20:X20"].Style.Font.Bold = true;
                                //intput ghim
                                ws.Cells["L21:X21"].Value = "";
                                ws.Cells["L21:X21"].Merge = true;
                                ws.Cells["L21:X21"].Style.Font.Bold = true;
                                //input cat tp
                                ws.Cells["L22:X22"].Value = "";
                                ws.Cells["L22:X22"].Merge = true;
                                ws.Cells["L22:X22"].Style.Font.Bold = true;

                                //input ngày yc
                                ws.Cells["L23:O23"].Value = SamplePaper.PaperOut.DATESTARTAPPROVE;
                                ws.Cells["L23:O23"].Merge = true;
                                ws.Cells["L23:O23"].Style.Font.Bold = true;
                                //
                                ws.Cells["P23:P23"].Value = "~";
                                ws.Cells["P23:P23"].Merge = true;
                                ws.Cells["P23:P23"].Style.Font.Bold = true; 
                                ws.Cells["P23:P23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input ngày hoàn thành
                                ws.Cells["Q23:T23"].Value = SamplePaper.PaperOut.DATEFINISHAPPROVE;
                                ws.Cells["Q23:T23"].Merge = true;
                                ws.Cells["Q23:T23"].Style.Font.Bold = true;
                                //input ngày hoàn thành                                
                                ws.Cells["u23:x23"].Merge = true;
                                ws.Cells["u23:x23"].Style.Font.Bold = true;
                                //input slyc
                                ws.Cells["L24:X24"].Value = SamplePaper.PaperOut.QTYREQUEST;
                                ws.Cells["L24:X24"].Merge = true;
                                ws.Cells["L24:X24"].Style.Font.Bold = true;
                                //input thong tin khac
                                ws.Cells["L25:X25"].Value = SamplePaper.PaperOut.NOTE1;
                                ws.Cells["L25:X25"].Merge = true;
                                ws.Cells["L25:X25"].Style.Font.Bold = true;
                                //input note
                                ws.Cells["L26:X26"].Value = SamplePaper.PaperOut.NOTE2;
                                ws.Cells["L26:X26"].Merge = true;
                                ws.Cells["L26:X26"].Style.Font.Bold = true;
                                //Xóa
                                ws.Cells["A27:X31"].Merge = true;

                                //Taixin

                                ws.Cells["A32:G32"].Value = "Page 1 of 1";
                                ws.Cells["A32:G32"].Merge = true;
                                ws.Cells["A32:G32"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["A32:G32"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                                ws.Cells["H32:Q32"].Value = "Taixin Printing Vina";
                                ws.Cells["H32:Q32"].Merge = true;
                                ws.Cells["H32:Q32"].Style.Font.Bold = true;
                                ws.Cells["H32:Q32"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["H32:Q32"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                //ws.Cells["R32:X32"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                ws.Cells["R32:X32"].Value = DateNow;
                                ws.Cells["R32:X32"].Merge = true;
                                ws.Cells["R32:X32"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["R32:X32"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                ws.PrinterSettings.PaperSize = ePaperSize.A4;
                                ws.PrinterSettings.Orientation = eOrientation.Portrait;
                                ws.PrinterSettings.FitToPage = true;

                                //ws.PrinterSettings.FitToWidth = 0;
                                //ws.PrinterSettings.FitToHeight = 0;
                                ws.PrinterSettings.TopMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.LeftMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.BottomMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.RightMargin = Decimal.Parse("0.25");
                                //ws.PrinterSettings.HorizontalCentered = true;
                                //ws.PrinterSettings.VerticalCentered = true;
                                File.Delete(pathFileExcel);
                                Byte[] bin = p.GetAsByteArray();
                                File.WriteAllBytes(pathFileExcel, bin);
                                exportFileExcel = false;
                                Printed(SamplePaper.PaperOut.IDNumber, "Manual", MainWindow.UserLogin);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Creat Excel File", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        break;
                    }
                case "Box":
                    {
                        if (exportFileExcel == true)
                        {
                            // tạo SaveFileDialog để lưu file excel
                            SaveFileDialog dialog = new SaveFileDialog();

                            // chỉ lọc ra các file có định dạng Excel
                            dialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";

                            // Nếu mở file và chọn nơi lưu file thành công sẽ lưu đường dẫn lại dùng
                            if (dialog.ShowDialog() == true)
                            {
                                pathFileExcel = dialog.FileName;
                            }

                            // nếu đường dẫn null hoặc rỗng thì báo không hợp lệ và return hàm
                            if (string.IsNullOrEmpty(pathFileExcel))
                            {
                                MessageBox.Show("Đường dẫn báo cáo không hợp lệ");
                                return;
                            }
                        }
                        try
                        {
                            using (ExcelPackage p = new ExcelPackage())
                            {
                                //p.Workbook.Properties.Author = DateTime.Now.ToShortDateString();
                                p.Workbook.Properties.Author = DateNow;
                                p.Workbook.Properties.Title = "Sample";
                                p.Workbook.Worksheets.Add("Sheet1");
                                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                                ws.Name = "Sheet1";
                                //ws.Cells.Style.Font.Size = 13;

                                for (int i = 1; i <= 38; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    ws.Cells.Style.Font.Size = 7;
                                    ws.Column(i).Width = 3.7;
                                    ws.Row(i).Height = 20;
                                }
                                //căn hàng và cột cho tất cả các ô
                                for (int i = 1; i < 38; i++)
                                {
                                    string strCell = "A" + i.ToString() + ":" + "X" + i.ToString();
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

                                //Bôi den backgroud
                                for (int i = 5; i < 7; i++)
                                {

                                    string strCell = "A" + i.ToString() + ":" + "A" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Bold = true;
                                }
                                for (int i = 5; i < 7; i++)
                                {

                                    string strCell = "I" + i.ToString() + ":" + "I" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Bold = true;
                                }
                                for (int i = 5; i < 7; i++)
                                {

                                    string strCell = "Q" + i.ToString() + ":" + "Q" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Bold = true;
                                }
                                //Bôi đen cột số 2 phần Qui trình
                                for (int i = 8; i < 38; i++)
                                {

                                    string strCell = "D" + i.ToString() + ":" + "K" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    fill.BackgroundColor.SetColor(System.Drawing.Color.Gainsboro);
                                    cell.Style.Font.Bold = true;
                                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                }

                                //can le ve ben trai 

                                for (int i = 8; i < 38; i++)
                                {

                                    string strCell = "E" + i.ToString() + ":" + "X" + i.ToString();
                                    var cell = ws.Cells[strCell];
                                    var fill = cell.Style.Fill;
                                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                }

                                ws.Cells["A1:A1"].Value = "▲";
                                ws.Cells["A1:X1"].Merge = true;
                                ws.Cells["A1:A1"].Style.Font.Size = 25;
                                ws.Cells["A1:A1"].Style.Font.Bold = true;
                                ws.Row(1).Height = 0;


                                ws.Cells["A2:A2"].Value = "LỆNH SẢN XUẤT MẪU";
                                ws.Cells["A2:X2"].Merge = true;
                                ws.Cells["A2:A2"].Style.Font.Size = 22;
                                ws.Cells["A2:A2"].Style.Font.Bold = true;
                                ws.Row(2).Height = 40;

                                //Ngày SX
                                ws.Cells["A3:B3"].Value = "Ngày : ";
                                ws.Cells["A3:B3"].Merge = true;
                                //ws.Cells["C3:X3"].Value = DateTime.Now.ToString("dd/MM/yyyy");
                                ws.Cells["C3:X3"].Value = DateNow;
                                ws.Cells["C3:X3"].Merge = true;                               
                                ws.Cells["C3:X3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;


                                //CÁCH                            

                                ws.Cells["A4:X4"].Merge = true;
                                ws.Row(4).Height = 3;

                                //MODEL
                                ws.Cells["A5:B5"].Value = "Model";
                                ws.Cells["A5:B5"].Merge = true;
                                ws.Cells["A5:B5"].Style.Font.Bold = true;
                                ws.Cells["C5:H5"].Value = SampleBox.box.modelcode;
                                ws.Cells["C5:H5"].Merge = true;

                                //VER
                                ws.Cells["A6:B6"].Value = "Ver";
                                ws.Cells["A6:B6"].Merge = true;
                                ws.Cells["A6:B6"].Style.Font.Bold = true;
                                ws.Cells["C6:H6"].Value = SampleBox.box.version;
                                ws.Cells["C6:H6"].Merge = true;

                                //Mã hàng
                                ws.Cells["I5:J5"].Value = "Mã Hàng";
                                ws.Cells["I5:J5"].Merge = true;
                                ws.Cells["I5:J5"].Style.Font.Bold = true;
                                ws.Cells["K5:P5"].Value = SampleBox.box.custpartcode;
                                ws.Cells["K5:P5"].Merge = true;

                                //Khách hàng
                                ws.Cells["I6:J6"].Value = "Khách Hàng";
                                ws.Cells["I6:J6"].Merge = true;
                                ws.Cells["I6:J6"].Style.Font.Bold = true;
                                ws.Cells["K6:P6"].Value = SampleBox.box.cust_gb;
                                ws.Cells["K6:P6"].Merge = true;

                                //Tình trạng
                                ws.Cells["Q5:R5"].Value = "Tình Trạng";
                                ws.Cells["Q5:R5"].Merge = true;
                                ws.Cells["Q5:R5"].Style.Font.Bold = true;
                                ws.Cells["S5:X5"].Value = SampleBox.box.information;
                                ws.Cells["S5:X5"].Merge = true;

                                //Yêu cầu
                                ws.Cells["Q6:R6"].Value = "Yêu cầu";
                                ws.Cells["Q6:R6"].Merge = true;
                                ws.Cells["Q6:R6"].Style.Font.Bold = true;
                                ws.Cells["S6:X6"].Value = SampleBox.box.imsempcode;
                                ws.Cells["S6:X6"].Merge = true;

                                //Khoảng trống
                                ws.Cells["A7:X7"].Merge = true;
                                ws.Row(7).Height = 10;

                                //Nguyên liệu
                                ws.Cells["A8:C18"].Value = "SPEC";
                                ws.Cells["A8:C18"].Merge = true;
                                ws.Cells["A8:C18"].Style.Font.Bold = true;
                                ws.Cells["A8:C18"].Style.Font.Size = 16;
                                ws.Cells["A8:C18"].Style.WrapText = true;

                                //GL-upper
                                ws.Cells["D8:K8"].Value = "Giấy in bìa";
                                ws.Cells["D8:K8"].Merge = true;
                                ws.Cells["D8:K8"].Style.Font.Bold = true;
                                //GL
                                ws.Cells["D9:K9"].Value = "Kích thước in";
                                ws.Cells["D9:K9"].Merge = true;
                                ws.Cells["D9:K9"].Style.Font.Bold = true;
                                //GL-lower
                                ws.Cells["D10:K10"].Value = "Giấy in ruột";
                                ws.Cells["D10:K10"].Merge = true;
                                ws.Cells["D10:K10"].Style.Font.Bold = true;
                                //GL
                                ws.Cells["D11:K11"].Value = "Kích thước in";
                                ws.Cells["D11:K11"].Merge = true;
                                ws.Cells["D11:K11"].Style.Font.Bold = true;
                                //GL-silver
                                ws.Cells["D12:K12"].Value = "";
                                ws.Cells["D12:K12"].Merge = true;
                                ws.Cells["D12:K12"].Style.Font.Bold = true;
                                //GL-kt silver
                                ws.Cells["D13:K13"].Value = "Kích thước in";
                                ws.Cells["D13:K13"].Merge = true;
                                ws.Cells["D13:K13"].Style.Font.Bold = true;
                                //In
                                ws.Cells["D14:G17"].Value = "Bìa Cứng";
                                ws.Cells["D14:G17"].Merge = true;
                                ws.Cells["D14:G17"].Style.Font.Bold = true;
                                //In_tren
                                ws.Cells["H14:K14"].Value = "Nắp trên";
                                ws.Cells["H14:K14"].Merge = true;
                                ws.Cells["H14:K14"].Style.Font.Bold = true;
                                //In_duoi
                                ws.Cells["H15:K15"].Value = "Nắp dưới";
                                ws.Cells["H15:K15"].Merge = true;
                                ws.Cells["H15:K15"].Style.Font.Bold = true;
                                //In_cover
                                ws.Cells["H16:K16"].Value = "Cover";
                                ws.Cells["H16:K16"].Merge = true;
                                ws.Cells["H16:K16"].Style.Font.Bold = true;
                                //In_midder
                                ws.Cells["H17:K17"].Value = "Midder";
                                ws.Cells["H17:K17"].Merge = true;
                                ws.Cells["H17:K17"].Style.Font.Bold = true;                                                     
                                //QT
                                ws.Cells["D18:K18"].Value = "Kích thước thành phẩm(Dài/Rộng/Cao)mm";
                                ws.Cells["D18:K18"].Merge = true;
                                ws.Cells["D18:K18"].Style.Font.Bold = true;
                                //Qui trình
                                ws.Cells["A19:C33"].Value = "QUI TRÌNH";
                                ws.Cells["A19:C33"].Merge = true;
                                ws.Cells["A19:C33"].Style.Font.Bold = true;
                                ws.Cells["A19:C33"].Style.Font.Size = 16;
                                ws.Cells["A19:C33"].Style.WrapText = true;
                                //In
                                ws.Cells["D19:G21"].Value = "IN";
                                ws.Cells["D19:G21"].Merge = true;
                                ws.Cells["D19:G21"].Style.Font.Bold = true;
                                //Input_mau
                                ws.Cells["H19:K19"].Value = "Nắp dưới";
                                ws.Cells["H19:K19"].Merge = true;
                                ws.Cells["H19:K19"].Style.Font.Bold = true;
                                //Input_mau
                                ws.Cells["H20:K20"].Value = "Nắp dưới";
                                ws.Cells["H20:K20"].Merge = true;
                                ws.Cells["H14:K20"].Style.Font.Bold = true;
                                //Input_mau
                                ws.Cells["H21:K21"].Value = "Cover";
                                ws.Cells["H21:K21"].Merge = true;
                                ws.Cells["H21:K21"].Style.Font.Bold = true;                               
                                //input_coating
                                ws.Cells["D22:K22"].Value = "Coating(Lớp phủ)";
                                ws.Cells["D22:K22"].Merge = true;
                                ws.Cells["D22:K22"].Style.Font.Bold = true;
                                //Glossy
                                ws.Cells["D23:G24"].Value = "Glossy";
                                ws.Cells["D23:G24"].Merge = true;
                                ws.Cells["D23:G24"].Style.Font.Bold = true;
                                //Glossy_màu
                                ws.Cells["H23:K23"].Value = "Màu";
                                ws.Cells["H23:K23"].Merge = true;
                                ws.Cells["H23:K23"].Style.Font.Bold = true;
                                //Glossy_nội dung
                                ws.Cells["H24:K24"].Value = "Nội dung";
                                ws.Cells["H24:K24"].Merge = true;
                                ws.Cells["H24:K24"].Style.Font.Bold = true;
                                //Stemping Hologram
                                ws.Cells["D25:G30"].Value = "Stemping Hologram";
                                ws.Cells["D25:G30"].Merge = true;
                                ws.Cells["D25:G30"].Style.Font.Bold = true;
                                //Stemping Hologram-mau
                                ws.Cells["H25:K25"].Value = "Màu";
                                ws.Cells["H25:K25"].Merge = true;
                                ws.Cells["H25:K25"].Style.Font.Bold = true;
                                //Stemping Hologram-noi dung
                                ws.Cells["H26:K26"].Value = "Nội dung";
                                ws.Cells["H26:K26"].Merge = true;
                                ws.Cells["H26:K26"].Style.Font.Bold = true;
                                //Stemping Hologram-mau
                                ws.Cells["H27:K27"].Value = "Màu";
                                ws.Cells["H27:K27"].Merge = true;
                                ws.Cells["H27:K27"].Style.Font.Bold = true;
                                //Stemping Hologram-noi dung
                                ws.Cells["H28:K28"].Value = "Nội dung";
                                ws.Cells["H28:K28"].Merge = true;
                                ws.Cells["H28:K28"].Style.Font.Bold = true;
                                //Stemping Hologram-mau
                                ws.Cells["H29:K29"].Value = "Màu";
                                ws.Cells["H29:K29"].Merge = true;
                                ws.Cells["H29:K29"].Style.Font.Bold = true;
                                //Stemping Hologram-noi dung
                                ws.Cells["H30:K30"].Value = "Nội dung";
                                ws.Cells["H30:K30"].Merge = true;
                                ws.Cells["H30:K30"].Style.Font.Bold = true;
                                //Debossing_
                                ws.Cells["D31:G31"].Value = "Debossing";
                                ws.Cells["D31:G31"].Merge = true;
                                ws.Cells["D31:G31"].Style.Font.Bold = true;
                                //Debossing_noi dung
                                ws.Cells["H31:K31"].Value = "Nội dung";
                                ws.Cells["H31:K31"].Merge = true;
                                ws.Cells["H31:K31"].Style.Font.Bold = true;
                                //Imbossing
                                ws.Cells["D32:G32"].Value = "Imbossing";
                                ws.Cells["D32:G32"].Merge = true;
                                ws.Cells["D32:G32"].Style.Font.Bold = true;
                                //Debossing_noi dung
                                ws.Cells["H32:K32"].Value = "Nội dung";
                                ws.Cells["H32:K32"].Merge = true;
                                ws.Cells["H32:K32"].Style.Font.Bold = true;
                                //Bồi
                                ws.Cells["D33:G33"].Value = "Bồi";
                                ws.Cells["D33:G33"].Merge = true;
                                ws.Cells["D33:G33"].Style.Font.Bold = true;
                                //Debossing_noi dung
                                ws.Cells["H33:K33"].Value = "Kiểu bồi";
                                ws.Cells["H33:K33"].Merge = true;
                                ws.Cells["H33:K33"].Style.Font.Bold = true;

                                //Ghi Chú
                                ws.Cells["A34:C37"].Value = "Ghi Chú";
                                ws.Cells["A34:C37"].Merge = true;
                                ws.Cells["A34:C37"].Style.Font.Bold = true;
                                ws.Cells["A34:C37"].Style.Font.Size = 16;
                                ws.Cells["A34:C37"].Style.WrapText = true;

                                //GC
                                ws.Cells["D34:G34"].Value = "Ngày Yêu Cầu";
                                ws.Cells["D34:G34"].Merge = true;
                                ws.Cells["D34:G34"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["H34:K34"].Value = "Ngày Hoàn Thành";
                                ws.Cells["H34:K34"].Merge = true;
                                ws.Cells["H34:K34"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["D35:K35"].Value = "Số lượng yêu cầu";
                                ws.Cells["D35:K35"].Merge = true;
                                ws.Cells["D35:K35"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["D36:K36"].Value = "Những thông tin khác";
                                ws.Cells["D36:K36"].Merge = true;
                                ws.Cells["D36:K36"].Style.Font.Bold = true;
                                //GC
                                ws.Cells["D37:K37"].Value = "Ghi chú";
                                ws.Cells["D37:K37"].Merge = true;
                                ws.Cells["D37:K37"].Style.Font.Bold = true;

                                //input giay bia
                                ws.Cells["L8:X8"].Value = SampleBox.box.paper_name1;
                                ws.Cells["L8:X8"].Merge = true;
                                ws.Cells["L8:X8"].Style.Font.Bold = true;
                                //input KICH THƯƠC
                                ws.Cells["L9:P9"].Value = SampleBox.box.paper_size1;
                                ws.Cells["L9:P9"].Merge = true;
                                ws.Cells["L9:P9"].Style.Font.Bold = true;
                                //DINH DANG
                                ws.Cells["Q9:R9"].Value = "Định dạng";
                                ws.Cells["Q9:R9"].Merge = true;
                                ws.Cells["Q9:R9"].Style.Font.Bold = true;
                                ws.Cells["Q9:R9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input DINH DANG
                                ws.Cells["S9:X9"].Value = SampleBox.box.paper_scale1;
                                ws.Cells["S9:X9"].Merge = true;
                                ws.Cells["S9:X9"].Style.Font.Bold = true;
                                //input LOWER
                                ws.Cells["L10:X10"].Value = SampleBox.box.paper_name2;
                                ws.Cells["L10:X10"].Merge = true;
                                ws.Cells["L10:X10"].Style.Font.Bold = true;
                                //input KICH THƯƠC
                                ws.Cells["L11:P11"].Value = SampleBox.box.paper_size2;
                                ws.Cells["L11:P11"].Merge = true;
                                ws.Cells["L11:P11"].Style.Font.Bold = true;                             
                                //DINH DANG
                                ws.Cells["Q11:R11"].Value = "Định dạng";
                                ws.Cells["Q11:R11"].Merge = true;
                                ws.Cells["Q11:R11"].Style.Font.Bold = true;
                                ws.Cells["Q11:R11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input DINH DANG
                                ws.Cells["S11:X11"].Value = SampleBox.box.paper_scale2;
                                ws.Cells["S11:X11"].Merge = true;
                                ws.Cells["S11:X11"].Style.Font.Bold = true;
                                //input SLIVER
                                ws.Cells["L12:X12"].Value = SampleBox.box.paper_name3;
                                ws.Cells["L12:X12"].Merge = true;
                                ws.Cells["L12:X12"].Style.Font.Bold = true;
                                //input KICH THƯƠC
                                ws.Cells["L13:P13"].Value = SampleBox.box.paper_size3;
                                ws.Cells["L13:P13"].Merge = true;
                                ws.Cells["L13:P13"].Style.Font.Bold = true;
                                //DINH DANG
                                ws.Cells["Q13:R13"].Value = "Định dạng";
                                ws.Cells["Q13:R13"].Merge = true;
                                ws.Cells["Q13:R13"].Style.Font.Bold = true;
                                ws.Cells["Q13:R13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input DINH DANG
                                ws.Cells["S13:X13"].Value = SampleBox.box.paper_scale3;
                                ws.Cells["S13:X13"].Merge = true;
                                ws.Cells["S13:X13"].Style.Font.Bold = true;

                                //input nắp trên
                                //input KICH THƯƠC
                                ws.Cells["L14:P14"].Value = SampleBox.box.cover_up_name1;
                                ws.Cells["L14:P14"].Merge = true;
                                ws.Cells["L14:P14"].Style.Font.Bold = true;
                                //DINH DANG
                                ws.Cells["Q14:R14"].Value = "Kích thước";
                                ws.Cells["Q14:R14"].Merge = true;
                                ws.Cells["Q14:R14"].Style.Font.Bold = true;
                                ws.Cells["Q14:R14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                //input DINH DANG
                                ws.Cells["S14:T14"].Value = SampleBox.box.cover_up_size1;
                                ws.Cells["S14:T14"].Merge = true;
                                ws.Cells["S14:T14"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["U14:V14"].Value = "Định dạng";
                                ws.Cells["U14:V14"].Merge = true;
                                ws.Cells["U14:V14"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["W14:X14"].Value = SampleBox.box.cover_up_scale1;
                                ws.Cells["W14:X14"].Merge = true;
                                ws.Cells["W14:X14"].Style.Font.Bold = true;

                                //input nắp dưới
                                ws.Cells["L15:P15"].Value = SampleBox.box.cover_up_name2;
                                ws.Cells["L15:P15"].Merge = true;
                                ws.Cells["L15:P15"].Style.Font.Bold = true;
                                //DINH DANG
                                ws.Cells["Q15:R15"].Value = "Kích thước";
                                ws.Cells["Q15:R15"].Merge = true;
                                ws.Cells["Q15:R15"].Style.Font.Bold = true;
                                ws.Cells["Q15:R15"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                //input DINH DANG
                                ws.Cells["S15:T15"].Value = SampleBox.box.cover_up_size2;
                                ws.Cells["S15:T15"].Merge = true;
                                ws.Cells["S15:T15"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["U15:V15"].Value = "Định dạng";
                                ws.Cells["U15:V15"].Merge = true;
                                ws.Cells["U15:V15"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["W15:X15"].Value = SampleBox.box.cover_up_scale2;
                                ws.Cells["W15:X15"].Merge = true;
                                ws.Cells["W15:X15"].Style.Font.Bold = true;

                                ws.Cells["L16:P16"].Value = SampleBox.box.cover_up_name3;
                                ws.Cells["L16:P16"].Merge = true;
                                ws.Cells["L16:P16"].Style.Font.Bold = true;
                                //cover
                                //DINH DANG
                                ws.Cells["Q16:R16"].Value = "Kích thước";
                                ws.Cells["Q16:R16"].Merge = true;
                                ws.Cells["Q16:R16"].Style.Font.Bold = true;
                                ws.Cells["Q16:R16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                //input DINH DANG
                                ws.Cells["S16:T16"].Value = SampleBox.box.cover_up_size3;
                                ws.Cells["S16:T16"].Merge = true;
                                ws.Cells["S16:T16"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["U16:V16"].Value = "Định dạng";
                                ws.Cells["U16:V16"].Merge = true;
                                ws.Cells["U16:V16"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["W16:X16"].Value = SampleBox.box.cover_up_scale3;
                                ws.Cells["W16:X16"].Merge = true;
                                ws.Cells["W16:X16"].Style.Font.Bold = true;

                                ws.Cells["L17:P17"].Value = SampleBox.box.cover_up_name4;
                                ws.Cells["L17:P17"].Merge = true;
                                ws.Cells["L17:P17"].Style.Font.Bold = true;
                                //DINH DANG
                                ws.Cells["Q17:R17"].Value = "Kích thước";
                                ws.Cells["Q17:R17"].Merge = true;
                                ws.Cells["Q17:R17"].Style.Font.Bold = true;
                                ws.Cells["Q17:R17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                                //input DINH DANG
                                ws.Cells["S17:T17"].Value = SampleBox.box.cover_up_size4;
                                ws.Cells["S17:T17"].Merge = true;
                                ws.Cells["S17:T17"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["U17:V17"].Value = "Định dạng";
                                ws.Cells["U17:V17"].Merge = true;
                                ws.Cells["U17:V17"].Style.Font.Bold = true;
                                //input DINH DANG
                                ws.Cells["W17:X17"].Value = SampleBox.box.cover_up_scale4;
                                ws.Cells["W17:X17"].Merge = true;
                                ws.Cells["W17:X17"].Style.Font.Bold = true;

                                //input full size
                                ws.Cells["L18:x18"].Value = SampleBox.box.fullsize;
                                ws.Cells["L18:x18"].Merge = true;
                                ws.Cells["L18:x18"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L19:X19"].Value = SampleBox.box.print_color1;
                                ws.Cells["L19:X19"].Merge = true;
                                ws.Cells["L19:X19"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L20:X20"].Value = SampleBox.box.print_color2;
                                ws.Cells["L20:X20"].Merge = true;
                                ws.Cells["L20:X20"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L21:X21"].Value = SampleBox.box.print_color3;
                                ws.Cells["L21:X21"].Merge = true;
                                ws.Cells["L21:X21"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L22:X22"].Value = SampleBox.box.coating;
                                ws.Cells["L22:X22"].Merge = true;
                                ws.Cells["L22:X22"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L23:X23"].Value = SampleBox.box.glossy_color;
                                ws.Cells["L23:X23"].Merge = true;
                                ws.Cells["L23:X23"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L24:X24"].Value = SampleBox.box.glossy_detail;
                                ws.Cells["L24:X24"].Merge = true;
                                ws.Cells["L24:X24"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L25:X25"].Value = SampleBox.box.holo_color1;
                                ws.Cells["L25:X25"].Merge = true;
                                ws.Cells["L25:X25"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L26:X26"].Value = SampleBox.box.holo_detail1;
                                ws.Cells["L26:X26"].Merge = true;
                                ws.Cells["L26:X26"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L27:X27"].Value = SampleBox.box.holo_color2;
                                ws.Cells["L27:X27"].Merge = true;
                                ws.Cells["L27:X27"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L28:X28"].Value = SampleBox.box.holo_detail2;
                                ws.Cells["L28:X28"].Merge = true;
                                ws.Cells["L28:X28"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L29:X29"].Value = SampleBox.box.holo_color3;
                                ws.Cells["L29:X29"].Merge = true;
                                ws.Cells["L29:X29"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L30:X30"].Value = SampleBox.box.holo_detail3;
                                ws.Cells["L30:X30"].Merge = true;
                                ws.Cells["L30:X30"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L31:X31"].Value = SampleBox.box.debosing_detail;
                                ws.Cells["L31:X31"].Merge = true;
                                ws.Cells["L31:X31"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["L32:X32"].Value = SampleBox.box.information;
                                ws.Cells["L32:X32"].Merge = true;
                                ws.Cells["L32:X32"].Style.Font.Bold = true;

                                //QT
                                ws.Cells["L33:P33"].Value = SampleBox.box.boi_detaiil;
                                ws.Cells["L33:P33"].Merge = true;
                                ws.Cells["L33:P33"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["Q33:S33"].Value = "Màu Sóng";
                                ws.Cells["Q33:S33"].Merge = true;
                                ws.Cells["Q33:S33"].Style.Font.Bold = true;
                                //QT
                                ws.Cells["T33:X33"].Value = SampleBox.box.boi_color;
                                ws.Cells["T33:X33"].Merge = true;
                                ws.Cells["T33:X33"].Style.Font.Bold = true;
                                //input ngày yc
                                ws.Cells["L34:O34"].Value = SampleBox.box.sadt;
                                ws.Cells["L34:O34"].Merge = true;
                                ws.Cells["L34:O34"].Style.Font.Bold = true;
                                //
                                ws.Cells["P34:P34"].Value = "~";
                                ws.Cells["P34:P34"].Merge = true;
                                ws.Cells["P34:P34"].Style.Font.Bold = true;
                                ws.Cells["P34:P34"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //input ngày hoàn thành
                                ws.Cells["Q34:T34"].Value = SampleBox.box.fadt;
                                ws.Cells["Q34:T34"].Merge = true;
                                ws.Cells["Q34:T34"].Style.Font.Bold = true;

                                ws.Cells["u34:x34"].Value = "";
                                ws.Cells["u34:x34"].Merge = true;
                                ws.Cells["u34:x34"].Style.Font.Bold = true;
                                //GC-sl
                                ws.Cells["L35:X35"].Value = SampleBox.box.qty;
                                ws.Cells["L35:X35"].Merge = true;
                                ws.Cells["L35:X35"].Style.Font.Bold = true;
                                //GC-thong tin khac
                                ws.Cells["L36:X36"].Value = SampleBox.box.remark1;
                                ws.Cells["L36:X36"].Merge = true;
                                ws.Cells["L36:X36"].Style.Font.Bold = true;
                                //GC-ghi chu
                                ws.Cells["L37:X37"].Value = SampleBox.box.remark2;
                                ws.Cells["L37:X37"].Merge = true;
                                ws.Cells["L37:X37"].Style.Font.Bold = true;
                                //Taixin
                                ws.Cells["A38:G38"].Value = "Page 1 of 1";
                                ws.Cells["A38:G38"].Merge = true;
                                ws.Cells["A38:G38"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["A38:G38"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                                ws.Cells["H38:Q38"].Value = "Taixin Printing Vina";
                                ws.Cells["H38:Q38"].Merge = true;
                                ws.Cells["H38:Q38"].Style.Font.Bold = true;
                                ws.Cells["H38:Q38"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["H38:Q38"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                //ws.Cells["R38:X38"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                ws.Cells["R38:X38"].Value = DateNow;
                                ws.Cells["R38:X38"].Merge = true;
                                ws.Cells["R38:X38"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                ws.Cells["R38:X38"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                ws.PrinterSettings.PaperSize = ePaperSize.A4;
                                ws.PrinterSettings.Orientation = eOrientation.Portrait;
                                ws.PrinterSettings.FitToPage = true;

                                //ws.PrinterSettings.FitToWidth = 0;
                                //ws.PrinterSettings.FitToHeight = 0;
                                ws.PrinterSettings.TopMargin = Decimal.Parse("0");
                                ws.PrinterSettings.LeftMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.BottomMargin = Decimal.Parse("0.25");
                                ws.PrinterSettings.RightMargin = Decimal.Parse("0.25");
                                //ws.PrinterSettings.HorizontalCentered = true;
                                //ws.PrinterSettings.VerticalCentered = true;
                                File.Delete(pathFileExcel);
                                Byte[] bin = p.GetAsByteArray();
                                File.WriteAllBytes(pathFileExcel, bin);
                                exportFileExcel = false;
                                Printed(SampleBox.box.samno, "Box",MainWindow.UserLogin);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Creat Excel File", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        break;
                    }
            };
            
        }

        public void PrintFileExcel()
        {
            CreatFileExcel("","", "", "", "");
            // Open the Workbook:           
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(pathFileExcel);
            PrintDialog dialog = new PrintDialog();
            dialog.UserPageRangeEnabled = true;
            PageRange rang = new PageRange(1, 3);
            dialog.PageRange = rang;
            PageRangeSelection seletion = PageRangeSelection.UserPages;
            dialog.PageRangeSelection = seletion;
            //dialog.ShowDialog();
            PrintDocument pd = workbook.PrintDocument;
            //PaperSize A5Sizes = new PaperSize("8.5x13", 1250, 800);
            PaperSize A4Sizes = new PaperSize("8.5x13", 1850, 1250);
            pd.PrinterSettings.DefaultPageSettings.PaperSize = A4Sizes;
            pd.Print();            
            File.Delete(pathFileExcel);
            //this.Hide();
        }

        int index = 0;
        public bool check = false;
        string pathTemp = "";
        string pathFolderExcel = "";
        public void ViewExcelFile(string _pathFileExcel)
        {
            try
            {

                if (check == false)
                {
                    if (!Directory.Exists(@"TempFile"))
                        Directory.CreateDirectory(@"TempFile");
                    for (int i = 0; i < 10000; i++)
                    {
                        string fileTemp = @"TempFile//TempFileXps" + i.ToString() + ".xpsx";
                        if (File.Exists(fileTemp))
                        {
                            File.Delete(fileTemp);
                        }
                    }
                    check = true;
                }
                pathFolderExcel = @"\excel.xlsx";
                if (!Directory.Exists(pathFolderExcel))
                {
                    File.Delete(pathFolderExcel);
                }    
                File.Copy(pathFileExcel, pathFolderExcel);
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(_pathFileExcel, ExcelVersion.Version2010);
                pathTemp = @"TempFile//TempFileXps" + index.ToString() + ".xpsx";
                workbook.SaveToFile(pathTemp, Spire.Xls.FileFormat.XPS);
                //workbook.SaveToFile(pathTemp, Spire.Xls.FileFormat.PDF);
                XpsDocument xpsDocument = new XpsDocument(pathTemp, FileAccess.Read);
                view.Document = xpsDocument.GetFixedDocumentSequence();
                view.FitToHeight();
                index++;
                MainWindow.checkPrint = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "View Excel File", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
            //Application.Current.Shutdown();
        }

        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            PrintFileExcel();
        }

        private void btnXls_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();                   
                sfd.ShowDialog();
                if(sfd.FileName != "")
                {                   
                    File.Copy(pathFolderExcel, sfd.FileName+".xlsx");
                    MessageBox.Show("Xuất file Excel thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                } 
                else
                {
                    MessageBox.Show("Vui lòng chọn vị trí lưu tập tin", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                } 
                
            }
            catch
            {
                MessageBox.Show("Tên file trùng với một file có sẵn.\nVui lòng nhập một tên mới", "Print/SaveFileExcel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public void CreatListExcel()
        {
            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    int numberRow = 0;
                    if(MainWindow.pl_Print=="Box")
                    {
                        foreach (var item in SampleBox.listSampleExportExcel)
                        {
                            if (item.checkXLS == "True")
                            {
                                numberRow++;
                            }
                        }
                    }
                    if (MainWindow.pl_Print == "Manual")
                    {
                        foreach (var item in Page_Sample_Manual.listSampleManual)
                        {
                            if (item.checkXLS == "True")
                            {
                                numberRow++;
                            }
                        }
                    }  
                    
                    numberRow = numberRow + 5;
                    p.Workbook.Properties.Author = DateNow;
                    //p.Workbook.Properties.Author = DateTime.Now.ToShortDateString();
                    p.Workbook.Properties.Author = DateNow;
                    p.Workbook.Properties.Title = "Sample";
                    p.Workbook.Worksheets.Add("Sheet1");
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];
                    ws.Name = "Sheet1";

                    //Cột 1 
                    ws.Column(1).Width = 10;//stt
                    ws.Column(2).Width = 40;//model
                    ws.Column(3).Width = 40;//code
                    ws.Column(4).Width = 10;//ver
                    ws.Column(5).Width = 30;//color
                    ws.Column(6).Width = 30;//sl
                    ws.Column(7).Width = 15;//ngày yc
                    ws.Column(8).Width = 15;//ngày ht                
                    ws.Column(9).Width = 15;//kh

                    ws.Row(1).Height = 10;
                    ws.Row(2).Height = 40;
                    ws.Row(3).Height = 20;
                    ws.Row(4).Height = 25;

                    //căn hàng và cột cho tất cả các ô                 


                    for (int i = 1; i < numberRow; i++)
                    {
                        string strCell = "A" + i.ToString() + ":" + "I" + i.ToString();
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
                        string strCell = "A" + i.ToString() + ":" + "I" + i.ToString();
                        var cell = ws.Cells[strCell];
                        ws.Row(i).Height = 25;
                        cell.Style.Font.Size = 11;
                        cell.Style.Font.Bold = false;

                        string strCell1 = "A" + i.ToString() + ":" + "A" + i.ToString();
                        ws.Cells[strCell1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell2 = "B" + i.ToString() + ":" + "B" + i.ToString();
                        ws.Cells[strCell2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        //--
                        string strCell3 = "C" + i.ToString() + ":" + "C" + i.ToString();
                        ws.Cells[strCell3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        //--
                        string strCell4 = "D" + i.ToString() + ":" + "D" + i.ToString();
                        ws.Cells[strCell4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //--
                        string strCell5 = "E" + i.ToString() + ":" + "E" + i.ToString();
                        ws.Cells[strCell5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        //--
                        string strCell6 = "F" + i.ToString() + ":" + "F" + i.ToString();
                        ws.Cells[strCell6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
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


                    //for (int i = 5; i < numberRow; i++)
                    //{
                    //    if (i % 2 == 0)
                    //    {
                    //        string strCell = "A" + i.ToString() + ":" + "I" + i.ToString();
                    //        var cell = ws.Cells[strCell];
                    //        var fill = cell.Style.Fill;
                    //        fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //        fill.BackgroundColor.SetColor(System.Drawing.Color.AliceBlue);
                    //    }
                    //}

                    //Bôi den backgroud
                    //

                    //ws.Cells["A2:I2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //ws.Cells["A2:I2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Azure);

                    //ws.Cells["A4:I4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //ws.Cells["A4:I4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Ivory);


                    ws.Cells["A1:A1"].Value = "";
                    ws.Cells["A1:I1"].Merge = true;
                    ws.Cells["A1:A1"].Style.Font.Size = 25;
                    ws.Cells["A1:A1"].Style.Font.Bold = true;


                    ws.Cells["A2:A2"].Value = "LIST SAMPLE";
                    ws.Cells["A2:I2"].Merge = true;
                    ws.Cells["A2:A2"].Style.Font.Size = 22;
                    ws.Cells["A2:A2"].Style.Font.Bold = true;


                    //Ngày SX
                    //ws.Cells["A3:A3"].Value = "Ngày : " + DateTime.Now.ToString("dd/MM/yyyy");
                    ws.Cells["A3:A3"].Value = "Ngày : " + DateNow;
                    ws.Cells["A3:I3"].Merge = true;
                    ws.Cells["A3:A3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    ws.Cells["A3:A3"].Style.Font.Bold = true;


                    //Head                  
                    ws.Cells["A4:I4"].Style.Font.Size = 12;
                    ws.Cells["A4:I4"].Style.Font.Bold = true;
                    ws.Cells["A4:A4"].Value = "STT";  
                    ws.Cells["B4:B4"].Value = "Model";         
                    ws.Cells["C4:C4"].Value = "Code";          
                    ws.Cells["D4:D4"].Value = "Ver";           
                    ws.Cells["E4:E4"].Value = "Color";                   
                    ws.Cells["F4:F4"].Value = "S/L";                   
                    ws.Cells["G4:G4"].Value = "Ngày YC";                   
                    ws.Cells["H4:H4"].Value = "Ngày HT";                   
                    ws.Cells["I4:I4"].Value = "Customer";                  

                    int index = 4;
                    int stt = 0;

                    if(MainWindow.pl_Print == "Box")
                    {
                        foreach (var item in SampleBox.listSampleExportExcel)
                        {
                            if (item.checkXLS == "True")
                            {
                                index++;
                                stt++;
                                //--
                                string strCell1 = "A" + index.ToString() + ":" + "A" + index.ToString();
                                ws.Cells[strCell1].Value = stt;
                                //--
                                string strCell2 = "B" + index.ToString() + ":" + "B" + index.ToString();
                                ws.Cells[strCell2].Value = item.modelcode;
                                //--
                                string strCell3 = "C" + index.ToString() + ":" + "C" + index.ToString();
                                ws.Cells[strCell3].Value = item.custpartcode;
                                //--
                                string strCell4 = "D" + index.ToString() + ":" + "D" + index.ToString();
                                ws.Cells[strCell4].Value = item.version;
                                //--
                                string strCell5 = "E" + index.ToString() + ":" + "E" + index.ToString();
                                ws.Cells[strCell5].Value = item.remark1;
                                //--
                                string strCell6 = "F" + index.ToString() + ":" + "F" + index.ToString();
                                ws.Cells[strCell6].Value = item.qty;
                                //--
                                string strCell7 = "G" + index.ToString() + ":" + "G" + index.ToString();
                                ws.Cells[strCell7].Value = item.sadt.Substring(0, 10);
                                //--
                                string strCell8 = "H" + index.ToString() + ":" + "H" + index.ToString();
                                ws.Cells[strCell8].Value = item.fadt.Substring(0, 10);
                                //--
                                string strCell9 = "I" + index.ToString() + ":" + "I" + index.ToString();
                                ws.Cells[strCell9].Value = item.cust_gb;
                            }

                        }
                    }
                    if (MainWindow.pl_Print == "Manual")
                    {
                        foreach (var item in Page_Sample_Manual.listSampleManual)
                        {
                            if (item.checkXLS == "True")
                            {
                                index++;
                                stt++;
                                //--
                                string strCell1 = "A" + index.ToString() + ":" + "A" + index.ToString();
                                ws.Cells[strCell1].Value = stt;
                                //--
                                string strCell2 = "B" + index.ToString() + ":" + "B" + index.ToString();
                                ws.Cells[strCell2].Value = item.CUSTMODELCODE;
                                //--
                                string strCell3 = "C" + index.ToString() + ":" + "C" + index.ToString();
                                ws.Cells[strCell3].Value = item.CUSTPARTCODE;
                                //--
                                string strCell4 = "D" + index.ToString() + ":" + "D" + index.ToString();
                                ws.Cells[strCell4].Value = item.VERSION;
                                //--
                                string strCell5 = "E" + index.ToString() + ":" + "E" + index.ToString();
                                ws.Cells[strCell5].Value = item.NOTE1;
                                //--
                                string strCell6 = "F" + index.ToString() + ":" + "F" + index.ToString();
                                ws.Cells[strCell6].Value = item.QTYREQUEST;
                                //--
                                string strCell7 = "G" + index.ToString() + ":" + "G" + index.ToString();
                                ws.Cells[strCell7].Value = item.DATESTARTAPPROVE.Substring(0, 10);
                                //--
                                string strCell8 = "H" + index.ToString() + ":" + "H" + index.ToString();
                                ws.Cells[strCell8].Value = item.DATEFINISHAPPROVE.Substring(0, 10);
                                //--
                                string strCell9 = "I" + index.ToString() + ":" + "I" + index.ToString();
                                ws.Cells[strCell9].Value = item.CUST_GB;
                            }

                        }
                    }

                    

                    ws.PrinterSettings.PaperSize = ePaperSize.A4;
                    ws.PrinterSettings.Orientation = eOrientation.Landscape;
                    ws.PrinterSettings.FitToPage = true;
                    //ws.Cells["A4:I4"].AutoFilter = true;

                    //ws.PrinterSettings.FitToWidth = 0;
                    //ws.PrinterSettings.FitToHeight = 0;
                    ws.PrinterSettings.TopMargin = Decimal.Parse("0");
                    ws.PrinterSettings.LeftMargin = Decimal.Parse("0.25");
                    ws.PrinterSettings.BottomMargin = Decimal.Parse("0.25");
                    ws.PrinterSettings.RightMargin = Decimal.Parse("0.25");
                    //ws.PrinterSettings.HorizontalCentered = true;
                    //ws.PrinterSettings.VerticalCentered = true;
                    File.Delete(pathFileExcel);
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(pathFileExcel, bin);
                    exportFileExcel = false;
                    //Printed(sampleBox.samno, "Box", MainWindow.UserLogin);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "CreatListExcel", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnListXls_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.ShowDialog();
                if (sfd.FileName != "")
                {
                    CreatListExcel();
                    File.Copy(pathFileExcel, sfd.FileName + ".xlsx");
                    MessageBox.Show("Xuất file Excel thành công", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                }                
                else
                {
                    MessageBox.Show("Vui lòng chọn vị trí lưu tập tin", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch
            {
                MessageBox.Show("Tên file trùng với một file có sẵn.\nVui lòng nhập một tên mới", "Print/SaveFileExcel", MessageBoxButton.OK, MessageBoxImage.Error);
            }              
        }
    }
    
}
