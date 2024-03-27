using DataHelper;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
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
using W_Opera.DAO;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Window_HistoryBox.xaml
    /// </summary>
    public partial class Window_HistoryBox : Window
    {
        public Window_HistoryBox()
        {
            InitializeComponent();
            Loaded += ListHistoryBox_Loaded;
        }

        public string path_sql;
        public string ModelCode;

        private void ListHistoryBox_Loaded(object sender, RoutedEventArgs e)
        {
            path_sql = MainWindow.path_sql;
            ModelCode = Page_Sample_Box.ModelCodeHistBox;

            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "select * from tbSampleBoxHistory where custpartcode = @Modelcode order by insdthis";
                    var listHist = DataProvider.Instance.executeQuery(path_sql, command, new object[] { ModelCode });

                    List<Helper_TaixinDB_SampleBox> listAll = new List<Helper_TaixinDB_SampleBox>();
                    foreach (DataRow rowA in listHist.Rows)
                    {
                        Helper_TaixinDB_SampleBox emp = new Helper_TaixinDB_SampleBox();
                        emp.modelcode = rowA["modelcode"].ToString();
                        emp.custmodelcode = rowA["custpartcode"].ToString();
                        emp.custpartcode = rowA["custpartcode"].ToString();
                        emp.information = rowA["information"].ToString();
                        emp.version = rowA["version"].ToString();

                        emp.cust_gb = rowA["cust_gb"].ToString();
                        emp.imsempcode = rowA["imsempcode"].ToString();
                        emp.paper_name1 = rowA["paper_name1"].ToString();
                        emp.paper_size1 = rowA["paper_size1"].ToString();
                        emp.paper_scale1 = rowA["paper_scale1"].ToString();

                        emp.paper_name2 = rowA["paper_name2"].ToString();
                        emp.paper_size2 = rowA["paper_size2"].ToString();
                        emp.paper_scale2 = rowA["paper_scale2"].ToString();

                        emp.paper_name3 = rowA["paper_name3"].ToString();
                        emp.paper_size3 = rowA["paper_size3"].ToString();
                        emp.paper_scale3 = rowA["paper_scale3"].ToString();

                        emp.cover_up_name1 = rowA["cover_up_name1"].ToString();
                        emp.cover_up_size1 = rowA["cover_up_size1"].ToString();
                        emp.cover_up_scale1 = rowA["cover_up_scale1"].ToString();

                        emp.cover_up_name2 = rowA["cover_up_name2"].ToString();
                        emp.cover_up_size2 = rowA["cover_up_size2"].ToString();
                        emp.cover_up_scale2 = rowA["cover_up_scale2"].ToString();

                        emp.cover_up_name3 = rowA["cover_up_name3"].ToString();
                        emp.cover_up_size3 = rowA["cover_up_size3"].ToString();
                        emp.cover_up_scale3 = rowA["cover_up_scale3"].ToString();

                        emp.cover_up_name4 = rowA["cover_up_name4"].ToString();
                        emp.cover_up_size4 = rowA["cover_up_size4"].ToString();
                        emp.cover_up_scale4 = rowA["cover_up_scale4"].ToString();

                        emp.printindex1 = rowA["printindex1"].ToString();
                        emp.printindex2 = rowA["printindex2"].ToString();
                        emp.printindex3 = rowA["printindex3"].ToString();

                        emp.print_color1 = rowA["print_color1"].ToString();
                        emp.print_color2 = rowA["print_color2"].ToString();
                        emp.print_color3 = rowA["print_color3"].ToString();

                        emp.coating = rowA["coating"].ToString();
                        emp.glossy_detail = rowA["glossy_detail"].ToString();
                        emp.holo_color1 = rowA["holo_color1"].ToString();
                        emp.holo_detail1 = rowA["holo_detail1"].ToString();
                        emp.holo_color2 = rowA["holo_color2"].ToString();
                        emp.holo_detail2 = rowA["holo_detail2"].ToString();

                        emp.holo_color3 = rowA["holo_color3"].ToString();
                        emp.holo_detail3 = rowA["holo_detail3"].ToString();
                        //emp.holo_color4 = rowA["holo_color4"].ToString();
                        //emp.holo_detail4 = rowA["holo_detail4"].ToString();

                        emp.debosing_detail = rowA["debosing_detail"].ToString();
                        emp.imbosing_detail = rowA["imbosing_detail"].ToString();
                        emp.boi_detaiil = rowA["boi_detaiil"].ToString();
                        emp.boi_color = rowA["boi_color"].ToString();

                        emp.sadt = rowA["sadt"].ToString();
                        emp.fadt = rowA["fadt"].ToString();
                        emp.qty = rowA["qty"].ToString();

                        emp.remark1 = rowA["remark1"].ToString();
                        emp.forecast = rowA["forecast"].ToString();
                        emp.remark2 = rowA["remark2"].ToString();

                        emp.qtyAttach = rowA["qtyAttach"].ToString();
                        emp.seqhis = rowA["SeqHis"].ToString();
                        emp.statushis = rowA["StatusHis"].ToString();
                        emp.insempcodehis = rowA["imsempcodehis"].ToString();

                        emp.insdthis = rowA["insdthis"].ToString();
                        emp.dep_mar = rowA["dep_mar"].ToString();
                        emp.dep_pro = rowA["dep_pro"].ToString();
                        emp.dep_qc = rowA["dep_qc"].ToString();
                        emp.dep_rnd = rowA["dep_rnd"].ToString();
                        emp.dep_pur = rowA["dep_pur"].ToString();





                        listAll.Add(emp);
                    }
                    // Sắp xếp và thêm STT
                    //listAll = listAll.OrderBy(x => x.EmpId).ToList();
                    int i = 1;
                    listAll.ForEach(x =>
                    {
                        x.ID = i;
                        i++;
                    });

                    ListHistory.ItemsSource = listAll;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Manual/ReadSampleNo", MessageBoxButton.OK, MessageBoxImage.Error);

            }

        }
    }
}
