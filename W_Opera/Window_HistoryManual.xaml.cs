using DataHelper;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System;
using System.Collections.Generic;
using System.Data;
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
using W_Opera.DAO;

namespace W_Opera
{
    /// <summary>
    /// Interaction logic for Window_HistoryManual.xaml
    /// </summary>
    public partial class Window_HistoryManual : Window
    {
        public Window_HistoryManual()
        {
            InitializeComponent();
            Loaded += ListHistoryManual_Loaded;
        }
        public string path_sql;
        public string ModelCode;
        //public string path_sql = "Data Source=(local);Initial Catalog=taixin_db;Integrated Security=True;";
        

        private void ListHistoryManual_Loaded(object sender, RoutedEventArgs e)
        {
            path_sql = MainWindow.path_sql;
            ModelCode = Page_Sample_Manual.ModelCodeHist;

            try
            {
                using (SqlConnection conn = new SqlConnection(path_sql))
                {
                    conn.Open();

                    var command = "select * from tbSampleManualHistory where custpartcode = @Modelcode order by insdthis";
                    var listHist = DataProvider.Instance.executeQuery(path_sql, command, new object[] { ModelCode });

                    List<Helper_TaixinDB_Model> listAll = new List<Helper_TaixinDB_Model>();
                    foreach(DataRow rowA in listHist.Rows)
                    {
                        Helper_TaixinDB_Model emp = new Helper_TaixinDB_Model();
                        emp.MODELCODE = rowA["modelcode"].ToString();
                        emp.CUSTMODELCODE = rowA["custmodelcode"].ToString();
                        emp.CUSTPARTCODE = rowA["custpartcode"].ToString();
                        emp.VERSIONUP = rowA["versionup"].ToString();
                        emp.VERSION = rowA["version"].ToString();

                        emp.CUST_GB = rowA["cust_gb"].ToString();
                        emp.INSEMPCODE = rowA["insempcode"].ToString();
                        emp.PAPERNAMEOut = rowA["papernameout"].ToString();
                        emp.HEIGHTOut = rowA["heightOut"].ToString();
                        emp.PHCOUNTOut = rowA["phcountOut"].ToString();
                        emp.PAPERNAME_FullOut = rowA["papernameFullOut"].ToString();

                        emp.PAPERNAMEIn = rowA["papernameIn"].ToString();
                        emp.HEIGHTIn = rowA["heightin"].ToString();
                        emp.PHCOUNTIn = rowA["phcountin"].ToString();
                        emp.PAPERNAME_FullIn = rowA["papernameFullIn"].ToString();
                        emp.MODELSPECOut = rowA["modelspecOut"].ToString();

                        emp.PAGECNT = rowA["pagecnt"].ToString();
                        emp.ETC1 = rowA["process"].ToString();
                        emp.BCOLORCODEOut = rowA["bcolorcodeOut"].ToString();
                        emp.BCOLORCODEIn = rowA["bcolorcodeIn"].ToString();
                        emp.TOTALPAGEOUT = rowA["totalpageOut"].ToString();

                        emp.TOTALPAGEIN = rowA["totalpageIn"].ToString();
                        emp.DATESTARTAPPROVE = rowA["Sadt"].ToString();
                        emp.DATEFINISHAPPROVE = rowA["Fadt"].ToString();
                        emp.QTYREQUEST = rowA["qtyRequest"].ToString();
                        emp.NOTE1 = rowA["RemarkDif"].ToString();

                        emp.NOTE2 = rowA["Remark"].ToString();
                        emp.ETC3 = rowA["qtyAttach"].ToString();
                        emp.MA = rowA["dep_mar"].ToString();
                        emp.SX = rowA["dep_pro"].ToString();
                        emp.CL = rowA["dep_qc"].ToString();

                        emp.RD = rowA["dep_rnd"].ToString();
                        emp.KD = rowA["dep_pur"].ToString();
                        emp.SeqHis = rowA["SeqHis"].ToString();
                        emp.StatusHis = rowA["StatusHis"].ToString();
                        emp.InsempcodeHis = rowA["imsempcodehis"].ToString();

                        emp.InsdtHis = rowA["insdthis"].ToString();
                        


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
