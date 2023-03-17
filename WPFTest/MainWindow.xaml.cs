using GridReport.Common;
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

namespace WPFTest
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            PrintReportHelper.Print(new ReportArg
            {
                GrfName = "Report/TestReport.grf",
                ShowPreview = true,
                Data = new
                {
                    Parameter = new
                    {
                        Name = "test",
                        BarCode = "TJ50001234"
                    },
                    Table = new List<DeptModel> { new DeptModel
                                         {
                        DeptName="一般检查",ItemName="身高\r\n体重\r\n身高\r\n体重\r\n",Memo="参数" },
                    new DeptModel
                                         {
                        DeptName="一般检查",ItemName="身高\r\n体重\r\n",Memo="参数" },
                    new DeptModel
                                         {
                        DeptName="一般检查",ItemName="身高\r\n体重\r\n",Memo="参数" },
                    new DeptModel
                                         {
                        DeptName="一般检查",ItemName="身高\r\n体重\r\n",Memo="参数" },}

                }
            });
        }
    }
    public class DeptModel
    {
        public string DeptName { get; set; }
        public string ItemName { get; set; }
        public string Memo { get; set; }
    }
}
