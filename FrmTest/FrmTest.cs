using GridReport.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FrmTest
{
    public partial class FrmTest : Form
    {
        public FrmTest()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PrintReportHelper.Print(new ReportArg
            {
                GrfName = "Report/TestReport.grf",
                ShowPreview = true,
                Data = new
                {
                    Parameter = new
                    {
                        Name = "漆鹏举",
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
