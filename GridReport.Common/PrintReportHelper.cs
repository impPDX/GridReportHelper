using gregn6Lib;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GridReport.Common
{
    public class PrintReportHelper
    {
        private const string CONST_DATA_SUBREPORTS = "SubReports";
        private const string CONST_DATA_SUBREPORTS_NAME = "Name";
        private const string CONST_DATA_PARAMETER = "Parameter";
        private const string CONST_DATA_TABLE = "Table";
        private const string CONST_DATA_PICTURE_PREFIX = "data:image/png;base64,";
        private const string CONST_REPORT_ROOT = "xml";
        private const string CONST_REPORT_ROW = "row";
        private const string CONST_REPORT_PICTURE_PREFIX = "Picture:";

        public enum ExportType
        {
            /// <summary>
            /// Excel
            /// </summary>
            Excel = 0,

            /// <summary>
            /// RTF
            /// </summary>
            RTF = 1,

            /// <summary>
            /// PDF
            /// </summary>
            PDF = 2,

            /// <summary>
            /// Html
            /// </summary>
            Html = 3,

            /// <summary>
            /// Image
            /// </summary>
            Image = 4,

            /// <summary>
            /// Text
            /// </summary>
            Text = 5,

            /// <summary>
            /// CSV
            /// </summary>
            CSV = 6,
        }


        public static GRExportType GetExportType(ExportType exportType)
        {
            GRExportType gExportType = GRExportType.gretPDF;
            switch (exportType)
            {
                case ExportType.Excel:
                    gExportType = GRExportType.gretXLS;
                    break;
                case ExportType.RTF:
                    gExportType = GRExportType.gretRTF;
                    break;
                case ExportType.PDF:
                    gExportType = GRExportType.gretPDF;
                    break;
                case ExportType.Html:
                    gExportType = GRExportType.gretHTM;
                    break;
                case ExportType.Image:
                    gExportType = GRExportType.gretIMG;
                    break;
                case ExportType.Text:
                    gExportType = GRExportType.gretTXT;
                    break;
                case ExportType.CSV:
                    gExportType = GRExportType.gretCSV;
                    break;
            }
            return gExportType;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arg">打印配置参数</param>
        /// <param name="isPrint">true打印 false导出</param>
        private static bool printAndExport(ReportArg arg, bool isPrint)
        {
            try
            {
                GridppReport report = new GridppReport();

                //加载报表模板
                if (string.IsNullOrEmpty(arg.GrfName))
                    throw new Exception($"未指定打印模板文件，无法打印");

                arg.GrfName = arg.GrfName.Replace("/", "\\");
                string grfPath = Path.Combine(Application.StartupPath, arg.GrfName);
                if (File.Exists(grfPath))
                    report.LoadFromFile(grfPath);
                else
                    throw new Exception($"打印模板文件{arg.GrfName}丢失，可能版本部署有误，请咨询管理员");
                //打印机
                report.Printer.PrinterName = arg.Printer ?? "";


                if (arg.paperSize != 0)
                {
                    report.Printer.PaperSize = short.Parse(arg.paperSize.GetHashCode().ToString());
                    report.DesignPaperSize = short.Parse(arg.paperSize.GetHashCode().ToString());
                }

                if (arg.Poaoero != PaperOrientation.默认)
                {
                    if (arg.Poaoero == PaperOrientation.纵向)
                    {
                        report.Printer.PaperOrientation = GRPaperOrientation.grpoPortrait;
                        report.DesignPaperOrientation = GRPaperOrientation.grpoPortrait;
                    }

                    if (arg.Poaoero == PaperOrientation.横向)
                    {
                        report.Printer.PaperOrientation = GRPaperOrientation.grpoLandscape;
                        report.DesignPaperOrientation = GRPaperOrientation.grpoLandscape;
                        if (arg.PaperWidth != 0.0)
                        {
                            report.DesignPaperWidth = arg.PaperWidth;
                        }
                        if (arg.PaperLength != 0.0)
                        {
                            report.DesignPaperLength = arg.PaperLength;
                        }
                        if (arg.PaperWidth != 0.0)
                        {
                            report.Printer.PaperWidth = arg.PaperWidth;
                        }
                        if (arg.PaperLength != 0.0)
                        {
                            report.Printer.PaperLength = arg.PaperLength;
                        }
                        report.PrintAsDesignPaper = true;
                    }
                }

                if (arg.LeftMargin != 0)
                {
                    report.Printer.LeftMargin = arg.LeftMargin;
                    report.DesignLeftMargin = arg.LeftMargin;
                }

                if (arg.RightMargin != 0)
                {
                    report.Printer.RightMargin = arg.RightMargin;
                    report.DesignRightMargin = arg.RightMargin;
                }

                if (arg.TopMargin != 0)
                {
                    report.Printer.TopMargin = arg.TopMargin;
                    report.DesignTopMargin = arg.TopMargin;
                }

                if (arg.BottomMargin != 0)
                {
                    report.Printer.BottomMargin = arg.BottomMargin;
                    report.DesignBottomMargin = arg.BottomMargin;
                }


                //序列化报表数据
                JObject data = JObject.FromObject(arg.Data);

                //加载参数数据
                loadParameter(report, data[CONST_DATA_PARAMETER]);

                //加载表格数据
                loadTable(report, data[CONST_DATA_TABLE]);

                //子报表
                if (data[CONST_DATA_SUBREPORTS] != null && data[CONST_DATA_SUBREPORTS] is JArray)
                {
                    Dictionary<GridppReport, JToken> subReportDatas = new Dictionary<GridppReport, JToken>();
                    //根据数据源自动识别子报表
                    foreach (var subReportData in data[CONST_DATA_SUBREPORTS] as JArray)
                    {
                        var ctrl = report.ControlByName(subReportData[CONST_DATA_SUBREPORTS_NAME].ToString());
                        if (ctrl == null) continue;

                        GridppReport subReport = ctrl.AsSubReport.Report;
                        subReportDatas.Add(subReport, subReportData);

                        //加载子报表参数
                        loadParameter(subReport, subReportData[CONST_DATA_PARAMETER]);
                    }

                    if (subReportDatas.Count > 0)
                        subReportDatas.First().Key.FetchRecord += () =>
                        {
                            //加载子报表表格数据
                            foreach (var item in subReportDatas)
                            {
                                GridppReport subReport = item.Key;
                                JToken subReportData = item.Value;

                                //子报表数据的筛选条件
                                Dictionary<string, string> dictParams = new Dictionary<string, string>();
                                foreach (IGRParameter para in subReport.Parameters)
                                {
                                    foreach (IGRField fld in report.DetailGrid.Recordset.Fields)
                                        if (string.Compare(fld.Name, para.Name, true) == 0)
                                            dictParams.Add(para.Name, para.AsString);
                                }

                                //填充数据
                                JArray subData = subReportData[CONST_DATA_TABLE] as JArray;
                                if (dictParams.Count > 0)
                                {
                                    IEnumerable<JToken> subDataList = subData.ToList();
                                    foreach (var param in dictParams)
                                        subDataList = subDataList.Where(o => o[param.Key] != null && o[param.Key].ToString() == param.Value);
                                    loadTable(subReport, new JArray(subDataList.ToArray()));
                                }
                                else
                                    loadTable(subReport, subData);
                            }
                        };
                }

                //打印预览/直接打印
                if (isPrint)
                {
                    if (arg.ShowPreview)
                        report.PrintPreview(true);
                    else
                    {
                        //设置打印机
                        report.Print(false);
                    }
                }
                else
                {
                    SaveFileDialog FileDialog = new SaveFileDialog();
                    ExportType exportType = ExportType.Excel;
                    switch (exportType)
                    {
                        case ExportType.Excel:
                            {
                                FileDialog.Filter = "Excel 97-2003 工作簿(*.xls)|*.xls|工作簿(*.xlsx)|*.xlsx";
                                break;
                            }
                        case ExportType.PDF:
                            {
                                FileDialog.Filter = "PDF(*.pdf)|*.pdf";
                                break;
                            }
                        case ExportType.Image:
                            {
                                FileDialog.Filter = "JPEG（*.jpg）|*.jpg|TIFF（*.tif;*.tiff）|*.tif";
                                break;
                            }
                        case ExportType.CSV:
                            {
                                FileDialog.Filter = "CSV(逗号分隔)|*.csv";
                                break;
                            }
                        case ExportType.RTF:
                            {
                                FileDialog.Filter = "RTF|*.rtf";
                                break;
                            }
                    }
                    FileDialog.RestoreDirectory = true;
                    if (FileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string localFilePath = FileDialog.FileName.ToString(); //获得文件路径 
                        report.ExportDirect(GetExportType(exportType), localFilePath, false, false);
                    }
                    else
                        return false;
                }
            }
            catch (Exception ex)
            {

                return false;
            }

            return true;
        }






        public static bool Print(ReportArg arg, bool choicePrinterWithDefault = false)
        {
            var printerGrfName = arg.GrfName;
            //if (!string.IsNullOrEmpty(arg.PrinterGrfName))
            //    printerGrfName = arg.PrinterGrfName;
            //else if (printerGrfName.Contains("/"))
            //    printerGrfName = printerGrfName.Substring(printerGrfName.LastIndexOf('/') + 1);
            return printAndExport(arg, true);
        }

        public static bool Export(ReportArg arg)
        {
            return printAndExport(arg, false);
        }

        /// <summary>
        /// 加载参数数据
        /// </summary>
        /// <param name="report"></param>
        /// <param name="paramData"></param>
        private static void loadParameter(GridppReport report, JToken paramData)
        {
            if (paramData != null)
                foreach (var parameter in paramData.ToObject<Dictionary<string, object>>())
                {
                    //字符串类型的参数
                    var reportParameter = report.ParameterByName(parameter.Key);
                    if (reportParameter != null)
                    {
                        if (parameter.Value == null)
                            reportParameter.AsString = "";
                        else if (parameter.Value.GetType().IsPrimitive || parameter.Value is string)
                            reportParameter.AsString = parameter.Value.ToString();
                    }
                    else
                    {
                        //图片类型的参数
                        var picCtrl = report.ControlByName($"{CONST_REPORT_PICTURE_PREFIX}{parameter.Key}");
                        if (picCtrl != null && parameter.Value is string && parameter.Value.ToString().StartsWith(CONST_DATA_PICTURE_PREFIX))
                        {
                            string base64 = parameter.Value.ToString().Substring(CONST_DATA_PICTURE_PREFIX.Length);
                            byte[] bytes = Convert.FromBase64String(base64);
                            MemoryStream memStream = new MemoryStream(bytes);
                            BinaryReader br = new BinaryReader(memStream);
                            byte[] buffer = new byte[memStream.Length];
                            br.Read(buffer, 0, (int)memStream.Length);
                            picCtrl.AsPictureBox.LoadFromMemory(ref buffer[0], buffer.Length);
                        }
                    }
                }
        }

        /// <summary>
        /// 加载表格数据
        /// </summary>
        /// <param name="report"></param>
        /// <param name="tableData"></param>
        private static void loadTable(GridppReport report, JToken tableData)
        {
            if (tableData != null && tableData is JArray)
            {
                JArray reportRows = new JArray();
                foreach (var row in tableData as JArray)
                    reportRows.Add(row);

                JObject reportTable = new JObject();
                JObject reportRoot = new JObject();
                reportRoot[CONST_REPORT_ROW] = reportRows;
                reportTable[CONST_REPORT_ROOT] = reportRoot;

                //将对象数据序列化为xml后加载
                report.LoadDataFromXML(JsonConvert.DeserializeXmlNode(reportTable.ToString()).InnerXml);
            }
        }


        #region "获取报表相关属性值"
        /// <summary>
        /// 获取报表总行数
        /// </summary>
        /// <returns>返回总行数</returns>
        public int GetRowCount(GridppReport report)
        {
            return report.DetailGrid.Recordset.RecordCount;
        }

        /// <summary>
        /// 获取相应列宽
        /// </summary>
        /// <param name="ColumnName">列名称></param>
        /// <returns></returns>
        public double GetColumnWidth(GridppReport report, string ColumnName)
        {
            IGRColumn col = report.ColumnByName(ColumnName);
            return col.Width;
        }

        /// <summary>
        /// 获取参数值
        /// </summary>
        /// <param name="ParaName">参数名称</param>
        /// <returns></returns>
        public string GetParameter(GridppReport report, string ParaName)
        {
            string ParaValue = "";
            if (report != null)
            {
                try
                {
                    ParaValue = report.ParameterByName(ParaName).Value.ToString();
                }
                catch (Exception)
                {
                }
            }
            return ParaValue;

        }

        /// <summary>
        /// 获取报表信息中设置的每页明细行
        /// </summary>
        /// <returns>行数</returns>
        public int GetRowsPerPage(GridppReport report)
        {
            return int.Parse(report.DetailGrid.ColumnContent.RowsPerPage.ToString());
        }

        /// <summary>
        /// 打印机是否在线状态
        /// </summary>
        /// <returns>true：在线，false：不在线</returns>
        public bool PrinterIsOnLine(GridppReport report)
        {
            return report.Printer.Online;
        }

        #endregion

    }
}
