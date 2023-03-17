using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GridReport.Common
{
    public enum PaperSize
    {
        LETTER = 1,
        A3 = 8,
        A4 = 9,
        A5 = 11,
        B4 = 12,
        B5 = 13,
        自定义 = 256
    }

    /// <summary>
    /// 纸张方向
    /// </summary>
    public enum PaperOrientation
    {
        默认 = 0,
        纵向 = 1,
        横向 = 2
    }

    public class ReportArg
    {
        /// <summary>
        /// 模板名称
        /// </summary>
        public string GrfName { get; set; }

        /// <summary>
        /// 打印机模板名称
        /// </summary>
        public string PrinterGrfName { get; set; }

        /// <summary>
        /// 数据序列化字符串
        /// </summary>
        public object Data { get; set; }

        /// <summary>
        /// 是否显示打印预览
        /// </summary>
        public bool ShowPreview { get; set; } = false;

        /// <summary>
        /// 打印设置，传空值取系统默认打印机
        /// </summary>
        public string Printer { get; set; }

        /// <summary>
        /// 纸张大小
        /// </summary>
        public PaperSize paperSize { get; set; }

        /// <summary>
        /// 打印纸张方向
        /// </summary>
        public PaperOrientation Poaoero { get; set; }
        /// <summary>
        /// 宽度
        /// </summary>
        public double PaperWidth { get; set; }
        /// <summary>
        /// 高度
        /// </summary>
        public double PaperLength { get; set; }
        /// <summary>
        /// 边距 左
        /// </summary>
        public double LeftMargin { get; set; }
        /// <summary>
        /// 边距 右
        /// </summary>
        public double RightMargin { get; set; }
        /// <summary>
        /// 边距 上
        /// </summary>
        public double TopMargin { get; set; }
        /// <summary>
        /// 边距 下
        /// </summary>
        public double BottomMargin { get; set; }
    }
}
