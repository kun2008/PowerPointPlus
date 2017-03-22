using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointPlus
{
    /// <summary>
    /// 矩形，表示对象的区域和大小
    /// </summary>
    public struct Rect
    {
        public long X
        {get; set;}
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }

        public Rect(long x, long y, long width, long height)
        {
            this.X = x;
            this.Y = y;
            this.Width = width;
            this.Height = height;
        }
    }

    public enum TextLocation
    {
        Top,
        Center,
        Bottom
    }

    public enum TextAlign
    {
        Left,
        Center,
        Right,
    }
    /// <summary>
    /// 字体
    /// </summary>
    public class FontInfo
    {
        /// <summary>
        /// 字体名称
        /// </summary>
        public string FontName { get; set; }
        /// <summary>
        /// 字体大小
        /// </summary>
        public double FontSize { get; set; }
        /// <summary>
        /// 字体前景色(RGB编码:如FFFFFF)
        /// </summary>
        public string FontColor { get; set; }
        /// <summary>
        /// 字体背景色(RGB编码:如FFFFFF)
        /// </summary>
        public string BackColor { get; set; }
        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool Bold { get; set; }
        /// <summary>
        /// 是否斜体
        /// </summary>
        public bool Italic { get; set; }

        /// <summary>
        /// 文本对齐:Left、Right、Center
        /// </summary>
        public TextAlign TextAlign { get; set; }
        /// <summary>
        /// 文本位置:Top、Center、Bottom
        /// </summary>
        public TextLocation TextLocation { get; set; }
    }
    /// <summary>
    /// 文字信息
    /// </summary>
    public class TextData: FontInfo
    {
        public Rect RectArea { get; set; }
        public string TextValue { get; set; }

    }
}
