using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointPlus
{
    public class Table
    {
        public Rect RectArea { get; set; }
        public Row RowHeader { get; set; }
        public List<Row> RowData { get; set; }

        public List<long> ColWidths { get; set; }
    }

    public class Row
    {
        public List<Cell> RowData { get; set; }
        public long Height { get; set; }
    }

    public class Cell:TextData
    {
        public long Height { get; set; }
        public long Width { get; set; }

        /// <summary>
        /// 单元格合并(行)
        /// </summary>
        public int RowSpan { get; set; }
    }
}
