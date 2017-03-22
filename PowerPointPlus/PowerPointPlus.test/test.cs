using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using A= DocumentFormat.OpenXml.Drawing;

namespace PowerPointPlus.test
{
    [TestFixture]
    public class Test
    {
        [Test]
        public void TestPPT()
        {
            var ppt = PPTPlus.CreatePPT(@"table.pptx",@"test.pptx");
            uint currentId = uint.MaxValue;
            var slidePart = PPTPlus.InsertNewPage(ppt, 1, out currentId);
            slidePart.CreateText(new TextData()
            {
                FontName = "微软雅黑",
                FontColor = "FFFFFF",
                FontSize = 20,
                BackColor = "3CB371",
                Bold = false,
                Italic = false,
                TextValue = "罗唐坤爱李志凤",
                TextAlign = A.TextAlignmentTypeValues.Left,
                TextLocation = A.TextAnchoringTypeValues.Center,
                RectArea = new Rect(812800L, 812799L, 2380343L, 961665L)
            });

            Table table = new Table();
            table.RectArea = new Rect(2032000L, 719666L, 7431314L, 1471991L);
            table.ColWidths = new List<long>() {2249714L, 2249714L};
            table.RowHeader = new Row()
            {
                Height = 370840L,
                RowData = new List<Cell>()
                {
                    new Cell()
                    {
                        TextValue = "姓名",FontName = "微软雅黑",FontSize = 20,TextAlign = A.TextAlignmentTypeValues.Left,TextLocation = A.TextAnchoringTypeValues.Center,Bold=false,Italic=false
                    },
                    new Cell() {TextValue = "年龄",FontName = "微软雅黑",FontSize = 20,TextAlign = A.TextAlignmentTypeValues.Left,TextLocation = A.TextAnchoringTypeValues.Center,Bold=false,Italic=false}
                }
            };
            table.RowData = new List<Row>()
            {
                new Row()
                {
                    Height = 370840L,
                    RowData = new List<Cell>()
                    {
                        new Cell() {TextValue = "罗唐坤",FontName = "微软雅黑",FontSize = 20,TextAlign = A.TextAlignmentTypeValues.Left,TextLocation = A.TextAnchoringTypeValues.Center},
                        new Cell() {TextValue = "30",FontName = "微软雅黑",FontSize = 20,TextAlign = A.TextAlignmentTypeValues.Left,TextLocation = A.TextAnchoringTypeValues.Center}
                    }
                }
            };
            slidePart.CreateTable(table);

        PPTPlus.Save(ppt);

        }
    }
}
