using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace PowerPointPlus.test
{
    [TestFixture]
    public class Test
    {
        [Test]
        public void TestPPT()
        {
           var ppt= PPTPlus.CreatePPT(@"C:\Users\Administrator\Desktop\ppt\table.pptx",
                @"C:\Users\Administrator\Desktop\ppt\test.pptx");
            uint currentId = uint.MaxValue;
            var slidePart= PPTPlus.InsertNewPage(ppt, 1, out currentId);
            PPTPlus.CreateText(slidePart, new TextData()
            {
                FontName = "微软雅黑",
                FontColor = "FFFFFF",
                FontSize = 20,
                BackColor = "3CB371",
                Bold = false,
                Italic = false,
                TextValue = "罗唐坤爱李志凤",
                TextAlign = TextAlign.Left,
                TextLocation = TextLocation.Center,
                RectArea = new Rect(812800L, 812799L, 2380343L, 961665L)
            });
            PPTPlus.Save(ppt);

        }
    }
}
