/**
* 命名空间: PowerPointPlus
*
* 功 能： PPT操作类，复制模版新建PPT文件，
* 支持新增PPT页、创建文本、表格、折线图、柱状图、饼图、3D饼图、插入图片、附件
* 目前仅支持office 2013
* 类 名： PPTPlus.cs
*
* mail:415895442@qq.com
* Ver 变更日期 负责人 变更内容
* ───────────────────────────────────
* V1.0 2017/3/18 天地有忧 
*
* 引用修改请注明出处和作者
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace PowerPointPlus
{
    public static class PPTPlus
    {
        /// <summary>
        /// ID计数器，防止ID重复
        /// </summary>
        private static uint _idCounter = 0;
        /// <summary>
        /// 拷贝模板，新建一个PPT文件
        /// </summary>
        /// <param name="templateFile">目标文件</param>
        /// <param name="pptFile">ppt文件</param>
        /// <returns></returns>
        public static PresentationDocument CreatePPT(string templateFile, string pptFile)
        {
            if (!File.Exists(templateFile))
            {
                return null;
            }
            try
            {
                File.Copy(templateFile, pptFile, true);
                var document = PresentationDocument.Open(pptFile, true);
                return document;
            }
            catch
            {
                return null;
            }
        }


        public static bool CreateTable(this SlidePart slidePart, Table table)
        {
            if (slidePart == null || table == null)
            {
                return false;
            }
            _idCounter++;
            // Declare and instantiate the graphic Frame of the new slide
            GraphicFrame graphicFrame = slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(new GraphicFrame());

            // Specify the required Frame properties of the graphicFrame
            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension = new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" };
            P14.ModificationId modificationId1 = new P14.ModificationId() { Val = 3229994563U };
            modificationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
            applicationNonVisualDrawingPropertiesExtension.Append(modificationId1);
            graphicFrame.NonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties
            (new NonVisualDrawingProperties() { Id = _idCounter, Name = "table "+ _idCounter },
            new NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(new ApplicationNonVisualDrawingPropertiesExtensionList(applicationNonVisualDrawingPropertiesExtension)));

            //table的位置，大小
            graphicFrame.Transform = new Transform(new A.Offset()
            {
                X = table.RectArea.X, Y = table.RectArea.Y
            }, new A.Extents()
            {
                Cx = table.RectArea.Width, Cy = table.RectArea.Height
            });

            // Specify the Griaphic of the graphic Frame
            graphicFrame.Graphic = new A.Graphic(new A.GraphicData(CreateTable(table)) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" });
            return true;
        }
        private static A.Table CreateTable(Table data)
        {
            
            A.Table table = new A.Table();

            // Specify the required table properties for the table
            A.TableProperties tableProperties = new A.TableProperties() { FirstRow = true, BandRow = true };
            A.TableStyleId tableStyleId = new A.TableStyleId();
            tableStyleId.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

            tableProperties.Append(tableStyleId);

            // Declare and instantiate tablegrid and colums
            A.TableGrid tableGrid1 = new A.TableGrid();
            foreach (var item in data.ColWidths)
            {
                tableGrid1.Append(new A.GridColumn()
                {
                    Width = item
                });
            }
            table.Append(tableProperties);
            table.Append(tableGrid1);
            A.TableRow rowHeader = new A.TableRow() { Height = data.RowHeader.Height };
            foreach (var item in data.RowHeader.RowData)
            {

                rowHeader.Append(CreateTextCell(item));
                
            }
            table.Append(rowHeader);
            foreach (var item in data.RowData)
            {
                A.TableRow row = new A.TableRow() { Height = data.RowHeader.Height };
                foreach (var cell in item.RowData)
                {
                    row.Append(CreateTextCell(cell));
                }
                table.Append(row);
            }
            return table;
        }

        private static A.TableCell CreateTextCell(Cell cell)
        {

            A.TableCell tableCell4 = new A.TableCell();

            A.TextBody textBody4 = new A.TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Alignment = cell.TextAlign};

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "zh-CN", AlternativeLanguage = "en-US", FontSize = (Int32Value)cell.FontSize*100, Bold = cell.Bold, Italic = cell.Italic, Dirty = false };
            runProperties2.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));

            if (!string.IsNullOrEmpty(cell.BackColor))
            {
                A.SolidFill solidFill10 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex()
                {
                    Val =cell.BackColor
                };

                solidFill10.Append(rgbColorModelHex1);
                runProperties2.Append(solidFill10);
            }
  
            A.LatinFont latinFont10 = new A.LatinFont() { Typeface = cell.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = cell.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };

            
            runProperties2.Append(latinFont10);
            runProperties2.Append(eastAsianFont10);
            A.Text text2 = new A.Text();
            text2.Text = cell.TextValue;

            run2.Append(runProperties2);
            run2.Append(text2);

            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "zh-CN", AlternativeLanguage = "en-US", FontSize = 2000, Bold = true, Italic = true, Dirty = false };

            if (!string.IsNullOrEmpty(cell.BackColor))
            {
                A.SolidFill solidFill11 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex()
                {
                    Val = cell.BackColor
                };

                solidFill11.Append(rgbColorModelHex2);
                endParagraphRunProperties4.Append(solidFill11);
            }

            A.LatinFont latinFont11 = new A.LatinFont() { Typeface = cell.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = cell.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };

            
            endParagraphRunProperties4.Append(latinFont11);
            endParagraphRunProperties4.Append(eastAsianFont11);

            paragraph4.Append(paragraphProperties1);
            paragraph4.Append(run2);
            paragraph4.Append(endParagraphRunProperties4);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);

            A.TableCellProperties tableCellProperties4 = new A.TableCellProperties() { Anchor = cell.TextLocation };

            
            if (!string.IsNullOrEmpty(cell.FontColor))
            {
                A.SolidFill solidFill12 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex()
                {
                    Val = cell.FontColor
                };
                solidFill12.Append(rgbColorModelHex3);
                tableCellProperties4.Append(solidFill12);
            }
            tableCell4.Append(textBody4);
            tableCell4.Append(tableCellProperties4);
            return tableCell4;

        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="presentationDocument"></param>
        /// <returns></returns>
        public static bool Save(PresentationDocument presentationDocument)
        {
            if (presentationDocument == null)
            {
                return false;
            }
            try
            {
                presentationDocument.PresentationPart.Presentation.Save();
                presentationDocument.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 创建标题，标题会使用模板中的样式
        /// </summary>
        /// <param name="slidePart"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        public static bool CreateTitle(SlidePart slidePart, string title)
        {
            if (string.IsNullOrEmpty(title))
            {
                return false;
            }
            Slide slide = slidePart.Slide;

            Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
            titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
                (new NonVisualDrawingProperties()
                {
                    Id = _idCounter, Name = "Title"+ _idCounter

                },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title }));
            titleShape.ShapeProperties = new ShapeProperties();

            // Specify the text of the title shape.
            titleShape.TextBody = new TextBody(new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text() { Text = title })));

            slide.Save(slidePart);
            return true;
        }
        /// <summary>
        /// 创建文字
        /// </summary>
        /// <param name="slidePart"></param>
        /// <param name="textData"></param>
        /// <returns></returns>
        public static bool CreateText(this SlidePart slidePart,TextData textData)
        {
            if (slidePart == null || textData == null)
            {
                return false;
            }
            _idCounter++;
            Slide slide = slidePart.Slide;
            Shape shape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties()
            {
                Id = (UInt32Value)(_idCounter), Name = "文本框 "+_idCounter
            };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties() { TextBox = true };
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = textData.RectArea.X, Y = textData.RectArea.Y };
            A.Extents extents2 = new A.Extents()
            {
                Cx = textData.RectArea.Width, Cy = textData.RectArea.Height
            };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill10 = new A.SolidFill();
            //背景色
            if (!string.IsNullOrEmpty(textData.BackColor))
            {
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex()
                {
                    Val = textData.BackColor
                };
                solidFill10.Append(rgbColorModelHex3);
            }


            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill10);

            TextBody textBody1 = new TextBody();

            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square,RightToLeftColumns = false, Anchor = textData.TextLocation };

            A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

            bodyProperties1.Append(shapeAutoFit1);

            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties(){ Alignment = textData.TextAlign };

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "zh-CN", AlternativeLanguage = "en-US", FontSize = (int)Math.Round(textData.FontSize,2)*100,
                Bold = textData.Bold, Italic = textData.Italic, Dirty = false };
            runProperties1.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));

            A.SolidFill solidFill11 = new A.SolidFill();
            if (!string.IsNullOrEmpty(textData.FontColor))
            {
                A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex()
                {
                    Val = textData.FontColor
                };
                solidFill11.Append(rgbColorModelHex4);
            }

            A.LatinFont latinFont10 = new A.LatinFont() { Typeface = textData.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = textData.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };

            runProperties1.Append(solidFill11);
            runProperties1.Append(latinFont10);
            runProperties1.Append(eastAsianFont10);
            A.Text text1 = new A.Text();
            text1.Text = textData.TextValue;

            run1.Append(runProperties1);
            run1.Append(text1);

            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "zh-CN", AlternativeLanguage = "en-US", FontSize = 2400, Bold = true, Italic = true, Dirty = false };

            A.SolidFill solidFill12 = new A.SolidFill();
            if (!string.IsNullOrEmpty(textData.FontColor))
            {
                A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex()
                {
                    Val = textData.FontColor
                };

                solidFill12.Append(rgbColorModelHex5);
            }
       
            A.LatinFont latinFont11 = new A.LatinFont() { Typeface = textData.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = textData.FontName, Panose = "020B0503020204020204", PitchFamily = 34, CharacterSet = -122 };

            endParagraphRunProperties1.Append(solidFill12);
            endParagraphRunProperties1.Append(latinFont11);
            endParagraphRunProperties1.Append(eastAsianFont11);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape.Append(nonVisualShapeProperties1);
            shape.Append(shapeProperties1);
            shape.Append(textBody1);

            return true;
        }

        /// <summary>
        /// 在指定位置处新建ppt页
        /// </summary>
        /// <param name="presentationDocument">ppt对象</param>
        /// <param name="position">位置</param>
        /// <param name="currentId">新建页的id</param>
        /// <param name="preSlideId">上一页的id，默认为uint.max，设置之后position会失效</param>
        /// <returns></returns>
        public static SlidePart InsertNewPage(PresentationDocument presentationDocument, int position, out uint currentId , uint preSlideId=uint.MaxValue)
        {
            currentId = uint.MaxValue;
            if (presentationDocument == null)
            {
                return null;
            }
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            //新建SlidePart对象
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));

            // Construct the slide content.            
            // Specify the non-visual properties of the new slide.
            NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

            // Specify the group shape properties of the new slide.
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());
            slide.Save(slidePart);
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
            var maxId = slideIdList.ChildElements.Max(r => ((SlideId) r).Id) + 1;
            SlideId prevSlideId = null;

            if (preSlideId == uint.MaxValue)
            {
                foreach (SlideId slideId in slideIdList.ChildElements)
                {
                    position--;
                    if (position == 0)
                    {
                        prevSlideId = slideId;
                        break;
                    }
                }
            }
            else
            {
                prevSlideId = (SlideId)slideIdList.ChildElements.First(r => ((SlideId) r).Id == preSlideId);
            }
            SlidePart lastSlidePart;
            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart) presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }
            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
            currentId = maxId;
            return slidePart;

        }
    }
}
