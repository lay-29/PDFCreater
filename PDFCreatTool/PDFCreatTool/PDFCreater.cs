using DH.Util.Gasbox;
using Microsoft.SqlServer.Server;
using PdfSharp.Charting;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Filters;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace Utilities.PDFCreater
{


    public class PDFCreater : IPDFCreater
    {
        #region Const
        private const int DEFAULT_MARGIN_TOP_OR_BOTTOM = 32;
        private const int DEFAULT_MARGIN_LEFT_OR_RIGHT = 40;
        #endregion Const

        #region Member Data
        private FontFamily _STFangSo;
        private DHPDFContent _dHPDFContent;

        private PdfDocument _pdfDocument;
        private XFont _defualtWatermarkFont;
        private XFont _defualtTittleFont;

        private XFont _defualtContentFont;
        private XFont _defualtCommentFont;

        private XFont _defualtOneLevelHeading;  //默认一级标题字体
        private XFont _defualtTwoLevelHeading;
        private XFont _defualtThreeLevelHeading;
        private XFont _defualtFourLevelHeading;    //默认四级标题字体

        private double _currentXPostion;
        private double _currentYPostion;
        private PdfPage _currentPage;
        #endregion  Member Data

        #region Life Cycle
        public PDFCreater()
        {
            _pdfDocument = new PdfDocument();
            _currentPage = _pdfDocument.AddPage();

            FindFont();
            _dHPDFContent = new DHPDFContent();


            _currentYPostion = MarginTop;
        }
        #endregion   Life Cycle

        #region   Methods
        public void SetMargin(int? left = null, int? top = null, int? right = null, int? bottom = null)
        {
            if (left != null)
                MarginLeft = (int)left;
            if (top != null)
                MarginTop = (int)top;
            if (right != null)
                MarginRight = (int)right;
            if (bottom != null)
                MarginBottom = (int)bottom;
        }

        public void SetTittle(string tittle, string tittleContent = "")
        {
            try
            {
                if (_dHPDFContent.Content.Count != 0)
                {
                    if (_dHPDFContent.Content[0] is OneLevelContent levelContent)
                    {
                        if (levelContent.LevelIndex == 0)
                        {
                            levelContent.LevelTittle = tittle;
                            levelContent.LevelContent = tittleContent;
                            levelContent.ContentLineSpacing = 10;
                            return;
                        }
                    }
                }

                _dHPDFContent.Content.Add(new OneLevelContent(0, tittle, tittleContent));

            }
            catch (Exception em)
            {

                throw;
            }
        }

        public void AddContent(PDFContentBase content)
        {
            try
            {
                _dHPDFContent.AddContent(content);
            }
            catch (Exception em)
            {

                throw;
            }

        }



        #region AddWatermark
        public void AddWatermark(string text = "德鸿半导体设备（浙江）有限公司")
        {
            StringContent watermark = new StringContent(text);
            watermark.FontSize = 20;
            watermark.Color = DHColor.Red;
            _dHPDFContent.Watermark = watermark;
        }
        public void AddWatermark(StringContent watermark)
        {
            _dHPDFContent.Watermark = watermark;
        }

        #endregion AddWatermark

        public void AddHeader(StringContent header)
        {
            try
            {
                _dHPDFContent.Header = header;
            }
            catch (Exception em)
            {

                throw;
            }
        }
        public void AddHeader(string header = "德鸿半导体设备（浙江）有限公司")
        {
            try
            {
                StringContent headers = new StringContent(header);
                headers.FontSize = 8;
                headers.Color = DHColor.BLack;
                headers.HorizontalAlignment = Alignment.Center;
                _dHPDFContent.Header = headers;
            }
            catch (Exception em)
            {

                throw;
            }
        }
        public void AddFooter(StringContent footer)
        {

            try
            {
                _dHPDFContent.Footer = footer;
            }
            catch (Exception em)
            {

                throw;
            }
        }
        public void AddFooter(string footer = "德鸿半导体设备（浙江）有限公司")
        {
            try
            {

                StringContent footers = new StringContent(footer);
                footers.FontSize = 8;
                footers.Color = DHColor.BLack;
                footers.HorizontalAlignment = Alignment.Center;
                _dHPDFContent.Footer = footers;
            }
            catch (Exception em)
            {

                throw;
            }
        }
        public void SetPageNumber(Alignment horizontalAlignment, Alignment verticalAlignment)
        {
            _dHPDFContent.PageNumberAlignment = (horizontalAlignment, verticalAlignment);
        }

        public async void Save(string path, bool openAfterSave = false, bool isSync = true)
        {
            try
            {
                Task<Exception> task = Task.Run(() =>
                {
                    try
                    {

                        AddContentHelper(_dHPDFContent);
                        if (_dHPDFContent.Watermark != null)
                        {
                            AddWatermarkHelper(_dHPDFContent.Watermark);
                        }
                        if (_dHPDFContent.Header != null)
                        {
                            AddHeaterHelper(_dHPDFContent.Header);
                        }
                        if (_dHPDFContent.Footer != null)
                        {
                            AddFooterkHelper(_dHPDFContent.Footer);
                        }
                        if (_dHPDFContent.PageNumberAlignment != null)
                        {
                            Alignment horizontalAlignment = (Alignment)_dHPDFContent.PageNumberAlignment.Value.Item1;
                            Alignment verticalAlignment = (Alignment)_dHPDFContent.PageNumberAlignment.Value.Item2;
                            AddPageNumberHelper(horizontalAlignment, verticalAlignment);
                        }
                        CreatDirectory(path);
                        _pdfDocument.Save(path);
                        return null;
                    }
                    catch (Exception em)
                    {
                        return em;
                    }
                });

                if (isSync)
                {
                    task.Wait();
                }
                else
                {
                    await task;
                }
                if (task.Result != null)
                {
                    throw task.Result;
                }
                if (openAfterSave == true)
                {
                    Process.Start(path);
                }
            }
            catch (Exception em)
            {

                throw;
            }
        }
        #endregion  Methods

        #region  Helper:用于写入信息在PDF文档中
        private void AddContentHelper(DHPDFContent content)
        {
            try
            {
                foreach (var oneContentBase in content.Content)
                {
                    AddContentHelper(oneContentBase);
                }
            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddContentHelper(PDFContentBase content)
        {
            try
            {
                switch (content.ContentType)
                {
                    case ContentType.LevelText:
                        if (content is OneLevelContent levelContent)
                        {
                            AddContentHelper(levelContent);
                        }
                        else
                        {
                            throw new Exception($"ContentType is {content.ContentType},But is not OneLevelContent Object.");
                        }

                        break;
                    case ContentType.String:
                        if (content is StringContent stringContent)
                        {
                            AddContentHelper(stringContent);
                        }
                        else
                        {
                            throw new Exception($"ContentType is {content.ContentType},But is not StringContent Object.");

                        }

                        break;
                    case ContentType.Table:
                        if (content is Table table)
                        {
                            AddContentHelper(table);
                        }
                        else
                        {
                            throw new Exception($"ContentType is {content.ContentType},But is not Table Object.");

                        }

                        break;
                    case ContentType.Image:
                        if (content is Image image)
                        {
                            AddContentHelper(image);
                        }
                        else
                        {
                            throw new Exception($"ContentType is {content.ContentType},But is not Table Object.");

                        }

                        break;
                    case ContentType.Line:
                        if (content is Line line)
                        {
                            AddContentHelper(line);
                        }
                        else
                        {
                            throw new Exception($"ContentType is {content.ContentType},But is not Line Object.");

                        }
                        break;
                    default:
                        break;

                }

            }
            catch (Exception em)
            {

                throw em;
            }

        }
        private void AddContentHelper(OneLevelContent oneLevelContent)
        {
            try
            {
                XFont tittleFont = _defualtOneLevelHeading;
                XFont contentFont = _defualtContentFont;
                XBrush xBrushTittle = XBrushes.Black;
                XBrush xBrushContent = XBrushes.Black;
                if (oneLevelContent.TittleFont != null)
                {

                    tittleFont = FindXFont(oneLevelContent.TittleFont);
                    xBrushTittle = FindBruse(oneLevelContent.TittleFont.Color);
                }
                switch (oneLevelContent.LevelIndex)
                {

                    case 0:
                        if (oneLevelContent.TittleFont == null)
                        {
                            tittleFont = _defualtTittleFont;
                            xBrushTittle = XBrushes.Red;
                        }
                        break;
                    case 1:
                        if (oneLevelContent.TittleFont == null)
                        {
                            tittleFont = _defualtOneLevelHeading;
                        }
                        break;
                    case 2:
                        if (oneLevelContent.TittleFont == null)
                        {
                            tittleFont = _defualtTwoLevelHeading;
                        }
                        break;
                    case 3:
                        if (oneLevelContent.TittleFont == null)
                        {
                            tittleFont = _defualtThreeLevelHeading;
                        }
                        break;
                    case 4:
                        if (oneLevelContent.TittleFont == null)
                        {
                            tittleFont = _defualtFourLevelHeading;
                        }
                        break;


                    default:
                        throw new Exception($"Don't OneLevelContent LevelIndex:{oneLevelContent.LevelIndex}");

                }

                if (oneLevelContent.ContentFont == null)
                {
                    contentFont = _defualtContentFont;
                }
                else
                {
                    contentFont = FindXFont(oneLevelContent.ContentFont);
                    xBrushContent = FindBruse(oneLevelContent.ContentFont.Color);
                }


                //同一级标题及内容
                if (oneLevelContent.LevelIndex != 0)
                {
                    XStringFormat format = new XStringFormat();
                    format.Alignment = XStringAlignment.Near;
                    format.LineAlignment = XLineAlignment.Near;


                    if (oneLevelContent.LevelTittle != null && oneLevelContent.LevelTittle != string.Empty && oneLevelContent.LevelTittle != "")
                    {
                        AddString_AutoLineFeed(oneLevelContent.LevelTittle, tittleFont, xBrushTittle, startXPos: MarginLeft + oneLevelContent.TittleRetractSpacing, endXPos: (int)_currentPage.Width.Value - MarginRight, oneLevelContent.TittleLineSpacing, xStringFormat: format, ref _currentYPostion, ref _currentPage);

                    }
                    if (oneLevelContent.LevelContent != null && oneLevelContent.LevelContent != string.Empty && oneLevelContent.LevelContent != "")
                    {
                        AddString_AutoLineFeed(oneLevelContent.LevelContent, contentFont, xBrushContent, startXPos: MarginLeft + oneLevelContent.ContentRetractSpacing, endXPos: (int)_currentPage.Width.Value - MarginRight, oneLevelContent.ContentLineSpacing, xStringFormat: format, ref _currentYPostion, ref _currentPage);

                    }
                }
                else  //oneLevelContent.LevelIndex == 0 PDF 标题
                {
                    XStringFormat format = new XStringFormat();
                    format.Alignment = XStringAlignment.Center;
                    format.LineAlignment = XLineAlignment.Near;

                    XStringFormat format2 = new XStringFormat();
                    format2.Alignment = XStringAlignment.Near;
                    format2.LineAlignment = XLineAlignment.Near;

                    //找到图标的路径
                    string currentPath = System.Environment.CurrentDirectory;
                    int rootIndex = currentPath.IndexOf("DHSOFTWARE");
                    string rootPath = string.Empty;
                    if (rootIndex != -1)
                    {
                        rootPath = currentPath.Remove(rootIndex);

                    }
                    else
                    {

                        throw new Exception("未找到“DHSOFTWARE”根路径");
                    }

                    string ImagePath = rootPath + "DHSOFTWARE\\Utilities\\Utilities\\PDFCreater\\ImageFile\\DHBDT.jpg";

                    Image image = new Image(ImagePath);
                    image.Width = 80;
                    image.Height = 30;
                    AddImage(8, 25, image, _currentPage);
                    AddString_AutoLineFeed(oneLevelContent.LevelTittle, tittleFont, xBrushTittle, startXPos: 0, endXPos: (int)_currentPage.Width.Value, oneLevelContent.TittleLineSpacing, xStringFormat: format, ref _currentYPostion, ref _currentPage);
                    AddLine(_currentYPostion - 2, MarginLeft, _currentPage.Width.Value - MarginRight, _currentPage, width: 1, color: DHColor.Red);
                    AddString_AutoLineFeed(oneLevelContent.LevelContent, contentFont, xBrushContent, startXPos: MarginLeft + oneLevelContent.ContentRetractSpacing, endXPos: (int)_currentPage.Width.Value - MarginRight, oneLevelContent.ContentLineSpacing, xStringFormat: format2, ref _currentYPostion, ref _currentPage);

                }
                //内容
                foreach (var subContent in oneLevelContent.SubLevelContent)
                {
                    AddContentHelper(subContent);
                }

            }
            catch (Exception em)
            {

                throw em;
            }
        }
        private void AddContentHelper(StringContent stringContent)
        {
            try
            {

                XGraphics xGraphics = XGraphics.FromPdfPage(_currentPage);


                XBrush brush = FindBruse(stringContent.Color);
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                XFont font = new XFont(_STFangSo, stringContent.FontSize, XFontStyle.Regular, options);

                //判断是否需要添加页面
                XSize lineSize = xGraphics.MeasureString(stringContent.Content, font);
                if (_currentYPostion + lineSize.Height > _currentPage.Height - MarginBottom)
                {
                    AddPage(ref _currentYPostion, ref _currentPage);
                    xGraphics.Dispose();
                    xGraphics = XGraphics.FromPdfPage(_currentPage);
                }


                XStringFormat format = new XStringFormat();
                if (stringContent.HorizontalAlignment == Alignment.Center)
                {

                    format.Alignment = XStringAlignment.Center;
                    format.LineAlignment = XLineAlignment.Near;
                    xGraphics.DrawString(stringContent.Content, font, brush, new XRect(0, _currentYPostion, _currentPage.Width, _currentPage.Height), format);
                }
                else if (stringContent.HorizontalAlignment == Alignment.Near)
                {
                    format.Alignment = XStringAlignment.Near;
                    format.LineAlignment = XLineAlignment.Near;
                    xGraphics.DrawString(stringContent.Content, font, brush, new XRect(MarginLeft + stringContent.RetractSpacing, _currentYPostion, _currentPage.Width, _currentPage.Height), format);
                }
                else if (stringContent.HorizontalAlignment == Alignment.Far)
                {
                    format.Alignment = XStringAlignment.Far;
                    format.LineAlignment = XLineAlignment.Near;
                    xGraphics.DrawString(stringContent.Content, font, brush, new XRect(0, _currentYPostion, _currentPage.Width - MarginRight - stringContent.RetractSpacing, _currentPage.Height), format);
                }
                _currentYPostion += (int)stringContent.FontSize + stringContent.LineSpacing;
                xGraphics.Dispose();


            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddContentHelper(Table table)
        {
            try
            {
                Dictionary<int, double> columnWidthMaxValue = new Dictionary<int, double>();
                Dictionary<int, double> rowHeightMaxValue = new Dictionary<int, double>();
                for (int i = 1; i <= table.ColumnNumber; i++)
                {
                    columnWidthMaxValue.Add(i, 0);
                }
                for (int i = 1; i <= table.RowNumber; i++)
                {
                    rowHeightMaxValue.Add(i, 0);
                }

                //找到每列最宽，每行最高
                foreach (var cell in table.TableContent)
                {
                    XGraphics xGraphics = XGraphics.FromPdfPage(_currentPage);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                    XFont xFont = new XFont(_STFangSo, cell.Value.CellText.FontSize, XFontStyle.Regular, options);

                    double measureValueWidth = xGraphics.MeasureString(cell.Value.CellText.Content, xFont).Width;
                    double measureValueHeight = xGraphics.MeasureString(cell.Value.CellText.Content, xFont).Height;
                    xGraphics.Dispose();
                    double maxWithdValue = columnWidthMaxValue[cell.Key.Item2];
                    double maxHeightValue = rowHeightMaxValue[cell.Key.Item1];
                    columnWidthMaxValue[cell.Key.Item2] = measureValueWidth > maxWithdValue ? measureValueWidth : maxWithdValue;
                    rowHeightMaxValue[cell.Key.Item1] = measureValueHeight > maxHeightValue ? measureValueHeight : maxHeightValue;
                }

                //设定cell宽高
                foreach (var cell in table.TableContent)
                {
                    cell.Value.Width = columnWidthMaxValue[cell.Key.Item2] + 10;
                    cell.Value.Height = rowHeightMaxValue[cell.Key.Item1] + 10;
                }

                //设定StartXY EndXY
                double X = 0;
                double Y = 0;
                foreach (var cell in table.TableContent)
                {

                    if (cell.Key.Item2 == 1) //第一列
                    {
                        X = 0;
                        if (cell.Key.Item1 != 1)
                        {
                            Y += table.TableContent[(cell.Key.Item1 - 1, 1)].Height;
                        }

                    }
                    cell.Value.StartX = X;
                    cell.Value.StartY = Y;
                    cell.Value.EndX = X + cell.Value.Width;
                    cell.Value.EndY = Y + cell.Value.Height;
                    X += cell.Value.Width;
                }


                //获取table 宽高
                double totalHeight = 0;
                double totalWidth = 0;

                foreach (var value in rowHeightMaxValue)
                {
                    totalHeight += value.Value + 10;
                }

                foreach (var value in columnWidthMaxValue)
                {
                    totalWidth += value.Value + 10;
                }
                table.Width = totalWidth;
                table.Height = totalHeight;

                //判断是否换页显示
                if (_currentYPostion + table.Height + table.Comment.FontSize > _currentPage.Height - MarginBottom)
                {
                    AddPage(ref _currentYPostion, ref _currentPage);
                }

                //开始写入table
                double startX = (_currentPage.Width - totalWidth) / 2;
                double startY = _currentYPostion;

                foreach (var cell in table.TableContent)
                {
                    AddTableCell(cell.Value, startX, startY, _currentPage);
                }
                _currentYPostion += (int)table.Height + 6;
                if (table.Comment != null && table.Comment.Content != "")
                {
                    AddContentHelper(table.Comment);
                }


            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddContentHelper(Image image)
        {
            try
            {
                if (XImage.ExistsFile(image.FilePath) == false)
                {
                    throw new Exception($"image.FilePath={image.FilePath}::Don't Exists!");
                }
                XImage xImage = XImage.FromFile(image.FilePath);

                MeasureImage(image, out double height, out double width);
                if (_currentYPostion + height > _currentPage.Height - MarginBottom)
                {
                    AddPage(ref _currentYPostion, ref _currentPage);
                }

                //绘图
                switch (image.HorizontalAlignment)
                {
                    case Alignment.Center:
                        AddImage((_currentPage.Width - width) / 2, _currentYPostion, image, _currentPage);
                        break;
                    case Alignment.Near:
                        AddImage(MarginLeft, _currentYPostion, image, _currentPage);
                        break;
                    case Alignment.Far:
                        AddImage(_currentPage.Width - MarginRight - width, _currentYPostion, image, _currentPage);
                        break;
                }
                _currentYPostion += height + 5;
                //绘图名
                MeasureString(image.Comment, out double strHeight, out double strWidth);
                if (image.Comment != null && image.Comment.Content != null && image.Comment.Content != "")
                {
                    switch (image.HorizontalAlignment)
                    {
                        case Alignment.Center:
                            AddString((_currentPage.Width - strWidth) / 2, _currentYPostion, image.Comment, _currentPage);
                            break;
                        case Alignment.Near:
                            AddString(MarginLeft + width / 2 - strWidth / 2, _currentYPostion, image.Comment, _currentPage);
                            break;
                        case Alignment.Far:
                            AddString(_currentPage.Width - MarginRight - width / 2 - strWidth / 2, _currentYPostion, image.Comment, _currentPage);
                            break;

                    }
                }
                _currentYPostion += strHeight + 5;

            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddContentHelper(Line line)
        {
            try
            {
                if (line.IsHorizontalLine)
                {
                    AddLine(_currentYPostion, line.StartPoint, line.EndPoint, _currentPage, line.LineWidth, line.Color, true);
                    _currentYPostion += line.LineWidth + 3;
                }
                else
                {
                    //AddLine(_currentYPostion, line.StartPoint, line.EndPoint, _currentPage, line.LineWidth, line.Color, true);
                }

            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddWatermarkHelper(StringContent watermark)
        {
            try
            {
                try
                {
                    foreach (var page in _pdfDocument.Pages)
                    {

                        XGraphics xGraphics = XGraphics.FromPdfPage(page);
                        xGraphics.TranslateTransform(page.Width / 2, page.Height / 2);
                        xGraphics.RotateTransform(-Math.Atan(page.Height / page.Width) * 180 / Math.PI);
                        xGraphics.TranslateTransform(-page.Width / 2, -page.Height / 2);

                        XBrush brush = FindWatermarkBruse(watermark.Color);

                        for (int i = 0; i < page.Height; i += 100)
                        {
                            int x = 0;
                            if (i / 100 % 2 == 0)
                            {
                                x = -200;
                            }
                            else
                            {
                                x = 200;
                            }


                            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                            XFont watermarkFont = new XFont(_STFangSo, watermark.FontSize, XFontStyle.Bold, options);

                            xGraphics.DrawString(watermark.Content, watermarkFont, brush, x, i, XStringFormat.TopLeft);
                        }
                        xGraphics.Dispose();
                    }

                }
                catch (Exception em)
                {

                    throw em;
                }
            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddHeaterHelper(StringContent heater)
        {
            try
            {
                foreach (var page in _pdfDocument.Pages)
                {

                    XGraphics xGraphics = XGraphics.FromPdfPage(page);


                    XBrush brush = FindBruse(heater.Color);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                    XFont font = new XFont(_STFangSo, heater.FontSize, XFontStyle.Regular, options);


                    XStringFormat format = new XStringFormat();
                    if (heater.HorizontalAlignment == Alignment.Center)
                    {

                        format.Alignment = XStringAlignment.Center;
                        format.LineAlignment = XLineAlignment.Near;
                        xGraphics.DrawString(heater.Content, font, brush, new XRect(0, 10, page.Width, page.Height), format);
                    }
                    else if (heater.HorizontalAlignment == Alignment.Near)
                    {
                        format.Alignment = XStringAlignment.Near;
                        format.LineAlignment = XLineAlignment.Near;
                        xGraphics.DrawString(heater.Content, font, brush, new XRect(MarginLeft, 10, page.Width, page.Height), format);
                    }
                    else if (heater.HorizontalAlignment == Alignment.Far)
                    {
                        format.Alignment = XStringAlignment.Far;
                        format.LineAlignment = XLineAlignment.Near;
                        xGraphics.DrawString(heater.Content, font, brush, new XRect(0, 10, page.Width - MarginRight, page.Height), format);
                    }
                    xGraphics.Dispose();
                    AddLine(10 + heater.FontSize + 1, 5, page.Width - 5, page, width: 0.3, color: DHColor.BLack);
                }
            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddFooterkHelper(StringContent footer)
        {
            try
            {
                foreach (var page in _pdfDocument.Pages)
                {

                    XGraphics xGraphics = XGraphics.FromPdfPage(page);


                    XBrush brush = FindBruse(footer.Color);
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                    XFont font = new XFont(_STFangSo, footer.FontSize, XFontStyle.Regular, options);


                    XStringFormat format = new XStringFormat();
                    if (footer.HorizontalAlignment == Alignment.Center)
                    {

                        format.Alignment = XStringAlignment.Center;
                        format.LineAlignment = XLineAlignment.Near;
                        xGraphics.DrawString(footer.Content, font, brush, new XRect(0, page.Height - MarginBottom + 10, page.Width, page.Height), format);
                    }
                    else if (footer.HorizontalAlignment == Alignment.Near)
                    {
                        format.Alignment = XStringAlignment.Near;
                        format.LineAlignment = XLineAlignment.Near;
                        xGraphics.DrawString(footer.Content, font, brush, new XRect(MarginLeft, page.Height - MarginBottom + 10, page.Width, page.Height), format);
                    }
                    else if (footer.HorizontalAlignment == Alignment.Far)
                    {
                        format.Alignment = XStringAlignment.Far;
                        format.LineAlignment = XLineAlignment.Near;
                        xGraphics.DrawString(footer.Content, font, brush, new XRect(0, page.Height - MarginBottom + 10, page.Width - MarginRight, page.Height), format);
                    }
                    xGraphics.Dispose();
                    AddLine(page.Height - MarginBottom + 9, 5, page.Width - 5, page, width: 0.3, color: DHColor.BLack);
                }
            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddPageNumberHelper(Alignment horizontalAlignment, Alignment verticalAlignment)
        {
            try
            {
                int i = 1;
                foreach (var page in _pdfDocument.Pages)
                {


                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                    XFont font = new XFont(_STFangSo, 8, XFontStyle.Regular, options);

                    XStringFormat format = new XStringFormat();

                    if (horizontalAlignment == Alignment.Near)
                    {
                        format.Alignment = XStringAlignment.Near;
                        format.LineAlignment = XLineAlignment.Near;
                    }
                    else if (horizontalAlignment == Alignment.Center)
                    {
                        format.Alignment = XStringAlignment.Near;
                        format.LineAlignment = XLineAlignment.Near;
                    }
                    else //horizontalAlignment==Alignment.Far
                    {
                        format.Alignment = XStringAlignment.Far;
                        format.LineAlignment = XLineAlignment.Near;
                    }

                    double yPostion = 0;
                    if (verticalAlignment == Alignment.Near)
                    {
                        yPostion = 10;
                    }
                    else
                    {
                        yPostion = page.Height - MarginBottom + 10;
                    }


                    XGraphics xGraphics = XGraphics.FromPdfPage(page);
                    string pageNumber = $"- {i} -";
                    xGraphics.DrawString(pageNumber, font, XBrushes.Black, new XRect(MarginLeft, yPostion, page.Width - MarginRight - MarginLeft, page.Height), format);
                    xGraphics.Dispose();
                    i++;
                }
            }
            catch (Exception em)
            {

                throw;
            }
        }
        #endregion Helper

        #region tools
        /// <summary>
        /// 可自动换行的写入PDF字符串
        /// </summary>
        /// <param name="WriteStr"></param>
        /// <param name="font"></param>
        /// <param name="brush"></param>
        /// <param name="startXPos"></param>
        /// <param name="endXPos"></param>
        /// <param name="lineSpacing"></param>
        /// <param name="YPos"></param>
        /// <param name="page"></param>
        private void AddString_AutoLineFeed(string WriteStr, XFont font, XBrush brush, int startXPos, int endXPos, int lineSpacing, XStringFormat xStringFormat, ref double YPos, ref PdfPage page)
        {
            try
            {

                XGraphics xGraphics = XGraphics.FromPdfPage(page);


                foreach (var lineStr in SplitTextIntoLines(WriteStr, font, endXPos - startXPos, xGraphics))
                {
                    //判断是否需要添加页面
                    XSize lineSize = xGraphics.MeasureString(lineStr, font);
                    if (YPos + lineSize.Height > page.Height - MarginBottom)
                    {
                        AddPage(ref YPos, ref page);
                        xGraphics.Dispose();
                        xGraphics = XGraphics.FromPdfPage(page);
                    }

                    xGraphics.DrawString(lineStr, font, brush, new XRect(startXPos, YPos, page.Width, page.Height), xStringFormat);
                    YPos += (int)lineSize.Height + lineSpacing;


                }
                xGraphics.Dispose();
            }
            catch (Exception em)
            {

                throw;
            }

        }
        private void AddString(double X, double Y, StringContent content, PdfPage page)
        {
            try
            {


                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                XFont font = new XFont(_STFangSo, content.FontSize, XFontStyle.Regular, options);
                XBrush xBrush = FindBruse(content.Color);

                XGraphics xGraphics = XGraphics.FromPdfPage(page);
                xGraphics.DrawString(content.Content, font, xBrush, X, Y, XStringFormat.TopLeft);
                xGraphics.Dispose();


            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddImage(double X, double Y, Image Image, PdfPage page)
        {
            try
            {
                XImage xImage = XImage.FromFile(Image.FilePath);
                MeasureImage(Image, out double height, out double width);

                xImage.Interpolate = true;
                XGraphics xGraphics = XGraphics.FromPdfPage(page);
                xGraphics.DrawImage(xImage, X, Y, width, height);
                xGraphics.Dispose();

            }
            catch (Exception em)
            {

                throw;
            }
        }

        private void AddLine(double XorY, double startPoint, double endPoint, PdfPage page, double width = 1, DHColor color = DHColor.BLack, bool isTransverseLine = true)
        {
            try
            {
                XGraphics xGraphics = XGraphics.FromPdfPage(page);
                XColor xColor = FindColor(color);

                if (isTransverseLine == true)
                {
                    xGraphics.DrawLine(new XPen(xColor, width), new XPoint(startPoint, XorY), new XPoint(endPoint, XorY));
                }
                else
                {
                    xGraphics.DrawLine(new XPen(xColor, width), new XPoint(XorY, startPoint), new XPoint(XorY, endPoint));
                }

                xGraphics.Dispose();
            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void AddPage(ref double YPos, ref PdfPage page)
        {
            try
            {
                YPos = MarginTop;
                page = _pdfDocument.AddPage();
            }
            catch (Exception em)
            {

                throw em;
            }
        }

        private void AddTableCell(TableCell cell, double X, double Y, PdfPage page)
        {
            try
            {
                double acStartX = X + (double)cell.StartX;
                double acStartY = Y + (double)cell.StartY;
                double acEndX = X + (double)cell.EndX;
                double acEndY = Y + (double)cell.EndY;
                if (cell.TopLine)
                {
                    AddLine(acStartY, acStartX, acEndX, page);
                }
                if (cell.BottomLine)
                {
                    AddLine(acStartY + cell.Height, acStartX, acEndX, page);
                }
                if (cell.LeftLine)
                {
                    AddLine(acStartX, acStartY, acEndY, page, isTransverseLine: false);
                }
                if (cell.RightLine)
                {
                    AddLine(acStartX + cell.Width, acStartY, acEndY, page, isTransverseLine: false);
                }

                XGraphics xGraphics = XGraphics.FromPdfPage(page);
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                XFont xFont = new XFont(_STFangSo, cell.CellText.FontSize, XFontStyle.Regular, options);
                XBrush xBrush = FindBruse(cell.CellText.Color);
                XStringFormat format = new XStringFormat();
                format.Alignment = XStringAlignment.Center;
                format.LineAlignment = XLineAlignment.Center;

                xGraphics.DrawString(cell.CellText.Content, xFont, xBrush, new XRect(acStartX, acStartY, acEndX - acStartX, acEndY - acStartY), format);
                xGraphics.Dispose();

            }
            catch (Exception em)
            {

                throw;
            }
        }

        private void MeasureString(StringContent content, out double height, out double width)
        {
            try
            {
                if (content == null || content.Content == null)
                {
                    height = 0;
                    width = 0;
                    return;
                }
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
                XFont font = new XFont(_STFangSo, content.FontSize, XFontStyle.Regular, options);

                XGraphics xGraphics = XGraphics.FromPdfPage(_currentPage);
                XSize xSize = xGraphics.MeasureString(content.Content, font);
                xGraphics.Dispose();
                height = xSize.Height;
                width = xSize.Width;
            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void MeasureImage(Image image, out double height, out double width)
        {
            height = 0;
            width = 0;
            try
            {
                XImage xImage = XImage.FromFile(image.FilePath);
                if (image.Width != null || image.Height != null)
                {
                    width = image.Width != null ? (double)image.Width : xImage.Size.Width;
                    height = image.Height != null ? (double)image.Height : xImage.Size.Height;
                }
                else if (image.TransformationMultiple != null)
                {
                    width = xImage.Size.Width * (double)image.TransformationMultiple;
                    height = xImage.Size.Height * (double)image.TransformationMultiple;
                }
                else
                {
                    width = xImage.Size.Width;
                    height = xImage.Size.Height;
                }


            }
            catch (Exception em)
            {

                throw;
            }
        }
        /// <summary>
        /// 将Text分割成多行文本以适应给定的宽度
        /// </summary>
        /// <param name="text"></param>
        /// <param name="xFont"></param>
        /// <param name="maxWidth"></param>
        /// <returns></returns>
        private string[] SplitTextIntoLines(string text, XFont xFont, double maxWidth, XGraphics xGraphics)
        {
            try
            {



                List<string> lines = new List<string>();

                StringBuilder currentLine = new StringBuilder();
                foreach (var chars in text.ToArray())
                {
                    if (chars.ToString() == "\n" || chars.ToString() == "\r\n")
                    {
                        lines.Add(currentLine.ToString());
                        currentLine = new StringBuilder();
                        continue;
                    }

                    if (chars.ToString() == "\t")
                    {

                        currentLine.Append("      ");
                        if (xGraphics.MeasureString(currentLine.ToString(), xFont).Width > maxWidth)
                        {
                            lines.Add(currentLine.ToString());
                            currentLine = new StringBuilder();
                        }
                    }

                    switch (GetCharRepresent(chars))
                    {
                        case 0: //字母
                            currentLine.Append(chars);
                            if (xGraphics.MeasureString(currentLine.ToString(), xFont).Width > maxWidth)
                            {
                                lines.Add(currentLine.ToString());
                                currentLine = new StringBuilder();
                            }
                            break;

                        case 1:  //非中文宽字符
                        case 2:  //中文
                        case 3:  // 数字

                            currentLine.Append(chars);
                            if (xGraphics.MeasureString(currentLine.ToString(), xFont).Width > maxWidth - 10)
                            {
                                lines.Add(currentLine.ToString());
                                currentLine = new StringBuilder();
                            }
                            break;
                        case 4:  //其他字符
                            //把标点留在上一行
                            if (currentLine.Length == 0 && lines.Count != 0)
                            {
                                var lastLine = lines[lines.Count - 1];

                                if (GetCharRepresent(lastLine[lastLine.Length - 1]) != 4)
                                {
                                    lastLine.Append(chars);
                                    break;
                                }
                            }


                            currentLine.Append(chars);
                            if (xGraphics.MeasureString(currentLine.ToString(), xFont).Width > maxWidth - 10)
                            {
                                lines.Add(currentLine.ToString());
                                currentLine = new StringBuilder();
                            }
                            break;
                        default:
                            break;
                    }

                }
                lines.Add(currentLine.ToString());


#if DEBUG
                foreach (var line in lines)
                {
                    string writeLine = line;
                    if (line.Contains("\r") && line.LastIndexOf("\r") == line.Length - 1)
                    {
                        writeLine = line.Remove(line.Length - 2);
                    }
                    Console.WriteLine(":start:" + writeLine + ":end:");
                }

#endif



                return lines.ToArray();

            }
            catch (Exception em)
            {

                throw;
            }
        }

        /// <summary>
        /// 判断char类型
        /// </summary>
        /// <param name="ch"></param>
        /// <returns> 0--字母 1--非中文宽字符 2--中文 3--数字 4--其他字符</returns>
        /// <exception cref="ArgumentException"></exception>
        private int GetCharRepresent(char ch)
        {
            if (char.IsPunctuation(ch) || char.IsSeparator(ch) || char.IsSymbol(ch))
                return 4;
            if ((int)ch > 255)
            {
                string chStr = ch.ToString();
                if (!IsChineseCh(chStr))
                    return 1;//非中文宽字符
                else
                    return 2;//中文
            }
            else if ((int)ch <= 255)
            {
                if (char.IsLetter(ch))
                    return 0;//字母
                else if (char.IsNumber(ch))
                    return 3;//数字
                else
                    return 4;//其他字符
            }
            else
                throw new ArgumentException(ch.ToString());
        }
        private bool CreatDirectory(string path)
        {
            try
            {
                do
                {
                    try
                    {
                        if (path[path.Length - 1].ToString() == "\\")
                        {
                            break;
                        }
                        path = path.Remove(path.Length - 1);
                    }
                    catch (Exception em)
                    {

                        throw em;
                    }


                } while (true);

                if (Directory.Exists(path) == false)
                {
                    Directory.CreateDirectory(path);
                }

                return true;
            }
            catch (Exception em)
            {

                throw em;
            }

        }

        /// <summary>
        /// 判断是否是中文字符
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private bool IsChineseCh(string input)
        {
            Regex regex = new Regex("^[\u4e00-\u9fa5]+$");
            return regex.IsMatch(input);
        }
        private XBrush FindBruse(DHColor dHColor)
        {
            try
            {
                switch (dHColor)
                {
                    case DHColor.Red:
                        return XBrushes.Red;

                    case DHColor.BLack:
                        return XBrushes.Black;
                    case DHColor.Green:
                        return XBrushes.Green;
                    case DHColor.Blue:
                        return XBrushes.Blue;
                    default:
                        return XBrushes.Black;


                }


            }
            catch (Exception em)
            {

                throw;
            }
        }
        private XFont FindXFont(DHFont dHFont)
        {
            XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);
            XFontStyle xFontStyle = FindFontStyle(dHFont.Style);
            switch (dHFont.Type)
            {
                case DHFontType.STFANGSO:
                    return new XFont(_STFangSo, dHFont.Size, xFontStyle, options);
                    break;
                default:
                    return new XFont(_STFangSo, dHFont.Size, xFontStyle, options);
                    break;
            }



        }
        private XFontStyle FindFontStyle(DHFontStyle fontStyle)
        {
            switch (fontStyle)
            {
                case DHFontStyle.Regular: return XFontStyle.Regular;
                case DHFontStyle.Bold: return XFontStyle.Bold;
                case DHFontStyle.Italic: return XFontStyle.Italic;
                case DHFontStyle.Strikeout: return XFontStyle.Strikeout;
                case DHFontStyle.Underline: return XFontStyle.Underline;
                case DHFontStyle.BoldItalic:return XFontStyle.BoldItalic;
                default:
                    return XFontStyle.Regular;
            }

        }
        private XBrush FindWatermarkBruse(DHColor dHColor)
        {
            try
            {
                switch (dHColor)
                {
                    case DHColor.Red:
                        return new XSolidBrush(XColor.FromArgb(90, 255, 0, 0));

                    case DHColor.BLack:
                        return new XSolidBrush(XColor.FromArgb(90, 255, 255, 255));
                    case DHColor.Green:
                        return new XSolidBrush(XColor.FromArgb(90, 0, 255, 0));
                    case DHColor.Blue:
                        return new XSolidBrush(XColor.FromArgb(90, 0, 0, 255));
                    default:
                        return new XSolidBrush(XColor.FromArgb(90, 255, 0, 0));


                }


            }
            catch (Exception em)
            {

                throw;
            }
        }
        private XColor FindColor(DHColor dHColor)
        {
            try
            {
                switch (dHColor)
                {
                    case DHColor.Red:
                        return XColors.Red;

                    case DHColor.BLack:
                        return XColors.Black;
                    case DHColor.Green:
                        return XColors.Green;
                    case DHColor.Blue:
                        return XColors.Blue;
                    default:
                        return XColors.Black;


                }


            }
            catch (Exception em)
            {

                throw;
            }
        }
        private void FindFont()
        {


            try
            {
                //找到系统字体
                System.Drawing.Text.PrivateFontCollection pdfFonts = new System.Drawing.Text.PrivateFontCollection();

                byte[] font1 = Properties.FontResource.STFANGSO;
                IntPtr intPtr = Marshal.UnsafeAddrOfPinnedArrayElement(font1, 0);
                pdfFonts.AddMemoryFont(intPtr, font1.Length);

                _STFangSo = pdfFonts.Families[0];


                //设定PDF默认字体
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

                _defualtWatermarkFont = new XFont(_STFangSo, 20, XFontStyle.Bold, options);
                _defualtTittleFont = new XFont(_STFangSo, 20, XFontStyle.Bold, options);

                _defualtOneLevelHeading = new XFont(_STFangSo, 18, XFontStyle.Bold, options);
                _defualtTwoLevelHeading = new XFont(_STFangSo, 16, XFontStyle.Bold, options);
                _defualtThreeLevelHeading = new XFont(_STFangSo, 14, XFontStyle.Bold, options);
                _defualtFourLevelHeading = new XFont(_STFangSo, 12, XFontStyle.Bold, options);

                _defualtContentFont = new XFont(_STFangSo, 10, XFontStyle.Regular, options);
                _defualtCommentFont = new XFont(_STFangSo, 8, XFontStyle.Italic, options);

                return;
            }
            catch (Exception em)
            {


            }

            try
            {
                //找到适配中文的字体
                System.Drawing.Text.PrivateFontCollection pdfFonts = new System.Drawing.Text.PrivateFontCollection();

                string currentPath = System.Environment.CurrentDirectory;
                int rootIndex = currentPath.IndexOf("DHSOFTWARE");
                string rootPath = string.Empty;
                if (rootIndex != -1)
                {
                    rootPath = currentPath.Remove(rootIndex);

                }
                else
                {

                    throw new Exception("未找到“DHSOFTWARE”根路径");
                }

                string strfontPath = rootPath + "DHSOFTWARE\\Utilities\\Utilities\\PDFCreater\\FontFile\\STFANGSO.TTF";
                pdfFonts.AddFontFile(strfontPath);

                //设定PDF默认字体
                XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

                _defualtWatermarkFont = new XFont(_STFangSo, 20, XFontStyle.Bold, options);
                _defualtTittleFont = new XFont(_STFangSo, 20, XFontStyle.Bold, options);

                _defualtOneLevelHeading = new XFont(_STFangSo, 18, XFontStyle.Bold, options);
                _defualtTwoLevelHeading = new XFont(_STFangSo, 16, XFontStyle.Bold, options);
                _defualtThreeLevelHeading = new XFont(_STFangSo, 14, XFontStyle.Bold, options);
                _defualtFourLevelHeading = new XFont(_STFangSo, 12, XFontStyle.Bold, options);

                _defualtContentFont = new XFont(_STFangSo, 10, XFontStyle.Regular, options);
                _defualtCommentFont = new XFont(_STFangSo, 8, XFontStyle.Italic, options);

                return;
            }
            catch (Exception em)
            {

                throw;
            }




        }
        #endregion tools



        #region   Properties
        public int MarginLeft { get; set; } = DEFAULT_MARGIN_LEFT_OR_RIGHT;
        public int MarginTop { get; set; } = DEFAULT_MARGIN_TOP_OR_BOTTOM;
        public int MarginRight { get; set; } = DEFAULT_MARGIN_LEFT_OR_RIGHT;
        public int MarginBottom { get; set; } = DEFAULT_MARGIN_TOP_OR_BOTTOM;

        public DHPDFContent DHPDFContent
        {
            get { return _dHPDFContent; }
            set { _dHPDFContent = value; }
        }

        #endregion   Properties
    }
}
