using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFCreatTool.Properties
{
    public interface IPDFCreater
    {

        #region  methods
        /// <summary>
        /// Set Tittle For Pdf File
        /// </summary>
        void SetTittle(string tittle, string tittleContent = "");


        /// <summary>
        /// 向PDF文档中添加内容
        /// </summary>
        /// <param name="content"></param>
        void AddContent(PDFContentBase content);

        /// <summary>
        /// 添加统一格式的公司水印
        /// </summary>
        /// <param name="text"></param>
        void AddWatermark(string text = "***有限公司");
        /// <summary>
        /// add  WaterMark on All Pages
        /// </summary>
        /// <param name="waterMark"></param>
        void AddWatermark(StringContent watermark);


        void AddHeader(StringContent header);
        void AddHeader(string header = "***有限公司");
        void AddFooter(StringContent footer);
        void AddFooter(string footer = "***有限公司");
        void SetPageNumber(Alignment horizontalAlignment, Alignment verticalAlignment);
        /// <summary>
        /// Set Margin For Pdf
        /// </summary>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="right"></param>
        /// <param name="bottom"></param>
        void SetMargin(int? left = null, int? top = null, int? right = null, int? bottom = null);
        #endregion   methods


        /// <summary>
        /// 保存PDF文件
        /// </summary>
        /// <param name="path">保存地址，包含文件名</param>
        /// <param name="openAfterSave">是否打开文件</param>
        /// <param name="isSync">如果是异步模式，</param>
        void Save(string path, bool openAfterSave = false, bool isSync = true);

        #region   properties
        int MarginLeft { get; set; }
        int MarginTop { get; set; }
        int MarginRight { get; set; }
        int MarginBottom { get; set; }

        #endregion    properties
    }
}


/*Demo
 * public static void CreatPDFTest()
        {
            try
            {
                IPDFCreater pDFCreater = new PDFCreater();

                #region 设置文档标题及署名内容

                pDFCreater.SetTittle($"*****有限公司{DateTime.Now.ToString("g")}", "检查项目：*****\t检查时间：2023年5月1日");

                #endregion 设置文档标题及署名内容


                #region  添加一个等级的内容   //可以添加一级标题的内容 也可以添加二级标题的内容 是何等级由OneLevelContent.LevelIndex控制

                OneLevelContent oneLevelContent = new OneLevelContent(1, $"1.一级标题", "一````内容"); //参数1：LevelIndex 表示是一级内容

                for (int i = 1; i <= 2; i++)
                {

                    //创建一个二级内容
                    OneLevelContent oneLevelContent2 = new OneLevelContent($"1.{i}二级标题", "二````内容");
                    //设置该内容的标题缩进 如不设置默认值0
                    oneLevelContent2.TittleRetractSpacing = 8;
                    //设置该内容的文本内容缩进  如不设置默认值0
                    oneLevelContent2.ContentRetractSpacing = 16;
                    //设置该内容的标题行间距 如不设置默认值4
                    //oneLevelContent2.TittleLineSpacing = 10;
                    //设置该内容的标题行间距 如不设置默认值2
                    //oneLevelContent2.ContentLineSpacing = 16;
                    //添加到一级内容下
                    oneLevelContent.AddSubLevelContent(oneLevelContent2);

                    for (int j = 1; j <= 2; j++)
                    {
                        //三级同二级
                        OneLevelContent oneLevelContent3 = new OneLevelContent($"1.{i}.{j}三级标题", "三````内容");
                        oneLevelContent3.TittleRetractSpacing = 16;
                        oneLevelContent3.ContentRetractSpacing = 24;
                        //添加到二级内容下
                        oneLevelContent2.AddSubLevelContent(oneLevelContent3);
                        //如作为子等级添加到父等价下时，构造时不用声明LevelIndex，AddSubLevelContent方法自动将LevelIndex设置为父等级的下一等级
                    }
                }
                //将一个等级的文本内容添加到文档中
                pDFCreater.AddContent(oneLevelContent);
                #endregion    添加一个等级的内容


                #region  添加一个字符串   //短文本适用 
                StringContent TimeStr = new StringContent($"{DateTime.Now.ToString("g")}");
                TimeStr.HorizontalAlignment = Alignment.Far;//靠右
                TimeStr.RetractSpacing = 0;
                //设定字体大小
                TimeStr.FontSize = 30;
                TimeStr.LineSpacing = 0;
                pDFCreater.AddContent(TimeStr);

                StringContent TimeStr2 = new StringContent($"{DateTime.Now.ToString("g")}");
                TimeStr2.HorizontalAlignment = Alignment.Near;//靠左
                //设定字体颜色
                TimeStr2.Color = DHColor.Red;
                TimeStr2.RetractSpacing = 0;
                pDFCreater.AddContent(TimeStr2);

                StringContent TimeStr3 = new StringContent($"{DateTime.Now.ToString("g")}");
                TimeStr3.HorizontalAlignment = Alignment.Center;//居中
                TimeStr3.RetractSpacing = 0;
                pDFCreater.AddContent(TimeStr3);
                #endregion 添加一个字符串


                #region 表格
                //简单的创建一个表格
                {
                    //创建一个12行4列的表格
                    Table table = new Table(12, 4);
                    //设定表头内容
                    table.SetCellText(1, 1, "检查项目");
                    table.SetCellText(1, 2, "检查信息");
                    table.SetCellText(1, 3, "检查人");
                    table.SetCellText(1, 4, "检查时间");
                    for (int k = 2; k <= 12; k++)
                    {
                        //k是第几行，1，2，3，4是第几列
                        table.SetCellText(k, 1, $"{k - 1}.检查项目");
                        table.SetCellText(k, 2, $"{k - 1}.检查信息");
                        table.SetCellText(k, 3, $"{k - 1}.检查人");
                        table.SetCellText(k, 4, $"{k - 1}.{DateTime.Now.ToString("G")}");
                    }
                    //设定图名
                    table.Comment.Content = "1-1 图一";
                    pDFCreater.AddContent(table);
                }


                //设定一个独特格式的表格
                {
                    Table table = new Table(12, 4);
                    table.SetCellText(1, 1, "检查项目");
                    table.SetCellText(1, 2, "检查信息");
                    table.SetCellText(1, 3, "检查人");
                    table.SetCellText(1, 4, "检查时间");
                    for (int k = 2; k <= 12; k++)
                    {
                        table.SetCellText(k, 1, $"{k - 1}.检查项目");
                        table.SetCellText(k, 2, $"{k - 1}.检查信息");
                        table.SetCellText(k, 3, $"{k - 1}.检查人");
                        table.SetCellText(k, 4, $"{k - 1}.{DateTime.Now.ToString("G")}");
                    }
                    table.Comment.Content = "1-2 图二";
                    //以上内容同上

                    //遍历表格的每个单元格
                    foreach (var cell in table.TableContent)
                    {
                        //将第一行的上边框设置为不显示
                        if (cell.Key.Item1 == 1)
                        {
                            cell.Value.TopLine = false;
                        }

                        //将最后一行的下边框设置为不显示
                        if (cell.Key.Item1 == 12)
                        {
                            cell.Value.BottomLine = false;
                        }

                        //将第一列的左边框设置为不显示
                        if (cell.Key.Item2 == 1)
                        {
                            cell.Value.LeftLine = false;
                        }

                        //将最后一列的右边框设置为不显示
                        if (cell.Key.Item2 == 4)
                        {
                            cell.Value.RightLine = false;
                        }
                    }
                    pDFCreater.AddContent(table);
                }
              

                #endregion 表格


                #region 图片

                string ImagePath = "";
                //不用看 内容为找到图标的保存路径 
                {
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

                    ImagePath = rootPath + "DHSOFTWARE\\Utilities\\Utilities\\PDFCreater\\ImageFile\\DHBDT.jpg";
                }
                //图一     ImagePath为‘.../.../*.jpg’ 的格式
                Image image = new Image(ImagePath);
                //设定宽高
                image.Width = image.Height = 200;
                //设定居左
                image.HorizontalAlignment = Alignment.Near;
                //设定图名
                image.Comment.Content = "1-2 图二";
                pDFCreater.AddContent(image);

                //图二
                Image image2 = new Image(ImagePath);
                //设定缩放比例
                image2.TransformationMultiple = 0.02;
                //设定居中
                image2.HorizontalAlignment = Alignment.Center;
                image2.Comment.Content = "1-2 图二";
                pDFCreater.AddContent(image2);

                Image image3 = new Image(ImagePath);
                image3.TransformationMultiple = 0.02;
                //设定居右
                image3.HorizontalAlignment = Alignment.Far;
                image3.Comment.Content = "1-2 图二";
                pDFCreater.AddContent(image3);
                #endregion  图片


                //水印
                pDFCreater.AddWatermark();
                //页眉
                pDFCreater.AddHeader();
                //页脚
                pDFCreater.AddFooter();
                //设置页码      此设置为         居右    居下    及右下的页码 
                pDFCreater.SetPageNumber(Alignment.Far, Alignment.Far);

                //保存
                string savePath = @"C:\Users\123\Desktop\123.pdf";
                //保存地址，保存完是否打开
                pDFCreater.Save(savePath, true);
            }
            catch (Exception em)
            {

                Console.WriteLine(em.Message + em.StackTrace);
            }
        }
 * 
 */