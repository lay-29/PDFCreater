using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utilities.PDFCreater
{
    /// <summary>
    /// PDf内容对象的基类
    /// </summary>
    public abstract class PDFContentBase
    {
        #region member data
        private ContentType _contentType;
        #endregion member data

        #region  life cycle
        internal PDFContentBase(ContentType contentType)
        {
            _contentType= contentType;
        }
        #endregion life cycle

        #region methods

        #endregion methods

        #region  properties
        public ContentType ContentType
        {
            get { return _contentType; }
            set { _contentType = value; }
        }

        #endregion   properties


    }

    /// <summary>
    /// PDF的所有内容
    /// </summary>
    public class DHPDFContent :PDFContentBase
    {
        #region member data

        private List<PDFContentBase> _content;

        #endregion  member data

        #region life cycle 
        public DHPDFContent() : base(ContentType.PDFContent)
        {
            _content = new List<PDFContentBase>();
            //第一个内容是标题
            _content.Add(new OneLevelContent(0,"","")); 
        }
        #endregion  life cycle

        #region methods
        public void AddContent(PDFContentBase contentBase)
        {
            Content.Add(contentBase);
        }
        public void DeleteContent(PDFContentBase contentBase)
        {
            if (Content.Contains(contentBase))
            {
                Content.Remove(contentBase);
            }
        }
        public void DeleteContent(int contentIndex)
        {
            if (0<contentIndex&&contentIndex < Content.Count)
            {
                Content.RemoveAt(contentIndex);
            }
        }
        #endregion methods

        #region properties
        public List<PDFContentBase> Content
        {
            get { return _content; }
            set { _content = value; }
        }
        public StringContent Watermark { get; set; } = null;
        public StringContent Header { get; set; } = null;
        public StringContent Footer { get; set; } = null;
        public (Alignment horizontalAlignment, Alignment verticalAlignment)? PageNumberAlignment { get; internal set; } = null;
        #endregion  properties


    }

    /// <summary>
    /// 一个等级标题下的内容
    /// </summary>
    public class OneLevelContent :PDFContentBase
    {

        #region  member data
        private int _levelIndex;  //等级指数 ，0--pdf文档标题 。1--一级标题， 2--二级标题， ····· 4--四级标题
        private string _levelTittle;
        private string _levelContent;  //当前标题下的文本内容。
        #endregion member data

        #region life cycle
        public OneLevelContent(int level):base(ContentType.LevelText)
        {
            SubLevelContent = new List<OneLevelContent>();
            _levelIndex = level;
            _levelTittle = string.Empty;
            _levelContent = string.Empty;
        }
        public OneLevelContent(string contentText) : base(ContentType.LevelText)
        {
            SubLevelContent = new List<OneLevelContent>();
            _levelIndex = 4;
            _levelTittle = string.Empty;
            _levelContent = contentText;

        }
        public OneLevelContent(string levelTittle,string contentText) : base(ContentType.LevelText)
        {
            SubLevelContent = new List<OneLevelContent>();
            _levelIndex = 4;
            _levelTittle= levelTittle;
            _levelContent = contentText;

        }
        public OneLevelContent(int level, string levelTittle ,string contentText) : base(ContentType.LevelText)
        {
            SubLevelContent = new List<OneLevelContent>();
            _levelIndex = level;
            _levelTittle = levelTittle;
            _levelContent = contentText;

        }

        #endregion life cycle

        #region  methods
        public void AddSubLevelContent(OneLevelContent oneLevelContent)
        {
            oneLevelContent.LevelIndex = _levelIndex + 1;
            SubLevelContent.Add(oneLevelContent);
        }

        #endregion methods

        #region properties
        public int LevelIndex
        {
            get { return _levelIndex; }
            set
            {

                if (value > 4)
                {
                    _levelIndex = 4;
                }
                else if (value < 0)
                {
                    _levelIndex = 0;
                }
                else
                {
                    _levelIndex = value;
                }

            }
        }

        public string LevelTittle
        {
            get { return _levelTittle; }
            set { _levelTittle = value; }
        }
        public string LevelContent
        {
            get { return _levelContent; }
            set { _levelContent = value; }
        }

        public int TittleRetractSpacing { get; set; } = 0;
        public int TittleLineSpacing { get; set; } = 4;


        public int ContentRetractSpacing { get; set; } = 0;
        public int ContentLineSpacing { get; set; } = 2;


        public DHFont TittleFont { get; set; } = null;
        public DHFont ContentFont { get; set; } = null;
        internal List<OneLevelContent> SubLevelContent { get; private set; }

        #endregion properties

    }

    /// <summary>
    /// 一个字符串内容
    /// </summary>
    public class StringContent : PDFContentBase
    {
        public StringContent(string content) : base(ContentType.String)
        {
            Content = content;
        }


        public DHColor Color { get; set; } = DHColor.BLack;
        public string Content { get; set; } ="";
        public double FontSize { get; set; } = 10;
        public Alignment HorizontalAlignment { get; set; } = Alignment.Near;

        public int RetractSpacing { get; set; } = 0;
        public int LineSpacing { get; set; } = 4;
    }


   /// <summary>
   /// 表格内容
   /// </summary>
    public class Table : PDFContentBase
    {
        public Table(int rowNumber,int columnNumber) : base(ContentType.Table)
        {
            RowNumber = rowNumber;
            ColumnNumber = columnNumber;
            TableContent=new Dictionary<(int, int), TableCell> ();
            for (int row=1;row<=RowNumber;row++)
            {
                for (int col = 1; col <=ColumnNumber;col++)
                {
                    TableContent.Add((row, col), new TableCell("")); 
                }
            }

        }

        public bool SetCellText(int row,int col,string text)
        {
            try
            {
                if(GetCell(row,col,out TableCell tablecell)==true)
                {
                    tablecell.CellText.Content = text;
                    return true;
                }
                return false;
               
            }
            catch (Exception em)
            {

                return false;
            }
        }
        public bool SetCellText(int row, int col, StringContent stringContent)
        {
            try
            {
                GetCell(row, col, out TableCell tablecell);
                tablecell.CellText = stringContent;
                return true;
            }
            catch (Exception em)
            {

                return false;
            }
        }
        public bool GetCell(int row,int col,out TableCell tableCell)
        {
            tableCell = null;
            try
            {
                if(TableContent.ContainsKey((row, col)))
                {
                    tableCell=TableContent[(row, col)];
                }
                else
                {
                    return false;
                }


                return true;
            }
            catch (Exception em)
            {

                return false;
            }
        }

        public Dictionary<(int,int),TableCell> TableContent { get; set; }
        public int RowNumber { get;  } 
        public int ColumnNumber { get;  }
        public StringContent Comment { get; set; } = new StringContent("") { FontSize=8,Color=DHColor.BLack,HorizontalAlignment=Alignment.Center};
        public bool ShowComment { get; set; } = true;
        public double Width { get; set; }
        public double Height { get; set; }

    }

    /// <summary>
    /// 图片内容
    /// </summary>
    public class Image : PDFContentBase
    {
        private string _filePath;
        private double? _height;
        private double? _width;
        private double? _transformationMultiple;
        private Alignment _horizontalAlignment=Alignment.Center;
        private StringContent _comment = new StringContent("") { FontSize =8};

        public Image(string filePath) : base(ContentType.Image)
        {
            _filePath = filePath;
        }


        public string FilePath
        {
            get { return _filePath; }
            set { _filePath = value; }
        }
        public double? Width
        {
            get { return _width; }
            set { _width = value; }
        }

        public double? Height
        {
            get { return _height; }
            set { _height = value; }
        }

         /// <summary>
         /// 等比例放大缩小的倍数   如果Width/Height!=null  该参数不起作用
         /// </summary>
        public double? TransformationMultiple
        {
            get { return _transformationMultiple; }
            set { _transformationMultiple = value; }
        }
        public Alignment HorizontalAlignment
        {
            get { return _horizontalAlignment; }
            set { _horizontalAlignment = value; }
        }

        public StringContent Comment
        {
            get { return _comment; }
            set { _comment = value; }
        }

    }

    public class Line : PDFContentBase
    {
        private double _startPoint;
        private double  _endPoint;
        private bool _isHorizontalLine;
      
        public Line(double start,double end,bool isHorizontalLine=true) : base(ContentType.Line)
        {
            _startPoint=start;
            _endPoint = end;
            _isHorizontalLine = isHorizontalLine;
        }


        public double EndPoint
        {
            get { return _endPoint; }
            set { _endPoint = value; }
        }

        public double StartPoint
        {
            get { return _startPoint; }
            set { _startPoint = value; }
        }
        public bool IsHorizontalLine
        {
            get { return _isHorizontalLine; }
            set { _isHorizontalLine = value; }
        }
        public DHColor Color { get; set; } = DHColor.BLack;
        public double LineWidth { get; set; } = 1;
    }
    public class TableCell
    {

        public TableCell(string CellStr)
        {
            CellText = new StringContent(CellStr);
        }

        public StringContent CellText { get; set; }
        public double? StartX { get; set; } = null;
        public double? EndX { get; set; } = null;
        public double? StartY { get; set; } = null;
        public double? EndY { get; set; } = null;


        public double Height { get; set; }
        public double Width { get; set; }

        public bool TopLine { get; set; } = true;
        public bool BottomLine { get; set; } = true;
        public bool LeftLine { get; set; } = true;
        public bool RightLine { get; set; } = true;
    }
}
