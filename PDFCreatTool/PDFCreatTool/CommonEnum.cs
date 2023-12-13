using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFCreatTool.Properties
{
    public enum DHFontType
    {
        STFANGSO,// 华文仿宋
    }

    public enum DHColor
    {
       Red,
       BLack,
       Green,
       Blue
    }
  
    public enum DHFontStyle 
    {
        Regular,   //普通样式
        Bold,       //加粗 
        BoldItalic,  //加粗斜体
        Italic,      //斜体
        Strikeout, //删除线
        Underline, //下划线
    }
    public enum Alignment
    {
         Near, //靠左或上
         Center,
         Far  //靠右或下
    }
    public enum ContentType
    {   PDFContent,
        String,
        LevelText,
        Image,
        Table,
        Line
    }
}
