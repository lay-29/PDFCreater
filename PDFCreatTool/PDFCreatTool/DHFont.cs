using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utilities.PDFCreater
{
   
    public class DHFont
	{
        #region  member data
        private DHFontType _type;
        private DHColor _color;
        private double _size;
        private DHFontStyle _fontStyle;
        #endregion  member data


        #region	life cycle
        public DHFont()
        {
            _type = DHFontType.STFANGSO;
            _color = DHColor.BLack;
            _size = 10;
            _fontStyle = DHFontStyle.Regular;
        }
        public DHFont(DHFontType fontType,DHColor fontColor,double size, DHFontStyle fontStyle)                              
        {
            _type = fontType;
            _color = fontColor;
            _size= size;
            _fontStyle= fontStyle;
        }
        #endregion life cycle

        #region  methods

        #endregion methods

        #region properties
        public double Size
        {
            get { return _size; }
            set { _size = value; }
        }

        public DHColor Color
        {
            get { return _color; }
            set { _color = value; }
        }

        public DHFontType Type
        {
            get { return _type; }
            set { _type = value; }
        }
        public DHFontStyle Style
        {
            get { return _fontStyle; }
            set { _fontStyle = value; }
        }
        #endregion properties


    }
}
