﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranslateApp.Data
{
    public class WSData
    {                                           //validation indeces
        public int IDColumn { get; set; }       // 0
        public int SrcColumn { get; set; }      // 1
        public string SrcLangCode { get; set; } // 2
        public int TrgColumn { get; set; }      // 3
        public string TrgLangCode { get; set; } // 4
        public bool[] valOk { get; set; }
        public WSData()
        {
            IDColumn = 0;
            SrcColumn = 0;
            SrcLangCode = string.Empty;
            TrgColumn = 0;
            TrgLangCode = string.Empty;
            valOk = new bool[5];
        }
        public bool CheckData()
        {
            foreach( bool val in valOk)
            {
                if (!val) return false;         
            }
            return true;
        }
    }
}
