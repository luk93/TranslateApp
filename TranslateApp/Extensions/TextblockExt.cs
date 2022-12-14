﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Documents;

namespace TranslateApp.Extensions
{
    public static class TextblockExt
    {
        public static void AddLine(this TextBlock tb, string text)
        {
            if (!string.IsNullOrEmpty(text))
                tb.Inlines.InsertBefore(tb.Inlines.FirstInline, new Run(text));
        }
    }
}
