using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;

namespace wfp_lab1.Services
{
    internal interface IWordPrintable
    {
        public Document ToWordDoc();
    }
}
