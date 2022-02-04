using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyAutomationFramework
{
    public class EasyReturn
    {

        public class Web
        {

            public IWebDriver driver { get; set; }
            public IWebElement element { get; set; }
            public DataTable table { get; set; }
            public string Error { get; set; }
            public string Value { get; set; }
            public bool Sucesso { get; set; }
            public Bitmap Bitmap { get; set; }

        }
    }
}
