using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace Utilities.ScriptManager.Helpers
{
    public static class HtmlExtensions
    {
        public static string FormatPlainText(this HtmlHelper helper, string plainText)
        {
            return FormatPlainText(plainText);
        }

        public static string FormatPlainText(string plainText)
        {
            StringBuilder html = new StringBuilder(plainText);
            html.Replace(Environment.NewLine, "<br/>");
            return html.ToString();
        }
    }
}