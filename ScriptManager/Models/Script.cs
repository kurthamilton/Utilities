using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

using Utilities.ScriptManager.Helpers;

namespace Utilities.ScriptManager.Models
{
    public enum ScriptType : int
    {
        None = 0,
        Sql = 1,
        Javascript = 2
    }

    public class Script
    {
        public ScriptType Type { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }
        public IEnumerable<Tag> Tags { get; set; }
    }
}