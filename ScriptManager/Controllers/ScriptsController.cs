using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.Configuration;
using System.IO;

using Utilities.ScriptManager.Models;

namespace Utilities.ScriptManager.Controllers
{
    public class ScriptsController : Controller
    {
        //
        // GET: /Scripts/

        public ActionResult Index()
        {
            IEnumerable<Script> scripts = FindScripts();

            return View(scripts);
        }

        public ActionResult RefreshScripts()
        {
            IEnumerable<Script> scripts = FindScripts();

            return View(scripts);
        }


        // All of this stuff to go into a service layer
        private IEnumerable<Script> FindScripts()
        {
            string scriptPath = ConfigurationManager.AppSettings["ScriptPath"];
            if (!Directory.Exists(scriptPath))
            {
                throw new Exception(string.Format("ScriptPath not found: {0}", scriptPath));
            }

            return GetScripts(scriptPath);
        }

        private IEnumerable<Script> GetScripts(string scriptPath)
        {
            List<Script> scripts = new List<Script>();

            foreach (string filePath in Directory.GetFiles(scriptPath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                string value = System.IO.File.ReadAllText(filePath);
                ScriptType type = ScriptType.None;
                switch (fileInfo.Extension.ToLower())
                {
                    case ".sql":
                        type = ScriptType.Sql;
                        break;
                }
                scripts.Add(new Script() { Type = type, Name = filePath, Value = value });
            }

            foreach (string directory in Directory.GetDirectories(scriptPath))
            {
                scripts.AddRange(GetScripts(directory));
            }
            
            return scripts;
        }
    }
}