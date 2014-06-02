using System;
using System.Configuration;
using System.IO;
using System.Web;

namespace Utilities.Office
{
    public class OfficeIO
    {
        internal const string WorkingFolderName = @"Working Folder\";
        internal const string OfficeTemplateFolderName = @"Office Templates\";

        /***********************************
         * Static methods
         ************************************/

        public static string GetValidFileName(string fileName)
        {
            return string.Join("", fileName.Split(Path.GetInvalidFileNameChars()));
        }

        public static string GetMicrosoftOfficePath()
        {
            string microsoftOfficePath = ConfigurationManager.AppSettings["MicrosoftOfficePath"];
            if (!string.IsNullOrEmpty(microsoftOfficePath))
            {
                if (Directory.Exists(microsoftOfficePath))
                    return microsoftOfficePath;
                throw new Exception(string.Format("Microsoft Office path {0} not found", microsoftOfficePath));
            }
            throw new Exception("Microsoft Office path not set in web.config");
        }

        public static string GetOfficeTemplatePath()
        {
            string microsoftOfficePath = GetMicrosoftOfficePath();
            string officeTemplatePath = Path.Combine(microsoftOfficePath, OfficeTemplateFolderName);
            if (!Directory.Exists(officeTemplatePath))
                throw new Exception(string.Format("Microsoft Office template path {0} not found", officeTemplatePath));
            return officeTemplatePath;
        }

        public static string GetWorkingFolderPath()
        {
            string microsoftOfficePath = GetMicrosoftOfficePath();
            string officeTemplatePath = Path.Combine(microsoftOfficePath, WorkingFolderName);
            if (!Directory.Exists(officeTemplatePath))
                Directory.CreateDirectory(officeTemplatePath);
            return officeTemplatePath;
        }

        public static string UploadWorkingFile(HttpPostedFileBase postedFile, string targetPath, string targetFileName)
        {
            if (postedFile != null)
            {
                if (targetPath == "")
                    targetPath = GetWorkingFolderPath();

                FileInfo fileInfo = new FileInfo(postedFile.FileName);
                string extension = fileInfo.Extension;

                if (targetFileName == "")
                    targetFileName = fileInfo.Name.Replace(extension, "");

                string workingFilePath = "";
                workingFilePath = GetVersionedSaveFilePath(targetPath, string.Concat(targetFileName, extension));
                postedFile.SaveAs(workingFilePath);

                return workingFilePath;
            }
            else
                return "";
        }

        public static string UploadWorkingFile(HttpPostedFile postedFile, string targetPath, string targetFileName)
        {
            HttpPostedFileBase postedFileBase = Helpers.GetHttpPostedFileBaseFromHttpPostedFile(postedFile);
            return UploadWorkingFile(postedFileBase, targetPath, targetFileName);
        }

        public static string UploadWorkingFile(HttpPostedFile postedFile, string targetFileName)
        {
            return UploadWorkingFile(postedFile, "", targetFileName);
        }

        public static string UploadWorkingFile(HttpPostedFile postedFile)
        {
            return UploadWorkingFile(postedFile, "", "");
        }

        public static string TryCopyFile(string sourceFilePath, string targetPath, string targetFileName)
        {
            if (File.Exists(sourceFilePath))
            {
                string newFilePath = GetVersionedSaveFilePath(targetPath, targetFileName);

                if (!Directory.Exists(targetPath))
                    Directory.CreateDirectory(targetPath);

                File.Copy(sourceFilePath, newFilePath);
                return newFilePath;
            }
            else
                return "";
        }


        internal static string GetVersionedSaveFilePath(string path, string fileName)
        {
            fileName = GetValidFileName(fileName);

            int version = 1;
            string workingPath = GetWorkingFolderPath();
            FileInfo fileInfo = new FileInfo(fileName);
            string fileExtension = fileInfo.Extension;
            string targetFileName = fileName.Substring(0, fileName.Length - fileExtension.Length);
            string workingFilePath;
            do
            {
                string workingFileName = string.Concat(targetFileName, (version > 1 ? string.Format(" ({0})", version) : ""), fileExtension);
                workingFilePath = Path.Combine(path, workingFileName);

                version++;
            } while (File.Exists(workingFilePath));

            return workingFilePath;
        }

        public static bool TryDeleteFile(string filePath)
        {
            // attempt to delete a file. System file locks sometimes carry over, so loop until deleted
            // if there is still a more permanent file lock, then escape function if trying for more than 10 seconds
            if (File.Exists(filePath))
            {
                bool success = false;
                DateTime startTime = DateTime.Now;

                do
                {
                    try
                    {
                        File.Delete(filePath);
                        success = true;
                    }
                    catch (IOException ex)
                    {
                        string message = ex.Message;
                    }
                }
                while (!success && (DateTime.Now - startTime) < TimeSpan.FromSeconds(10));

                return success;
            }

            return true;
        }
    }
}
