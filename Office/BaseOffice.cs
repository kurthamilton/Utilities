using System;

namespace Utilities.Office
{
    // All OpenXml standards documents are here: http://www.ecma-international.org/news/TC45_current_work/TC45_available_docs.htm

    public abstract class BaseOffice
    {
        // NB - known bug
        // There is a bug in System.IO.Packaging when dealing with large (~ > 10MB) PackageParts. It unzips the whole Package to IsolatedStorage, but 
        // the user account doesn't have sufficient privileges to access it, so it crashes. 

        // Further explanation:
        // Basically did what this suggested giving full rights to "Network Service" on the IsolatedStorage folder:
        // http://excelpackage.codeplex.com/workitem/17586
        // except the profile was "Default User.Windows" for whatever reason.
    }
}
