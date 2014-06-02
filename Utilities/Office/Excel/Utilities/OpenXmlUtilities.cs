using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using OpenXml = DocumentFormat.OpenXml;
using OpenXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OpenXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Utilities.Office.Excel
{        
    internal static class OpenXmlUtilities
    {

        public static T CreatePart<P, T>(P parentPart)
            where P : OpenXmlPackaging.OpenXmlPart
            where T : OpenXmlPackaging.OpenXmlPart, OpenXmlPackaging.IFixedContentTypePart           
        {
            string relationshipId = OpenXmlUtilities.GenerateNewRelationshipId<P>(parentPart);
            T childPart = parentPart.AddNewPart<T>(relationshipId);
            return childPart;
        }

        private static string GenerateNewRelationshipId<P>(P parentPart) where P : OpenXmlPackaging.OpenXmlPart
        {
            int index = 1;
            string relationshipIdFormatString = "rId{0}";            
            while (OpenXmlUtilities.GetPartById(parentPart, string.Format(relationshipIdFormatString, index)) != null)
            {
                index++;
            }
            return string.Format(relationshipIdFormatString, index);
        }

        public static T RecreatePart<T>(T openXmlPart, out IEnumerable<KeyValuePair<string, string>> namespaceDeclarations)
            where T : OpenXmlPackaging.OpenXmlPart, OpenXmlPackaging.IFixedContentTypePart
        {
            if (openXmlPart.RootElement != null)
                namespaceDeclarations = openXmlPart.RootElement.NamespaceDeclarations;
            else
                namespaceDeclarations = null;

            // Get part parent
            IEnumerable<OpenXmlPackaging.OpenXmlPart> parentParts = openXmlPart.GetParentParts();
            if (parentParts.Count() != 1)
                throw new Exception("Number of part parents not equal to 1. Needs to be handled in the code");
            OpenXmlPackaging.OpenXmlPart parentPart = parentParts.First();

            // Get relationship id between parent and part
            string id = parentPart.GetIdOfPart(openXmlPart);

            // Change id to allow to new part to be created with existing id
            parentPart.ChangeIdOfPart(openXmlPart, id + "original");

            // Get part children
            Dictionary<string, OpenXmlPackaging.OpenXmlPart> childParts = new Dictionary<string, OpenXmlPackaging.OpenXmlPart>();
            foreach (OpenXmlPackaging.IdPartPair idPart in openXmlPart.Parts)
            {
                childParts.Add(idPart.RelationshipId, idPart.OpenXmlPart);
            }

            bool hasChildren = (childParts.Count() > 0);
            
            // Delete old part if no children. 
            // This is done here as some parts (SharedStringTable) can only exist once, so need to delete before creating new part.
            // This will need a re-think if there are parts that can only exist once and have child parts
            if (!hasChildren)
                parentPart.DeletePart(openXmlPart);                
            
            // Create new part
            T newPart = parentPart.AddNewPart<T>(id);

            // Copy existing part children to new part. Need to create relationships to new part before deleting the old part. 
            // Deleting a part seems to delete the unique relationships it has.
            if (hasChildren)
            {                
                foreach (KeyValuePair<string, OpenXmlPackaging.OpenXmlPart> childPart in childParts)
                {
                    newPart.CreateRelationshipToPart(childPart.Value, childPart.Key);
                }

                parentPart.DeletePart(openXmlPart);
            }

            return newPart;
        }

        public static OpenXmlPackaging.OpenXmlPart GetPartById(OpenXmlPackaging.OpenXmlPart parentPart, string id)
        {
            // GetPartById throws an exception when attempting to get an invalid id.
            foreach (OpenXmlPackaging.IdPartPair childPartPair in parentPart.Parts)
            {                
                if (string.Compare(childPartPair.RelationshipId, id, StringComparison.InvariantCultureIgnoreCase) == 0)
                    return childPartPair.OpenXmlPart;
            }

            return null;
        }

        /// <summary>
        /// Recursively find the part with the given uri.
        /// </summary>
        public static OpenXmlPackaging.OpenXmlPart GetPartByUri(OpenXmlPackaging.OpenXmlPart parentPart, string uri)
        {            
            foreach (OpenXmlPackaging.IdPartPair childPartPair in parentPart.Parts)
            {
                if (string.Compare(childPartPair.OpenXmlPart.Uri.ToString(), uri, StringComparison.InvariantCultureIgnoreCase) == 0)
                    return childPartPair.OpenXmlPart;
                else
                {
                    OpenXmlPackaging.OpenXmlPart part = GetPartByUri(childPartPair.OpenXmlPart, uri);
                    if (part != null)
                        return part;
                }
            }

            return null;
        }

        public static string GetAttributeValueFromReader(CustomOpenXmlReader reader, string name)
        {
            CustomOpenXmlAttribute attribute = reader.Attributes[name];
            if (attribute != null)
                return attribute.Value;
            return "";
        }
    }    
}
