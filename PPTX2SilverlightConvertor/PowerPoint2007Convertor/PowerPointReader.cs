using System;
using System.IO;
using System.IO.Packaging;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Windows.Forms;

using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;

namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    public static class PPTNameSpace
    {
        public readonly static string DocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public readonly static string PresentationmlNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    }

    public class PowerPointReader
    {
        public string ConvertToSilverLightImages(string pptFileName, StringBuilder sb)
        {
            POWERPOINT.Application app = new POWERPOINT.Application();
            POWERPOINT.Presentation pres = app.Presentations.Open(pptFileName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            SilverLight silverLight = new SilverLight(pptFileName, ConfigManager.SilverLightTemplatePath);
            XAMLWriter xamlWriter = new XAMLWriter(silverLight);

            SlideImagesInfo slideImagesInfo = silverLight.CreateImages(pres);
            xamlWriter.UpdateSlideImageInfo(slideImagesInfo);
            return silverLight.HtmlFile;
        }

        public string ConvertToSilverLightTitles(string pptFileName, StringBuilder sb)
        {
            SilverLight silverLight = new SilverLight(pptFileName, ConfigManager.SilverLightTemplatePath);
            XAMLWriter xamlWriter = new XAMLWriter(silverLight);

            //  Fill this collection with a list of all the titles of all the slides in the requested slide deck.
            List<string> titles = new List<string>();

            PackagePart documentPart = null;
            Uri documentUri = null;

            using (Package pptPackage = Package.Open(pptFileName, FileMode.Open, FileAccess.Read))
            {
                foreach (PackageRelationship relationship in pptPackage.GetRelationshipsByType(PPTNameSpace.DocumentRelationshipType))
                {
                    documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
                    documentPart = pptPackage.GetPart(documentUri);
                    break;
                }
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("p", PPTNameSpace.PresentationmlNamespace);

                // Iterate through the slides and extract the title string from each.
                XmlDocument xDoc = new XmlDocument(nt);
                xDoc.Load(documentPart.GetStream());

                XmlNodeList sheetNodes = xDoc.SelectNodes("//p:sldIdLst/p:sldId", nsManager);
                if (sheetNodes != null)
                {
                    XmlAttribute relAttr = null;
                    PackageRelationship sheetRelationship = null;
                    PackagePart sheetPart = null;
                    Uri sheetUri = null;
                    XmlDocument sheetDoc = null;
                    XmlNode titleNode = null;

                    //Look at each sheet node, retrieving the relationship id.
                    foreach (System.Xml.XmlNode xNode in sheetNodes)
                    {
                        relAttr = xNode.Attributes["r:id"];
                        if (relAttr != null)
                        {
                            //Retrieve the PackageRelationship object for the sheet.
                            sheetRelationship = documentPart.GetRelationship(relAttr.Value);
                            if (sheetRelationship != null)
                            {
                                sheetUri = PackUriHelper.ResolvePartUri(
                                  documentUri, sheetRelationship.TargetUri);
                                sheetPart = pptPackage.GetPart(sheetUri);
                                if (sheetPart != null)
                                {
                                    sheetDoc = new XmlDocument(nt);
                                    sheetDoc.Load(sheetPart.GetStream());
                                    titleNode = sheetDoc.SelectSingleNode("//p:sp//p:ph[@type='title' or @type='ctrTitle']", nsManager);
                                    if (titleNode != null)
                                    {
                                        titles.Add(titleNode.ParentNode.ParentNode.ParentNode.InnerText);
                                    }
                                }
                            }
                        }
                    }
                }

                pptPackage.Close();
            }

            for (int i = 0; i < titles.Count; i++)
            {
                Utility.Concat(sb, "Title of slide ", i+1, " is [", titles[i] + "]");
                xamlWriter.InsertPPTTitle(titles[i], 100, 20, 280, 40+ 12*i);
            }

            return silverLight.HtmlFile;
        }
    }
}