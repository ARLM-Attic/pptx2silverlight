using System;
using System.Xml;
using System.IO;

namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    class XAMLWriter
    {
        private XmlDocument xamlDoc;
        private SilverLight silverLight;

        public XAMLWriter(SilverLight silverLight)
        {
            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml");

            this.xamlDoc = new XmlDocument(nt);
            this.silverLight = silverLight;
        }

        private string XamlPage
        {
            get
            {
                return silverLight.XamlFile;
            }
        }

        public void InsertPPTTitle(string msg, int Width, int Height, int CanvasLeft, int CanvasTop)
        {
            XmlNode canvasNode;
            XmlNode childNode;
            XmlAttribute attrib;

            msg = msg.Replace('\n', ' ');

            xamlDoc.Load(XamlPage);
            canvasNode = xamlDoc.GetElementsByTagName("Canvas")[0];

            childNode = xamlDoc.CreateNode(XmlNodeType.Element, "TextBlock", "http://schemas.microsoft.com/client/2007");

            attrib = xamlDoc.CreateAttribute("Text");
            attrib.Value = msg;
            childNode.Attributes.Append(attrib);

            attrib = xamlDoc.CreateAttribute("Width");
            attrib.Value = Width.ToString();
            childNode.Attributes.Append(attrib);

            attrib = xamlDoc.CreateAttribute("Height");
            attrib.Value = Height.ToString();
            childNode.Attributes.Append(attrib);

            attrib = xamlDoc.CreateAttribute("Canvas.Left");
            attrib.Value = CanvasLeft.ToString();
            childNode.Attributes.Append(attrib);

            attrib = xamlDoc.CreateAttribute("Canvas.Top");
            attrib.Value = CanvasTop.ToString();
            childNode.Attributes.Append(attrib);

            canvasNode.AppendChild(childNode);

            xamlDoc.Save(XamlPage);
        }

        public void UpdateSlideImageInfo(SlideImagesInfo slideImagesInfo)
        {
            XmlNode childNode;
            XmlAttribute attrib;
            XmlNodeList childList;

            xamlDoc.Load(XamlPage);

            childList = xamlDoc.GetElementsByTagName("Image");
            for (int i = 0; i < childList.Count; i++)
            {
                childNode = childList[i];
                attrib = childNode.Attributes["x:Name"];
                if (attrib.Value == "BackgroundImage")
                {
                    childNode.Attributes["Source"].Value = slideImagesInfo.SlideImagePath + "/.." + Path.DirectorySeparatorChar + "Background.jpg";
                    childNode.Attributes["Source"].Value = childNode.Attributes["Source"].Value.Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    childNode.Attributes["Source"].Value = "file:///" + childNode.Attributes["Source"].Value;
                }
                else if (attrib.Value == "PptSlideImage")
                {
                    childNode.Attributes["Source"].Value = slideImagesInfo.SlideImagePath + Path.DirectorySeparatorChar + "Slide1.jpg";
                    childNode.Attributes["Source"].Value = childNode.Attributes["Source"].Value.Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    childNode.Attributes["Source"].Value = "file:///" + childNode.Attributes["Source"].Value;
                }
                else
                    continue;
            }

            childList = xamlDoc.GetElementsByTagName("TextBlock");
            for (int i = 0; i < childList.Count; i++)
            {
                childNode = childList[i];
                attrib = childNode.Attributes["x:Name"];
                if (attrib.Value == "SlidesCount")
                    childNode.Attributes["Text"].Value = slideImagesInfo.SlidesCount.ToString();
                else if (attrib.Value == "FullScreen")
                    childNode.Attributes["Text"].Value = "Full Screen";
                else if (attrib.Value == "SlideNumber")
                    childNode.Attributes["Text"].Value = string.Format("{0} Of {1}", 1, slideImagesInfo.SlidesCount);
                else
                    continue;
            }

            xamlDoc.Save(XamlPage);

            //Change the title of the Html page
            StreamReader sr = new StreamReader(silverLight.HtmlFile);
            string htmlText = sr.ReadToEnd();
            htmlText = htmlText.Replace("Silverlight Project Test Page", slideImagesInfo.PptName);
            sr.Close();
            StreamWriter sw = new StreamWriter(silverLight.HtmlFile);
            sw.Write(htmlText);
            sw.Flush();
            sw.Close();
        }
    }
}
