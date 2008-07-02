using System;
using System.IO;
using System.Reflection;

using POWERPOINT = Microsoft.Office.Interop.PowerPoint;

namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    public class SilverLight
    {
        const string targetFolder = "SLVersion_{0}";
        private DirectoryInfo sourceDirInfo;
        private DirectoryInfo targetDirInfo;

        public string HtmlFile
        {
            get
            {
                return this.targetDirInfo.FullName + Path.DirectorySeparatorChar + "Default.html";
            }
        }

        public string XamlFile
        {
            get
            {
                return this.targetDirInfo.FullName + Path.DirectorySeparatorChar + "Page.xaml";
            }
        }

        public SilverLight(string pptFullPath, string templatePathKey)
        {
            CreateTargetDirectory(pptFullPath);

            Assembly assembly = Assembly.GetExecutingAssembly();

            sourceDirInfo = new DirectoryInfo(Path.GetDirectoryName(assembly.Location) + Path.DirectorySeparatorChar + templatePathKey);
            
            Copy(sourceDirInfo, targetDirInfo);
        }

        private void CreateTargetDirectory(string pptFullPath)
        {
            string targetDirectory = Path.GetDirectoryName(pptFullPath) + Path.DirectorySeparatorChar + string.Format(SilverLight.targetFolder, Path.GetFileName(pptFullPath));

            if (Directory.Exists(targetDirectory))
            {
                Directory.Delete(targetDirectory, true);
            }
            targetDirInfo = Directory.CreateDirectory(targetDirectory);
        }

        private void Copy(DirectoryInfo source, DirectoryInfo target)
        {
            foreach (FileInfo f in source.GetFiles())
            {
                File.Copy(f.FullName, target.FullName + Path.DirectorySeparatorChar + f.Name, true);
            }

            string targetdirName;
            foreach (DirectoryInfo d in source.GetDirectories())
            {
                targetdirName = target.FullName + Path.DirectorySeparatorChar + d.Name;
                if(!Directory.Exists(targetdirName))
                {
                    Directory.CreateDirectory(targetdirName);
                }
                Copy(d, new DirectoryInfo(targetdirName));
            }
        }

        public SlideImagesInfo CreateImages(POWERPOINT.Presentation pres)
        {
            //pres.SaveAs(pptFileName + @".xml", POWERPOINT.PpSaveAsFileType.ppSaveAsXMLPresentation, Microsoft.Office.Core.MsoTriState.msoTrue);

            string slideImagePath = this.targetDirInfo.FullName + Path.DirectorySeparatorChar + "Slides";
            pres.SaveAs(slideImagePath, POWERPOINT.PpSaveAsFileType.ppSaveAsJPG, Microsoft.Office.Core.MsoTriState.msoTrue);

            DirectoryInfo dInfo = new DirectoryInfo(slideImagePath);
            FileInfo [] filesInfo = dInfo.GetFiles();

            string pptFileName = Path.GetFileName(pres.FullName);
            SlideImagesInfo slideImagesInfo = new SlideImagesInfo(pptFileName, slideImagePath, filesInfo.Length);
            return slideImagesInfo;
        }
    }

    public class SlideImagesInfo
    {
        private string pptName;
        private string slideImagePath;
        private long slidesCount;
        public SlideImagesInfo(string pptName, string slideImagePath, int slidesCount)
        {
            this.pptName = pptName;
            this.slideImagePath = slideImagePath;
            this.slidesCount = slidesCount;
        }

        public string PptName
        {
            get
            {
                return this.pptName;
            }
        }

        public string SlideImagePath
        {
            get
            {
                return this.slideImagePath;
            }
        }

        public long SlidesCount
        {
            get
            {
                return this.slidesCount;
            }
        }
    }
}
