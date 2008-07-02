using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

using System.Windows.Interop;

namespace HCLT.MSFT.TIL.SilverlightConvertor
{
    public partial class Page : Canvas
    {
        private string FILE_PREFIX = @"file:///";
        private string SLIDES_PATH = null;
        private int count = 1;

        public void Page_Loaded(object o, EventArgs e)
        {
            // Required to initialize variables
            InitializeComponent();
            InitializeImages();
        }

        private void InitializeImages()
        {
            try
            {
                string ImagePath = PptSlideImage.Source.ToString();
                SLIDES_PATH = ImagePath.Replace(FILE_PREFIX, "");
                SLIDES_PATH = FILE_PREFIX + System.IO.Path.GetDirectoryName(SLIDES_PATH);
                SLIDES_PATH = SLIDES_PATH.Replace(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);
                PptSlideImage.MouseLeftButtonUp += new MouseEventHandler(PptSlideImage_MouseLeftButtonUp);

                FullScreen.MouseLeftButtonUp += new MouseEventHandler(FullScreen_MouseLeftButtonUp);
                BrowserHost.FullScreenChange += new EventHandler(BrowserHost_FullScreenChange);

                tbPrev.MouseLeftButtonUp += new MouseEventHandler(tbPrev_MouseLeftButtonUp);
                tbNext.MouseLeftButtonUp += new MouseEventHandler(tbNext_MouseLeftButtonUp);
            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
        }

        void tbNext_MouseLeftButtonUp(object sender, MouseEventArgs e)
        {
            PptSlideImage.Source = new Uri(string.Format("{0}/Slide{1}.jpg", SLIDES_PATH, UpdateCount(1)));
        }

        void tbPrev_MouseLeftButtonUp(object sender, MouseEventArgs e)
        {
            PptSlideImage.Source = new Uri(string.Format("{0}/Slide{1}.jpg", SLIDES_PATH, UpdateCount(-1)));
        }

        void BrowserHost_FullScreenChange(object sender, EventArgs e)
        {
            AlterFullScreenText();
        }

        void FullScreen_MouseLeftButtonUp(object sender, MouseEventArgs e)
        {
            BrowserHost.IsFullScreen = !BrowserHost.IsFullScreen;
        }

        
        private const string NormalScreenText = "Full Screen";
        private const string FullScreenText = "Normal Mode";
        private void AlterFullScreenText()
        {
            switch(FullScreen.Text)
            {
                case Page.NormalScreenText:
                    FullScreen.Text = Page.FullScreenText;
                    break;
                case Page.FullScreenText:
                    FullScreen.Text = Page.NormalScreenText;
                    break;
            }
        }

        void PptSlideImage_MouseLeftButtonUp(object sender, MouseEventArgs e)
        {
            PptSlideImage.Source = new Uri(string.Format("{0}/Slide{1}.jpg", SLIDES_PATH, UpdateCount(1)));
        }

        private int UpdateCount(int x)
        {
            count += x;
            int maxCount = int.Parse(SlidesCount.Text);
            if (count == maxCount + 1)
                count = 1;
            else if (count == 0)
                count = maxCount;
            SlideNumber.Text = string.Format("{0} Of {1}", count, maxCount);
            return count;
        }
        
    }
}
