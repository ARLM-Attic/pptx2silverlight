using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.Threading;

namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    public partial class PowerPointReaderForm : Form
    {
        private PowerPointReader reader = new PowerPointReader();
        private delegate string HandleRequestDelegate(string pptFileName, StringBuilder sb);
        private delegate void HandleProgressBarDelegate();

        public PowerPointReaderForm()
        {
            InitializeComponent();
            SetProgressBarInitialValues();
        }

        private void HandleRequest(HandleRequestDelegate handle)
        {
            try
            {
                if (string.IsNullOrEmpty(this.txtPPTFileName.Text) || !File.Exists(this.txtPPTFileName.Text))
                {
                    MessageBox.Show("Please select a valid pptx file.");
                    btnOpenPPTFile.Focus();
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    SetProgressBarInitialValues();

                    pptConvertorProgressBar.Show();
                    ThreadPool.QueueUserWorkItem(new WaitCallback(HandleProgressBar));
                    string htmlFile = handle(this.txtPPTFileName.Text, sb);
                    pptConvertorProgressBar.Hide();

                    Utility.StarProcess(htmlFile);
                }
            }
            catch (Exception ex)
            {
                SetProgressBarInitialValues();
                Utility.ShowException(ex);
            }
        }

        private void HandleProgressBar(object o)
        {
            for (; pptConvertorProgressBar.Value <= pptConvertorProgressBar.Maximum; Thread.Sleep(100))
            {
                if (!pptConvertorProgressBar.Visible)
                    break;
                pptConvertorProgressBar.Value = (pptConvertorProgressBar.Value == pptConvertorProgressBar.Maximum ? pptConvertorProgressBar.Minimum : pptConvertorProgressBar.Value + 1);
            }
        }

        private void SetProgressBarInitialValues()
        {
            this.pptConvertorProgressBar.Hide();
            this.pptConvertorProgressBar.Value = this.pptConvertorProgressBar.Minimum;
        }

        private void BtnOpenPPTFile_Click(object sender, EventArgs e)
        {
            DialogResult result = this.openPPTFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                this.txtPPTFileName.Text = this.openPPTFileDialog.FileName;
            }
        }

        private void BtnConvertToSilverLight_Click(object sender, EventArgs e)
        {
            HandleRequest(new HandleRequestDelegate(reader.ConvertToSilverLightTitles));
        }

        private void BtnConvertToSilverLightImages_Click(object sender, EventArgs e)
        {
            HandleRequest(new HandleRequestDelegate(reader.ConvertToSilverLightImages));
        }
    }
}