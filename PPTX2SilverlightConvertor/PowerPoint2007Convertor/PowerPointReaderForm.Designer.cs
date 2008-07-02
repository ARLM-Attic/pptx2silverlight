namespace HCLT.MSFT.TIL.PPT2007Convertor
{
    partial class PowerPointReaderForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PowerPointReaderForm));
            this.openPPTFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.txtPPTFileName = new System.Windows.Forms.TextBox();
            this.btnOpenPPTFile = new System.Windows.Forms.Button();
            this.btnConvertToSilverLightImages = new System.Windows.Forms.Button();
            this.pptConvertorProgressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // openPPTFileDialog
            // 
            this.openPPTFileDialog.Filter = "PowerPoint 2007 Files(*.pptx)|*.pptx";
            this.openPPTFileDialog.Title = "Select a PPT File";
            // 
            // txtPPTFileName
            // 
            this.txtPPTFileName.BackColor = System.Drawing.Color.NavajoWhite;
            this.txtPPTFileName.Location = new System.Drawing.Point(12, 12);
            this.txtPPTFileName.Name = "txtPPTFileName";
            this.txtPPTFileName.Size = new System.Drawing.Size(335, 20);
            this.txtPPTFileName.TabIndex = 0;
            // 
            // btnOpenPPTFile
            // 
            this.btnOpenPPTFile.Location = new System.Drawing.Point(360, 12);
            this.btnOpenPPTFile.Name = "btnOpenPPTFile";
            this.btnOpenPPTFile.Size = new System.Drawing.Size(93, 23);
            this.btnOpenPPTFile.TabIndex = 1;
            this.btnOpenPPTFile.Text = "Open PPTX File";
            this.btnOpenPPTFile.UseVisualStyleBackColor = true;
            this.btnOpenPPTFile.Click += new System.EventHandler(this.BtnOpenPPTFile_Click);
            // 
            // btnConvertToSilverLightImages
            // 
            this.btnConvertToSilverLightImages.Location = new System.Drawing.Point(12, 38);
            this.btnConvertToSilverLightImages.Name = "btnConvertToSilverLightImages";
            this.btnConvertToSilverLightImages.Size = new System.Drawing.Size(131, 23);
            this.btnConvertToSilverLightImages.TabIndex = 5;
            this.btnConvertToSilverLightImages.Text = "Convert To SilverLight";
            this.btnConvertToSilverLightImages.UseVisualStyleBackColor = true;
            this.btnConvertToSilverLightImages.Click += new System.EventHandler(this.BtnConvertToSilverLightImages_Click);
            // 
            // pptConvertorProgressBar
            // 
            this.pptConvertorProgressBar.Location = new System.Drawing.Point(154, 38);
            this.pptConvertorProgressBar.Name = "pptConvertorProgressBar";
            this.pptConvertorProgressBar.Size = new System.Drawing.Size(299, 23);
            this.pptConvertorProgressBar.TabIndex = 6;
            // 
            // PowerPointReaderForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Wheat;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(455, 80);
            this.Controls.Add(this.pptConvertorProgressBar);
            this.Controls.Add(this.btnConvertToSilverLightImages);
            this.Controls.Add(this.btnOpenPPTFile);
            this.Controls.Add(this.txtPPTFileName);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(471, 116);
            this.MinimumSize = new System.Drawing.Size(471, 116);
            this.Name = "PowerPointReaderForm";
            this.Text = "PPT Convertor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openPPTFileDialog;
        private System.Windows.Forms.TextBox txtPPTFileName;
        private System.Windows.Forms.Button btnOpenPPTFile;
        private System.Windows.Forms.Button btnConvertToSilverLightImages;
        private System.Windows.Forms.ProgressBar pptConvertorProgressBar;

    }
}

