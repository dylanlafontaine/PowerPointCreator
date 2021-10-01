
namespace PowerPointCreator
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.richTextBoxBody = new System.Windows.Forms.RichTextBox();
            this.richTextBoxTitle = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.imageBrowser = new System.Windows.Forms.WebBrowser();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.prevButton = new System.Windows.Forms.Button();
            this.nextButton = new System.Windows.Forms.Button();
            this.groupBoxImages = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.checkBoxImageSelected = new System.Windows.Forms.CheckBox();
            this.axWebBrowser = new System.Windows.Forms.WebBrowser();
            this.openFileDialogPPT = new System.Windows.Forms.OpenFileDialog();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.createPPTButton = new System.Windows.Forms.Button();
            this.combineSlidesButton = new System.Windows.Forms.Button();
            this.imageList2 = new System.Windows.Forms.ImageList(this.components);
            this.listView2 = new System.Windows.Forms.ListView();
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.fileBrowserTemplate = new System.Windows.Forms.OpenFileDialog();
            this.saveNewSlide = new System.Windows.Forms.SaveFileDialog();
            this.saveFinalPresentation = new System.Windows.Forms.SaveFileDialog();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBoxImages.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupBox1.Controls.Add(this.searchButton);
            this.groupBox1.Controls.Add(this.richTextBoxBody);
            this.groupBox1.Controls.Add(this.richTextBoxTitle);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 58);
            this.groupBox1.MinimumSize = new System.Drawing.Size(1215, 308);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1246, 308);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Use Shortcut Ctrl + B to Bold Text";
            // 
            // searchButton
            // 
            this.searchButton.Location = new System.Drawing.Point(143, 235);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(1000, 37);
            this.searchButton.TabIndex = 4;
            this.searchButton.Text = "Search for Images";
            this.searchButton.UseVisualStyleBackColor = true;
            this.searchButton.Click += new System.EventHandler(this.searchButton_Click);
            // 
            // richTextBoxBody
            // 
            this.richTextBoxBody.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.richTextBoxBody.Location = new System.Drawing.Point(143, 108);
            this.richTextBoxBody.Name = "richTextBoxBody";
            this.richTextBoxBody.Size = new System.Drawing.Size(1000, 121);
            this.richTextBoxBody.TabIndex = 3;
            this.richTextBoxBody.Text = "";
            // 
            // richTextBoxTitle
            // 
            this.richTextBoxTitle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.richTextBoxTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
            this.richTextBoxTitle.Location = new System.Drawing.Point(143, 32);
            this.richTextBoxTitle.Name = "richTextBoxTitle";
            this.richTextBoxTitle.Size = new System.Drawing.Size(1000, 46);
            this.richTextBoxTitle.TabIndex = 2;
            this.richTextBoxTitle.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F);
            this.label2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.label2.Location = new System.Drawing.Point(6, 140);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(112, 46);
            this.label2.TabIndex = 1;
            this.label2.Text = "Body";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F);
            this.label1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.label1.Location = new System.Drawing.Point(23, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 46);
            this.label1.TabIndex = 0;
            this.label1.Text = "Title";
            // 
            // imageBrowser
            // 
            this.imageBrowser.Location = new System.Drawing.Point(888, 0);
            this.imageBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.imageBrowser.Name = "imageBrowser";
            this.imageBrowser.Size = new System.Drawing.Size(47, 37);
            this.imageBrowser.TabIndex = 1;
            this.imageBrowser.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(240, 21);
            this.pictureBox1.MaximumSize = new System.Drawing.Size(1112, 475);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(66, 34);
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // prevButton
            // 
            this.prevButton.Location = new System.Drawing.Point(252, 925);
            this.prevButton.Name = "prevButton";
            this.prevButton.Size = new System.Drawing.Size(98, 52);
            this.prevButton.TabIndex = 3;
            this.prevButton.Text = "Previous";
            this.prevButton.UseVisualStyleBackColor = true;
            this.prevButton.Click += new System.EventHandler(this.prevButton_Click);
            // 
            // nextButton
            // 
            this.nextButton.Location = new System.Drawing.Point(786, 925);
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(98, 52);
            this.nextButton.TabIndex = 4;
            this.nextButton.Text = "Next";
            this.nextButton.UseVisualStyleBackColor = true;
            this.nextButton.Click += new System.EventHandler(this.nextButton_Click);
            // 
            // groupBoxImages
            // 
            this.groupBoxImages.AutoSize = true;
            this.groupBoxImages.Controls.Add(this.label3);
            this.groupBoxImages.Controls.Add(this.listView1);
            this.groupBoxImages.Controls.Add(this.checkBoxImageSelected);
            this.groupBoxImages.Controls.Add(this.pictureBox1);
            this.groupBoxImages.Location = new System.Drawing.Point(12, 389);
            this.groupBoxImages.MinimumSize = new System.Drawing.Size(1215, 501);
            this.groupBoxImages.Name = "groupBoxImages";
            this.groupBoxImages.Size = new System.Drawing.Size(1246, 501);
            this.groupBoxImages.TabIndex = 5;
            this.groupBoxImages.TabStop = false;
            this.groupBoxImages.Text = "Click Images to Select or Deselect Them for a Slide";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(989, 38);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 17);
            this.label3.TabIndex = 10;
            this.label3.Text = "Selected Images";
            // 
            // listView1
            // 
            this.listView1.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(952, 58);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(191, 417);
            this.listView1.TabIndex = 9;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.ItemActivate += new System.EventHandler(this.listView1_ItemActivate);
            // 
            // checkBoxImageSelected
            // 
            this.checkBoxImageSelected.AutoCheck = false;
            this.checkBoxImageSelected.AutoSize = true;
            this.checkBoxImageSelected.Location = new System.Drawing.Point(44, 34);
            this.checkBoxImageSelected.Name = "checkBoxImageSelected";
            this.checkBoxImageSelected.Size = new System.Drawing.Size(127, 21);
            this.checkBoxImageSelected.TabIndex = 8;
            this.checkBoxImageSelected.Text = "Image Selected";
            this.checkBoxImageSelected.UseVisualStyleBackColor = true;
            // 
            // axWebBrowser
            // 
            this.axWebBrowser.Location = new System.Drawing.Point(954, 0);
            this.axWebBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.axWebBrowser.Name = "axWebBrowser";
            this.axWebBrowser.Size = new System.Drawing.Size(42, 37);
            this.axWebBrowser.TabIndex = 6;
            this.axWebBrowser.Visible = false;
            // 
            // openFileDialogPPT
            // 
            this.openFileDialogPPT.FileName = "openFileDialoguePPT";
            // 
            // createPPTButton
            // 
            this.createPPTButton.Location = new System.Drawing.Point(516, 925);
            this.createPPTButton.Name = "createPPTButton";
            this.createPPTButton.Size = new System.Drawing.Size(98, 52);
            this.createPPTButton.TabIndex = 7;
            this.createPPTButton.Text = "Create Slide";
            this.createPPTButton.UseVisualStyleBackColor = true;
            this.createPPTButton.Click += new System.EventHandler(this.createPPTButton_Click);
            // 
            // combineSlidesButton
            // 
            this.combineSlidesButton.Location = new System.Drawing.Point(252, 995);
            this.combineSlidesButton.Name = "combineSlidesButton";
            this.combineSlidesButton.Size = new System.Drawing.Size(632, 52);
            this.combineSlidesButton.TabIndex = 8;
            this.combineSlidesButton.Text = "Combine Slides";
            this.combineSlidesButton.UseVisualStyleBackColor = true;
            this.combineSlidesButton.Click += new System.EventHandler(this.combineSlidesButton_Click);
            // 
            // imageList2
            // 
            this.imageList2.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList2.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList2.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // listView2
            // 
            this.listView2.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.listView2.HideSelection = false;
            this.listView2.LabelWrap = false;
            this.listView2.Location = new System.Drawing.Point(1276, 58);
            this.listView2.MultiSelect = false;
            this.listView2.Name = "listView2";
            this.listView2.Size = new System.Drawing.Size(193, 832);
            this.listView2.TabIndex = 9;
            this.listView2.UseCompatibleStateImageBehavior = false;
            this.listView2.View = System.Windows.Forms.View.List;
            this.listView2.ItemActivate += new System.EventHandler(this.listView2_ItemActivate);
            // 
            // folderBrowser
            // 
            this.folderBrowser.Description = "Select folder to save slides to:";
            this.folderBrowser.SelectedPath = global::PowerPointCreator.Properties.Settings.Default.Folder_Path;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(671, 1066);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(0, 0);
            this.button1.TabIndex = 10;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(1588, 724);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(0, 0);
            this.button2.TabIndex = 11;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // fileBrowserTemplate
            // 
            this.fileBrowserTemplate.Filter = "Powerpoint Files|*.pptx";
            this.fileBrowserTemplate.InitialDirectory = global::PowerPointCreator.Properties.Settings.Default.Folder_Path;
            this.fileBrowserTemplate.RestoreDirectory = true;
            this.fileBrowserTemplate.Title = "Select a PowerPoint Template";
            // 
            // saveNewSlide
            // 
            this.saveNewSlide.Filter = "PowerPoint Files|*.pptx";
            this.saveNewSlide.Title = "Select a Location To Save Your Slide";
            // 
            // saveFinalPresentation
            // 
            this.saveFinalPresentation.Filter = "PowerPoint Files|*.pptx";
            this.saveFinalPresentation.Title = "Select a Location to Save Your Presentation";
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1618, 1150);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.listView2);
            this.Controls.Add(this.combineSlidesButton);
            this.Controls.Add(this.nextButton);
            this.Controls.Add(this.createPPTButton);
            this.Controls.Add(this.prevButton);
            this.Controls.Add(this.axWebBrowser);
            this.Controls.Add(this.groupBoxImages);
            this.Controls.Add(this.imageBrowser);
            this.Controls.Add(this.groupBox1);
            this.MinimumSize = new System.Drawing.Size(1636, 1028);
            this.Name = "Form1";
            this.Text = "Power Point Image Searcher";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBoxImages.ResumeLayout(false);
            this.groupBoxImages.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RichTextBox richTextBoxBody;
        private System.Windows.Forms.RichTextBox richTextBoxTitle;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.WebBrowser imageBrowser;
        private System.Windows.Forms.Button searchButton;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button prevButton;
        private System.Windows.Forms.Button nextButton;
        private System.Windows.Forms.GroupBox groupBoxImages;
        private System.Windows.Forms.WebBrowser axWebBrowser;
        private System.Windows.Forms.OpenFileDialog openFileDialogPPT;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button createPPTButton;
        private System.Windows.Forms.CheckBox checkBoxImageSelected;
        private System.Windows.Forms.Button combineSlidesButton;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ImageList imageList2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListView listView2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowser;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.OpenFileDialog fileBrowserTemplate;
        private System.Windows.Forms.SaveFileDialog saveNewSlide;
        private System.Windows.Forms.SaveFileDialog saveFinalPresentation;
    }
}

