using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Permissions;
using Microsoft.Win32;
using System.Diagnostics;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Control = System.Windows.Forms.Control;
using Font = System.Drawing.Font;
using System.IO;
using OpenXmlPowerTools;
using A = DocumentFormat.OpenXml.Drawing;


namespace PowerPointCreator
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    public partial class Form1 : Form
    {
        private int currentImage = 0;
        private List<Image> imageList1 = new List<Image>();
        private List<Image> selectedImages = new List<Image>();
        private int numberSlides = 0;
        private List<string> slideSources = new List<string>();
        public Form1()
        {
            //*Retreived from RooiWillie https://stackoverflow.com/questions/17922308/use-latest-version-of-internet-explorer-in-the-webbrowser-control *//
            int BrowserVer, RegVal;

            // get the installed IE version
            using (WebBrowser Wb = new WebBrowser())
                BrowserVer = Wb.Version.Major;

            // set the appropriate IE version
            if (BrowserVer >= 11)
                RegVal = 11001;
            else if (BrowserVer == 10)
                RegVal = 10001;
            else if (BrowserVer == 9)
                RegVal = 9999;
            else if (BrowserVer == 8)
                RegVal = 8888;
            else
                RegVal = 7000;

            // set the actual key
            using (RegistryKey Key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", RegistryKeyPermissionCheck.ReadWriteSubTree))
                if (Key.GetValue(System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe") == null)
                    Key.SetValue(System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe", RegVal, RegistryValueKind.DWord);
            //*//

            this.KeyPreview = true;
            this.Controls.Add(pictureBox1);
            this.Controls.Add(listView1);

            InitializeComponent();

            listView1.View = View.LargeIcon;
            imageList2.ImageSize = new Size(64, 64);
            listView1.LargeImageList = imageList2;
            Properties.Settings.Default.Folder_Path = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            folderBrowser.SelectedPath = Properties.Settings.Default.Folder_Path;
            listView2.View = View.Details;
            listView2.Columns.Add("Created Slides");
            listView2.Columns[0].Width = -2;
            listView2.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);
            listView2.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.HeaderSize);
            this.Scale(new SizeF((float).75, (float).75));
            fileBrowserTemplate.InitialDirectory = Properties.Settings.Default.Folder_Path;
            saveNewSlide.InitialDirectory = Properties.Settings.Default.Folder_Path;
            saveFinalPresentation.InitialDirectory = Properties.Settings.Default.Folder_Path;
        }

        public void BoldSelectedText(RichTextBox textbox)
        {
            //Stores initial selection
            int selectionStart = textbox.SelectionStart;
            int selectionLength = textbox.SelectionLength;
            //Sets the selected font to bold
            textbox.SelectionFont = new Font(textbox.Font, FontStyle.Bold);
            if (selectionLength != 0)
            {
                //Ensures text after selection isn't bolded
                textbox.SelectionStart = textbox.SelectionStart + textbox.SelectionLength;
                textbox.SelectionLength = 0;
                textbox.SelectionFont = textbox.Font;
                //Resets selected position
                textbox.Select(selectionStart, selectionLength);
            }
        }

        public void UnboldSelectedText(RichTextBox textbox)
        {
            //Stores initial selection
            int selectionStart = textbox.SelectionStart;
            int selectionLength = textbox.SelectionLength;
            //Sets the selected font to bold
            textbox.SelectionFont = new Font(textbox.Font, FontStyle.Regular);
            if (selectionLength != 0)
            {
                //Ensures text after selection isn't bolded
                textbox.SelectionStart = textbox.SelectionStart + textbox.SelectionLength;
                textbox.SelectionLength = 0;
                textbox.SelectionFont = textbox.Font;
                //Resets selected position
                textbox.Select(selectionStart, selectionLength);
            }
        }

        //Code retreived from Hinek https://stackoverflow.com/questions/435433/what-is-the-preferred-way-to-find-focused-control-in-winforms-app/439606#439606
        public Control FindFocusedControl(Control control)
        {
            var container = control as IContainerControl;
            while (container != null)
            {
                control = container.ActiveControl;
                container = control as IContainerControl;
            }
            return control;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.Control | Keys.B:
                    var focusedObj = FindFocusedControl(this);
                    try
                    {
                        if (((RichTextBox)focusedObj).SelectionFont.Bold.Equals(false))
                        {
                            BoldSelectedText((RichTextBox)focusedObj);
                        }
                        else if (((RichTextBox)focusedObj).SelectionFont.Bold.Equals(true))
                        {
                            UnboldSelectedText((RichTextBox)focusedObj);
                        }
                    }
                    catch (InvalidCastException)
                    {
                        Console.WriteLine("Cannot Cast Obj to RichTextBox");
                    }
                    return true;
                default:
                    break;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private List<string> splitRichText(RichTextBox rtbInput)
        {
            rtbInput.Enabled = false;
            List<string> boldWordsList = new List<string>();
            string boldWord = "";
            rtbInput.SelectionLength = 1;
            for (rtbInput.SelectionStart = 0; rtbInput.SelectionStart < rtbInput.TextLength; rtbInput.SelectionStart++)
            {
                if (rtbInput.SelectionFont.Style == FontStyle.Bold && char.IsLetterOrDigit(rtbInput.Text[rtbInput.SelectionStart]))
                {
                    boldWord += rtbInput.Text[rtbInput.SelectionStart];
                }
                else if (!char.IsLetterOrDigit(rtbInput.Text[rtbInput.SelectionStart]) && boldWord != "")
                {
                    boldWordsList.Add(boldWord);
                    boldWord = "";
                }
            }
            if (boldWord != "")
            {
                boldWordsList.Add(boldWord);
            }
            rtbInput.Enabled = true;
            rtbInput.Focus();
            return boldWordsList;
        }

        private void ImportHTMLImages()
        {
            if (imageBrowser.Document != null)
            {
                HtmlElementCollection imageList = null;
                HtmlDocument doc = imageBrowser.Document;
                imageList = doc.GetElementsByTagName("img");
                int i = 0;
                foreach (HtmlElement img in imageList)
                {
                    if (i > 8)
                    {
                        try
                        {
                            string[] src = Regex.Split(img.OuterHtml, "src=\"//");
                            src = Regex.Split(src[1], "\" data-");
                            string url = "https://" + src[0];
                            Image loadedImg = LoadImage(url);
                            imageList1.Add(loadedImg);
                        }
                        catch (System.IndexOutOfRangeException)
                        {

                        }
                    }
                    i++;
                }
            }
            pictureBox1.Image = imageList1[currentImage];
            pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
        }

        private Image LoadImage(string url)
        {
            System.Net.WebRequest request =
            System.Net.WebRequest.Create(url);

            System.Net.WebResponse response = request.GetResponse();
            System.IO.Stream responseStream =
                response.GetResponseStream();

            Bitmap bmp = new Bitmap(responseStream);

            responseStream.Dispose();

            return bmp;
        }

        private void updateSearch()
        {
            try
            {
                imageBrowser = null;
                imageBrowser = new System.Windows.Forms.WebBrowser();
                imageBrowser.ScriptErrorsSuppressed = true;
                string url = "https://www.duckduckgo.com/?q=";
                List<string> boldWordsBody;
                boldWordsBody = splitRichText(richTextBoxBody);
                var searchTermsTitle = richTextBoxTitle.Text.Split(' ');

                for (int i = 0; i < searchTermsTitle.Length; i++)
                {
                    url += searchTermsTitle[i] + "+";
                }
                for (int i = 0; i < boldWordsBody.Count; i++)
                {
                    url += boldWordsBody[i] + "+";
                }
                url += "&iax=images&ia=images";

                imageBrowser.Navigate(url);
                imageList1.Clear();
            }
            catch (System.TimeoutException)
            {

            }
        }

        private async void searchButton_Click(object sender, EventArgs e)
        {
            updateSearch();
            await imageBrowser.WaitForElementAsync("tile--img__img", 50000);
            ImportHTMLImages();
            try
            {
                if (selectedImages.Contains(imageList1[0]))
                {
                    checkBoxImageSelected.Checked = true;
                }
                else
                {
                    checkBoxImageSelected.Checked = false;
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("No images");
            }
        }

        private void prevButton_Click(object sender, EventArgs e)
        {
            if (imageList1.Count == 0)
            {
                return;
            }
            if (0 < currentImage)
            {
                currentImage--;
            }
            else
            {
                currentImage = imageList1.Count - 1;
            }

            pictureBox1.Image = imageList1[currentImage];
            pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;

            if (selectedImages.Contains(pictureBox1.Image))
            {
                checkBoxImageSelected.Checked = true;
            }
            else
            {
                checkBoxImageSelected.Checked = false;
            }
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            if (imageList1.Count == 0)
            {
                return;
            }
            if (imageList1.Count - 1 > currentImage)
            {
                currentImage++;
            }
            else
            {
                currentImage = 0;
            }

            pictureBox1.Image = imageList1[currentImage];
            pictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;

            if (selectedImages.Contains(pictureBox1.Image))
            {
                checkBoxImageSelected.Checked = true;
            }
            else
            {
                checkBoxImageSelected.Checked = false;
            }
        }

        private void ParseTemplate(List<A.Text> textBoxList)
        {
            foreach (A.Text textBox in textBoxList)
            {
                if (textBox.Text.ToLower().Contains("Title".ToLower()))
                {
                    textBox.Text = richTextBoxTitle.Text;
                }
                else if (textBox.Text.ToLower().Contains("Body".ToLower()))
                {
                    textBox.Text = richTextBoxBody.Text;
                }
            }
        }

        private void SetTemplateImages(Slide templateSlide)
        {
            int i = 0;
            var imagesList = templateSlide.SlidePart.ImageParts.OrderBy(o => o.Uri.ToString()).ToList();
            var columnsList = templateSlide.Descendants<GridColumn>().ToList();
            foreach (ImagePart img in imagesList)
            {
                try
                {
                    System.IO.Stream data = GetBinaryDataStream(ConvertImageToBase64(selectedImages[i]));
                    img.FeedData(data);
                    i++;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    Console.WriteLine("No image to insert.");
                    columnsList[i].Remove();
                    i++;
                }
            }
        }

        private void SaveSlides(Slide templateSlide, MemoryStream stream)
        {
            string path = null;
            var result = saveNewSlide.ShowDialog();
            if (result == DialogResult.OK)
            {
                path = saveNewSlide.FileName;
                templateSlide.Save();
                stream.Seek(0, SeekOrigin.Begin);
                FileStream outStream = null;
                try
                {
                    outStream = File.Create(path);
                    stream.CopyTo(outStream);
                    slideSources.Add(path);
                    outStream.Close();
                    listView2.Items.Add(path);
                    listView2.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);
                }
                catch (System.IO.IOException)
                {
                    MessageBox.Show("Cannot access file\n" + path + "\nsince it is currently open.");
                }

                stream.Close();
                saveNewSlide.FileName = "";
            }
        }

        //Opens PowerPoint Template and copies it for use as a template.
        private void OpenTemplatePPT(MemoryStream stream, string templatePath)
        {
            PresentationDocument presDoc = null;
            Slide templateSlide = null;
            List<A.Text> textBoxes = null;
            stream = new MemoryStream();
            A.Offset tableOffset = null;
            DocumentFormat.OpenXml.Int64Value cellOffset = null;
            
            try { 

                //Sets the default folder path in the application settings
                Properties.Settings.Default.Folder_Path = templatePath;
                Properties.Settings.Default.Save();

                //Opens the template and copies it to a new filestream
                FileStream template = null;
                string path = null;
                try
                {
                    template = File.Open(templatePath, FileMode.Open, FileAccess.Read);


                    template.CopyTo(stream);
                    presDoc = PresentationDocument.Open(stream, true);

                    templateSlide = presDoc.PresentationPart.SlideParts.First().Slide;
                    textBoxes = templateSlide.Descendants<A.Text>().ToList();

                    //Parses the template for the "title" and "body" placeholders and replaces them
                    ParseTemplate(textBoxes);

                    path = System.IO.Path.GetDirectoryName(Application.ExecutablePath);

                    tableOffset = templateSlide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().First().Parent.Parent.Parent.Descendants<DocumentFormat.OpenXml.Drawing.Offset>().ToList()[0];
                    cellOffset = templateSlide.Descendants<DocumentFormat.OpenXml.Drawing.GridColumn>().First().Width;

                    //Replaces placeholder images with user-provided images
                    SetTemplateImages(templateSlide);

                    //Offsets the images based on how many are present
                    OffsetImageTable(tableOffset, cellOffset);

                    //Saves updates made to the slides
                    SaveSlides(templateSlide, stream);
                    template.Close();
                }
                catch (System.IO.IOException)
                {
                    MessageBox.Show("Cannot access file\n" + templatePath + "\nsince it is already open.");
                }

            }
            catch (System.Exception ex)
            {
                if (ex is System.InvalidOperationException || ex is System.ArgumentOutOfRangeException)
                {
                    MessageBox.Show("Template is not in a valid format.\nA template must follow the following format:\n\nTextbox: \"Title\"\nTextbox: \"Body\"\n\t__________________\nTable: \t|img|img|img|img|\n\t-----------------------");
                }
            }
        }

        private void OpenFileExplorer()
        {
            DialogResult result = fileBrowserTemplate.ShowDialog();
            if (result == DialogResult.OK)
            {
                MemoryStream stream = new MemoryStream();
                string templatePath = fileBrowserTemplate.FileName;
                numberSlides++;
                OpenTemplatePPT(stream, templatePath);
                fileBrowserTemplate.FileName = "";
            }
        }
        private void createPPTButton_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileExplorer();
            }
            catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException)
            {
                numberSlides--;
                MessageBox.Show("Cannot create slide", "", MessageBoxButtons.OK);
            }
        }

        private string ConvertImageToBase64(Image file)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                file.Save(memoryStream, file.RawFormat);
                byte[] imageBytes = memoryStream.ToArray();
                return Convert.ToBase64String(imageBytes);
            }
        }

        private void combineSlidesButton_Click(object sender, EventArgs e)
        {
            List<PmlDocument> slideMasters = new List<PmlDocument>();
            List<SlideSource> sources = new List<SlideSource>();

            DialogResult result = saveFinalPresentation.ShowDialog();

            if (result == DialogResult.OK)
            {
                string path = saveFinalPresentation.FileName;
                foreach (string source in slideSources)
                {
                    sources.Add(new SlideSource(new PmlDocument(source), 0, 1, false));
                }

                try
                {
                    PresentationBuilder.BuildPresentation(sources, path);
                    MessageBox.Show("Successfully created presentation at" + path + ".");
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    MessageBox.Show("No slides to combine.");
                }
                saveFinalPresentation.FileName = "";
            }
        }

        private void OffsetImageTable(A.Offset tableOffset, DocumentFormat.OpenXml.Int64Value cellOffset)
        {
            switch (selectedImages.Count)
            {
                //If count is 4 don't need to adjust offset
                case 3:
                    tableOffset.X.Value += cellOffset.Value / 2;
                    break;
                case 2:
                    tableOffset.X.Value += cellOffset.Value;
                    break;
                case 1:
                    tableOffset.X.Value += (cellOffset.Value * 3) / 2;
                    break;
                default:
                    break;
            }
        }

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (selectedImages.Count > 3 && checkBoxImageSelected.Checked == false)
            {
                MessageBox.Show("Slide can only contain 4 images.", "", MessageBoxButtons.OK);
                return;
            }
            if (selectedImages.Contains(pictureBox1.Image))
            {
                listView1.Items.RemoveAt(imageList2.Images.IndexOfKey(ConvertImageToBase64(pictureBox1.Image)));
                imageList2.Images.RemoveByKey(ConvertImageToBase64(pictureBox1.Image));
                selectedImages.Remove(pictureBox1.Image);
                checkBoxImageSelected.Checked = false;
                updateListView();
            }
            else
            {
                imageList2.Images.Add(ConvertImageToBase64(pictureBox1.Image), pictureBox1.Image);
                selectedImages.Add(pictureBox1.Image);
                checkBoxImageSelected.Checked = true;
                updateListView();
            }
        }

        private void updateListView()
        {
            listView1.Items.Clear();
            foreach (Image img in selectedImages)
            {
                ListViewItem item = new ListViewItem();
                item.ImageIndex = imageList2.Images.IndexOfKey(ConvertImageToBase64(img));
                listView1.Items.Add(item);
            }
        }

        private void listView1_ItemActivate(Object sender, EventArgs e)
        {
            ListViewItem selectedItem = listView1.SelectedItems[0];
            if (selectedImages.ElementAt(selectedItem.Index) == pictureBox1.Image)
            {
                checkBoxImageSelected.Checked = false;
            }
            selectedImages.RemoveAt(selectedItem.Index);
            imageList2.Images.RemoveAt(selectedItem.Index);
            selectedItem.Remove();
            updateListView();
        }

        private void listView2_ItemActivate(object sender, EventArgs e)
        {
            ListViewItem selectedItem = listView2.SelectedItems[0];
            DialogResult result = MessageBox.Show("Are you sure you want to delete this slide?\n\n" + selectedItem.Text, "Delete Slide", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                try
                {
                    File.Delete(selectedItem.Text);
                    selectedItem.Remove();
                    numberSlides--;
                }
                catch (IOException)
                {

                }
            }
        }
    }
    public static class WebBrowserExtensions
    {
        public static async Task<string> WaitForElementAsync(this WebBrowser wb,
            string attributeName, int timeout = 30000, int interval = 500)
        {
            try
            {
                var stopwatch = Stopwatch.StartNew();
                while (true)
                {
                    try
                    {
                        if (wb.Document != null)
                        {
                            var elements = wb.Document.GetElementsByTagName("img");
                            foreach (HtmlElement elem in elements)
                            {
                                string className = elem.GetAttribute("className");
                                if (className != null && className.Length != 0 && className.Contains(attributeName)) return className;
                            }
                        }
                    }
                    catch (System.NullReferenceException) { }
                    if (stopwatch.ElapsedMilliseconds > timeout) throw new TimeoutException();
                    await Task.Delay(interval);
                }
            }
            catch (System.TimeoutException)
            {
                return "";
            }
        }
    }
}