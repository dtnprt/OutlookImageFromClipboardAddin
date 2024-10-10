using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OutlookImageFromClipboardAddin.Properties;
using OutlookImageFromClipboardAddin;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Clippy
{
    public partial class ImageForm : Form
    {
        string FolderPath = null;
        string Filename = null;
        string FileExtension = ".png";
        Image CurrentImage = null;
        public string FilePath = null;
        public string AttachmentName = null;

        protected override void OnHandleCreated(EventArgs e)
        {
            WindowsApiHelper.UseImmersiveDarkMode(Handle, WindowsApiHelper.IsDarkMode());
        }

        public ImageForm()
        {
            InitializeComponent();
            this.Icon = Icon.FromHandle(Resources.insert_image.GetHicon());
            if (WindowsApiHelper.IsDarkMode())
            {
                WindowsApiHelper.SetWindowTheme(this.Handle, "DarkMode_Explorer", null);

                this.BackColor = CustomColorTable.dark2;
                this.ForeColor = CustomColorTable.light1;
                pictureBox1.BackgroundImage = GenerateBackgroundTiles(16, CustomColorTable.dark3);
                pictureBox1.BackColor = CustomColorTable.dark2;

                foreach (Control c in this.Controls)
                {
                    c.BackColor = CustomColorTable.dark2;
                    c.ForeColor = CustomColorTable.light1;
                    if(c.GetType() == typeof(System.Windows.Forms.Button))
                       ((System.Windows.Forms.Button)c).FlatAppearance.BorderColor = CustomColorTable.dark3;
                }
                txtFilename.BackColor = CustomColorTable.dark3;
            }
        }

        private void ImageForm_Load(object sender, EventArgs e)
        {



            FolderPath = GuaranteeBackslash(Path.GetTempPath()) + "Clippy-Images\\";
            Filename = GetFilepath();
            txtFilename.Text = Filename;
         
            if (Clipboard.ContainsImage())
            {
                CurrentImage = Clipboard.GetImage();
                pictureBox1.Image = CurrentImage;
                lblInfo.Text = $"Size: {CurrentImage.Size.Width.ToString()}x{CurrentImage.Size.Height.ToString()}px";
            }
            txtFilename.Focus();
        }

        string GetFilepath()
        {
            int counter = 0;
            string date = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
            
            string tmpFileName = "Image-" + date;
            
            string fullPath = GuaranteeBackslash(this.FolderPath) + tmpFileName;
            string fullPathWithExtension = fullPath + FileExtension;
            do
            {
                fullPathWithExtension = fullPath + "-" + counter++.ToString() + FileExtension;
            } while (File.Exists(fullPathWithExtension));

            return tmpFileName + FileExtension;
        }

        string GuaranteeBackslash(string Path)
        {
            return Path.EndsWith("\\") ? Path : Path + "\\";
        }

        string GuaranteeExtension(string Path, string Extension)
        {
            return Path.EndsWith(Extension) ? Path : Path + Extension;
        }


        void SaveImage()
        {
            FilePath = GuaranteeBackslash(FolderPath) + GuaranteeExtension(txtFilename.Text, FileExtension);

            System.IO.Directory.CreateDirectory(this.FolderPath);
            
            pictureBox1.Image.Save(FilePath, System.Drawing.Imaging.ImageFormat.Png);
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveImage();
        }

        public static Bitmap GenerateBackgroundTiles(int size, Color col)
        {
            Bitmap bmp = new Bitmap(size * 2, size * 2);
            using (SolidBrush brush = new SolidBrush(col))
            using (Graphics G = Graphics.FromImage(bmp))
            {
                G.FillRectangle(brush, 0, 0, size, size);
                G.FillRectangle(brush, size, size, size, size);
            }

            return bmp;
        }

    }
}
