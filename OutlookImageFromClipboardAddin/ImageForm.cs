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
        public ImageForm()
        {
            InitializeComponent();
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

    }
}
