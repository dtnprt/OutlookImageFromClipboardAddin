using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Text.RegularExpressions;

namespace OutlookImageFromClipboardAddin
{
    public partial class Ribbon1
    {

        Clippy.ImageForm imgForm = new Clippy.ImageForm();


        private void btnAddImageFromClipboard_Click(object sender, RibbonControlEventArgs e)
        {
            string Filename = null;

            if (Clipboard.ContainsImage())
            {
                
                if(imgForm.ShowDialog() == DialogResult.OK)
                {
                    Filename = imgForm.FilePath;

                    // Get the Application object
                    Outlook.Application application = Globals.ThisAddIn.Application;

                    // Get the active Inspector object and check if is type of MailItem
                    Outlook.Inspector inspector = application.ActiveInspector();
                    Outlook.Explorer explorer = application.ActiveExplorer();
                    Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
                    if (mailItem != null)
                    {
                        // make sure a filename was passed
                        if (string.IsNullOrEmpty(Filename) == false)
                        {
                            // need to check to see if file exists before we attach !
                            if (!File.Exists(Filename))
                                MessageBox.Show("Attached document " + Filename + " does not exist", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            else
                            {
                                Outlook.Attachment attachment = mailItem.Attachments.Add(Filename, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                                File.Delete(Filename);
                            }
                        }
                    }
                }

            }
            else
            {
                MessageBox.Show("No image found in clipboard.","Image error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
