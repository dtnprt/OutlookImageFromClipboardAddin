using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
//using Word = Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace OutlookImageFromClipboardAddin
{
    public partial class ThisAddIn
    {

        Outlook.MailItem mailItem;
        public static string lastJob = string.Empty;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            /*
            // Get the Application object
            Outlook.Application application = this.Application;

            // Get the Inspector object
            Outlook.Inspectors inspectors = application.Inspectors;

            // Get the active Inspector object
            Outlook.Inspector activeInspector = application.ActiveInspector();
            if (activeInspector != null)
            {
                // Get the title of the active item when the Outlook start.
                //MessageBox.Show("Active inspector: " + activeInspector.Caption);
            }

            // Get the Explorer objects
            Outlook.Explorers explorers = application.Explorers;


            // Get the active Explorer object
            Outlook.Explorer activeExplorer = application.ActiveExplorer();
            if (activeExplorer != null)
            {
                // Get the title of the active folder when the Outlook start.
                //MessageBox.Show("Active explorer: " + activeExplorer.Caption);
            }


            // ...
            // Add a new Inspector to the application
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_AddTextToNewMail);
            

            application.ItemLoad += Application_ItemLoad;
            */


        }


        private void Application_ItemLoad(object Item)
        {
            /*
            if(Item.GetType() == typeof(Outlook.MailItem))
            {
                mailItem = (Outlook.MailItem)Item;
                ;
            }
            */
        }

        void Inspectors_AddTextToNewMail(Outlook.Inspector inspector)
        {
            /*
            this.mailItem = inspector.CurrentItem as Outlook.MailItem;
            
            
            if (mailItem != null)
            {
                ;
                if (mailItem.EntryID == null)
                {
                        //mailItem.Open += new Outlook.ItemEvents_10_OpenEventHandler(MailItem_Open);
                    //mailItem.Unload += MailItem_Unload;
                   
                }
            }
            */
        }

        private void MailItem_Unload()
        {
        }

        private void App_WindowSelectionChange(Word.Selection Sel)
        {

        }

        private void MailItem_PropertyChange(string Name)
        {
            
        }

        private void MailItem_Open(ref bool Cancel)
        {
            
           // mailItem.Open -= MailItem_Open;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }



        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
