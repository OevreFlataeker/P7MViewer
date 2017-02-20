using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Redemption;

namespace P7MViewer_Plugin
{
    public partial class ThisAddIn
    {

        private Office.CommandBar PacktOldMenuBar;
        // Defining old Menubar
        private Office.CommandBarPopup PacktNewMenuBar;
        // Defining instance of button for menu item
        private Office.CommandBarButton PacktButton1;
        // Tag string for our Menu item
        private string strMenuString = "Analyze PKCS#7 structure";

        Office.CommandBar PacktCustomToolBar;
        // Declare the button
        Office.CommandBarButton PacktButtonA;

        Outlook._Application olApp;
        Outlook.Explorer olExplorer;
        string mailID;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Define the Old Menu Bar
            PacktOldMenuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
            // Define the new Menu Bar into the existing menu bar
            PacktNewMenuBar = (Office.CommandBarPopup)PacktOldMenuBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, false);
            //If PacktNewMenuBar not found then the code will add it
            if (PacktNewMenuBar != null)
            {
                // Set caption for the Menu
                PacktNewMenuBar.Caption = "Analyze PKCS#7 structure";
                // Tag string value passing
                PacktNewMenuBar.Tag = strMenuString;
                // Assigning button type
                PacktButton1 = (Office.CommandBarButton)PacktNewMenuBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                // Setting up the button style
                PacktButton1.Style = Office.MsoButtonStyle.msoButtonIconAndCaptionBelow;
                // Set button caption
                PacktButton1.Caption = "Analyze PKCS#7 structure";
                // Set the menu visible
                PacktNewMenuBar.Visible = true;
            }



            // Verify the PacktCustomToolBar exist and add to the application
            if (PacktCustomToolBar == null)
            {
                // Adding the commandbar to Active explorer
                Office.CommandBars PacktBars = this.Application.ActiveExplorer().CommandBars;
                // Adding PacktCustomToolBar to the commandbars
                PacktCustomToolBar = PacktBars.Add("NewPacktToolBar", Office.MsoBarPosition.msoBarTop, false, true);
            }
            // Adding button to the custom tool bar
            Office.CommandBarButton MyButton1 = (Office.CommandBarButton)PacktCustomToolBar.Controls.Add(1, missing, missing, missing, missing);
            // Set the button style
            MyButton1.Style = Office.MsoButtonStyle.msoButtonCaption;
            // Set the caption and tag string
            MyButton1.Caption = "Analyze PKCS#7 structure";
            MyButton1.Tag = "Analyze PKCS#7 structure";
            if (this.PacktButtonA == null)
            {

                // Adding the event handler for the button in the toolbar
                this.PacktButtonA = MyButton1;
                PacktButtonA.Click += new Office._CommandBarButtonEvents_ClickEventHandler(ButtonClick);
            }
            olApp = new Outlook.ApplicationClass();
            Outlook._NameSpace olNs = olApp.GetNamespace("MAPI");

            olExplorer = olApp.ActiveExplorer();
           //  olExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
            
            

        }

    private void PreviewEventHandler (Outlook.Attachment Attachment,ref bool Cancel) {
        Cancel = false;
    }
 
        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder =
            olApp.ActiveExplorer().CurrentFolder;

           
            Outlook.Selection oSel = olExplorer.Selection;
           /* if (oSel.Count != 1)
            {
                MessageBox.Show("Please select only one message for PKCS#7 structure inspection!");
                return;
            }
            * */
            System.Collections.IEnumerator i = oSel.GetEnumerator();
            i.MoveNext();

            Object selObject = (Outlook.MailItem)i.Current;
            if (selObject is Outlook.MailItem)
            {
                Outlook.MailItem mailItem =
                   (selObject as Outlook.MailItem);
                try
                {
                    bool doIt = false;

                   

                    SafeMailItem smi = new SafeMailItemClass();
                    smi.Item = mailItem;

                    mailID = mailItem.EntryID;
                    /*Outlook.NameSpace ns = olApp.GetNamespace("MAPI");
                    
                    Outlook.MAPIFolder f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    Outlook.Folders folders = f.Folders;
                    foreach (Outlook.Folder fol in folders) {
                        if (fol.Name == "TEST") {
                            //
                            break;
                        }

                    }

                    */
                    // @"\\Mailbox - Dauberschmidt, Markus\Inbox\TEST"

                    
                    
                    

                    

                    
                    
                }
                catch (Exception ie)
                {
                    MessageBox.Show("Exception: " + ie.Message);
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void ButtonClick(Office.CommandBarButton ButtonContrl, ref bool CancelOption)
        {

            // Message box displayed on button click

            
            P7MViewer.frmMain m = new P7MViewer.frmMain();
            Outlook.MAPIFolder selectedFolder =
            olApp.ActiveExplorer().CurrentFolder;

           
            Outlook.Selection oSel = olExplorer.Selection;
            if (oSel.Count != 1)
            {
                MessageBox.Show("Please select only one message for PKCS#7 structure inspection!");
                return;
            }
           
            System.Collections.IEnumerator i = oSel.GetEnumerator();
            i.MoveNext();

            Object selObject = (Outlook.MailItem)i.Current;
            if (selObject is Outlook.MailItem)
            {
                Outlook.MailItem mailItem =
                   (selObject as Outlook.MailItem);
                try
                {
                    SafeMailItem smi = new SafeMailItemClass();
                    smi.Item = mailItem;
                    


                    m.setOutlookMessage(mailItem);
                    m.Show();
                    m.readMail();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
