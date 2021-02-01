using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.IO;
using iwantedue;
using iwantedue.Windows.Forms;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;

namespace P7MViewer
{
    public partial class frmMain : Form
    {
        private bool smimefound = false;
        private Dictionary<string, X509Certificate2> certs = new Dictionary<string, X509Certificate2>();
        public frmMain()
        {
            InitializeComponent();

            statusBar.Text = "Please use the \"Open...\" menu item or drag'n'drop an Outlook message";
            
        }

        private string printCertDetails(X509Certificate2 c)
        {
            return string.Format("SubjectName: {0}, valid from: {1}, valid until: {2}", c.Subject, c.NotBefore, c.NotAfter);
        }
        public void parsePKCS7(byte[] pkcs7stream, TreeNode node)
        {            
            byte[] data = pkcs7stream;
            try
            {
                EnvelopedCms envData = new EnvelopedCms();
                envData.Decode(data);

                TreeNode pkcs = node;
                String str;
                str = "Encryption Algorithm";
                txtBox.AppendText(str +"\n");
                TreeNode encAlgo = pkcs.Nodes.Add(str);
                str = "    Name=" + envData.ContentEncryptionAlgorithm.Oid.FriendlyName + " / OID=" + envData.ContentEncryptionAlgorithm.Oid.Value;
                txtBox.AppendText(str + "\n");
                TreeNode encName = pkcs.Nodes.Add(str.Trim());

                str = "    Key length="+ envData.ContentEncryptionAlgorithm.KeyLength+" bit";
                txtBox.AppendText(str+"\n\n");
                TreeNode encKey = pkcs.Nodes.Add(str.Trim());

                str = "Number of recipients " + envData.RecipientInfos.Count;
                txtBox.AppendText(str + "\n");
                TreeNode numRecp = node.Nodes.Add(str);
                
                int num = 1;
                foreach (RecipientInfo r in envData.RecipientInfos)
                {
                    txtBox.AppendText("=================\n");
                    TreeNode rec = numRecp.Nodes.Add("Recipient Nr: " + num++);
                    str = "    Encrypted key=" + BitConverter.ToString(r.EncryptedKey) + " (" + r.EncryptedKey.Length + " bytes)";
                    txtBox.AppendText(str+"\n");
                    TreeNode n = rec.Nodes.Add(str.Trim());

                    str = "    Encryption alg=" + r.KeyEncryptionAlgorithm.Oid.FriendlyName + ", OID=" + r.KeyEncryptionAlgorithm.Oid.Value;                                        
                    txtBox.AppendText(str + "\n");
                    n = rec.Nodes.Add(str.Trim());

                    str = "    RecipientIdentifier.Type=" + r.RecipientIdentifier.Type;
                    txtBox.AppendText(str + "\n");
                    n = rec.Nodes.Add(str.Trim());
                    
                    switch (r.RecipientIdentifier.Type)
                    {
                        case SubjectIdentifierType.IssuerAndSerialNumber:

                            X509IssuerSerial xi =
                                (X509IssuerSerial)r.RecipientIdentifier.Value;
                            str = "    Issuer=" + xi.IssuerName;
                            txtBox.AppendText(str + "\n");
                            rec.Nodes.Add(str.Trim());

                            str = "    SerialNumber=" + xi.SerialNumber + " (hex)";
                            txtBox.AppendText(str + "\n");
                            rec.Nodes.Add(str.Trim());

                            bool found = false;
                            foreach (X509Certificate2 c in certs.Values)
                            {
                                if (c.SerialNumber == xi.SerialNumber)
                                {
                                    str = "    Certificate Details=" + printCertDetails(c);
                                    txtBox.AppendText(str + "\n");
                                    rec.Nodes.Add(str.Trim());
                                    found = true;
                                }
                            }

                            if (!found)
                            {
                                str = "    Certificate Details=<could not be identified in AD>";
                                txtBox.AppendText(str + "\n");
                                rec.Nodes.Add(str.Trim());
                            }
                            break;
                        case SubjectIdentifierType.NoSignature:
                            str = "    No signature";
                            txtBox.AppendText(str + "\n");
                            rec.Nodes.Add(str.Trim());
                            break;
                        case SubjectIdentifierType.SubjectKeyIdentifier:
                        case SubjectIdentifierType.Unknown:                        
                            str = "    SubjectKeyInfo=" + r.RecipientIdentifier.Value + "(no X.509 tag)";
                            txtBox.AppendText(str + "\n");
                            rec.Nodes.Add(str.Trim());
                            break;
                    }
                    
                }

                txtBox.AppendText("\n");

                str = "Attached certificates " + envData.Certificates.Count;
                txtBox.AppendText(str+"\n");
                TreeNode a = node.Nodes.Add(str);
                foreach (X509Certificate2 cert in envData.Certificates)
                {
                    str = "    Subject=" + cert.SubjectName;
                    txtBox.AppendText(str +"\n");
                    TreeNode n = a.Nodes.Add(str);
                }

                txtBox.AppendText("\n");
                str = "Unprotected Attributes " + envData.UnprotectedAttributes.Count;
                txtBox.AppendText(str +"\n");
                TreeNode u = node.Nodes.Add(str);
                foreach (CryptographicAttributeObject obj in
                    envData.UnprotectedAttributes)
                {
                    str = obj.Oid.FriendlyName;
                    txtBox.AppendText(str + " ");
                    TreeNode unprot = u.Nodes.Add(str);
                }
                txtBox.AppendText("\n");
                                
            }
            catch (Exception ex)
            {
                txtBox.AppendText(ex.Message);
            }
            finally
            {
                // reader.Close();
            }
            
        }
        private void openFile()
        {
            DialogResult msgFileSelectResult = this.ofDlg.ShowDialog();
            if (msgFileSelectResult == DialogResult.OK)
            {
                foreach (string msgfile in this.ofDlg.FileNames)
                {
                    readFile(msgfile);
                }
            }
        }

        private void readFile(String filename) {
            txtBox.Clear();
            treeView.Nodes.Clear();
            Stream messageStream = File.Open(filename, FileMode.Open, FileAccess.Read);
            OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
            messageStream.Close();
            if (!isOutlookMessage(message))
            {
                MessageBox.Show("No Outlook MSG file!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                message.Dispose();
                return;
            }            
            smimefound = false;

            this.loadMessage(message,treeView.Nodes.Add("Message"));
            message.Dispose();
            mnuExport.Enabled = true;
            treeView.ExpandAll();
            txtBox.SelectionStart = 0;
            txtBox.SelectionLength = 0;
            copyContentToClipboardToolStripMenuItem.Enabled = true;

        }

        private bool isOutlookMessage(OutlookStorage.Message message)
        {
            return (!String.IsNullOrEmpty(message.BodyRTF) || !String.IsNullOrEmpty(message.BodyText) || !String.IsNullOrEmpty(message.From) || !String.IsNullOrEmpty(message.Subject) || message.Attachments.Count > 0);
        }

       

        private void btnOpen_Click(object sender, EventArgs e)
        {
            openFile();
        }

        /// <summary>
        /// Get user certificate from Active Directory (Source adapted from https://mjc.si/2016/12/10/get-user-certificate-from-active-directory/")
        /// </summary>
        /// <param name="email">EMail address to get the certificate for</param>
        /// <returns></returns>
        private static X509Certificate2 GetUserCertificateFromAD(string email)
        {
            try
            {
                DirectoryEntry de;

                de = Domain.GetCurrentDomain().GetDirectoryEntry();
                                
                DirectorySearcher dsearch = new DirectorySearcher(de);
                dsearch.Filter = String.Format("(&(objectCategory=person)(objectClass=user)(userPrincipalName={0}))", email);
                SearchResultCollection searchResults = dsearch.FindAll();

                foreach (System.DirectoryServices.SearchResult result in searchResults)
                {
                    //Find userCertificate
                    if (result.Properties.Contains("userCertificate"))
                    {
                        Byte[] certBytes = (Byte[])result.Properties["userCertificate"][0];

                        X509Certificate2 certificate = null;
                        certificate = new X509Certificate2(certBytes);

                        return certificate;
                    }
                    else
                    {
                        //implement logging
                        return null;
                    }
                }

                de.Close();
                de.Dispose();
                return null;
            }
            catch (Exception ex)
            {
                //implement logging
                return null;
            }
        }
		
        private void loadMessage(OutlookStorage.Message message, TreeNode messageNode)
        {                  
            String str;

            str = "Outlook Envelope data";
            messageNode.Text = str;
            txtBox.AppendText(str + "\n");

            str = "SMTP Routing Header\n " + message.GetMapiProperty("007D").ToString(); 
            messageNode.Text = str;
            txtBox.AppendText(str + "\n");

            str = message.Subject;
            messageNode.Nodes.Add(str);
            txtBox.AppendText(str + "\n");

            str = "Subject: " + message.Subject;
            messageNode.Nodes.Add(str);
            txtBox.AppendText(str + "\n");
            
            
            /*TreeNode bodyNode = messageNode.Nodes.Add("Body: (double click to view)");
            bodyNode.Tag = new string[] { message.BodyText, message.BodyRTF };
            */
            str = "Recipients: " + message.Recipients.Count;
            TreeNode recipientNode = messageNode.Nodes.Add(str);
            txtBox.AppendText(str + "\n");
            foreach (OutlookStorage.Recipient recipient in message.Recipients)
            {
                // Load SMIME certificates from AD
                if (!certs.ContainsKey(recipient.Email))
                { 
                    X509Certificate2 cert = GetUserCertificateFromAD(recipient.Email);
                    if (cert != null)
                    {
                        certs.Add(recipient.Email, cert);
                    }
                }
                str = recipient.Type + ": " + recipient.Email;
                recipientNode.Nodes.Add(str);
                txtBox.AppendText(str + "\n");
            }
        
            str = "Attachments: " + message.Attachments.Count;
            TreeNode attachmentNode = messageNode.Nodes.Add(str);
            txtBox.AppendText(str+"\n");
            foreach (OutlookStorage.Attachment attachment in message.Attachments)
            {
                str = attachment.Filename + ": " + attachment.Data.Length + " bytes";
                attachmentNode.Nodes.Add(str);
                txtBox.AppendText(str + "\n");
                // Check for SMIME attachment
                if (attachment.Filename.Contains("p7m")) // Weaken the check
                {
                    smimefound = true;
                    txtBox.AppendText("==== PKCS#7 Enveloped data ====\n");
                    
                    parsePKCS7(attachment.Data, attachmentNode);
                    
                }
            }

            str = "Sub Messages: " + message.Messages.Count;
            TreeNode subMessageNode = messageNode.Nodes.Add(str);
            txtBox.AppendText(str + "\n");
            foreach (OutlookStorage.Message subMessage in message.Messages)
            {               
                this.loadMessage(subMessage, subMessageNode.Nodes.Add("MSG"));
            }
            if (smimefound)
            {
                statusBar.Text = "S/MIME attachment(s) found!";
            }
            else
            {
                statusBar.Text = "No S/MIME attachment(s) found!";
            }
        }

        private void about()
        {
            MessageBox.Show("(c) 2010 by Markus Dauberschmidt");
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFile();
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            about();
        }

        private void viewAsTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            txtBox.Visible = true;
            treeView.Visible = false;
            viewAsTextToolStripMenuItem.Checked = true;
            viewAsTreeToolStripMenuItem.Checked = false;
        }

        private void viewAsTreeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            txtBox.Visible = false;
            treeView.Visible = true;
            viewAsTextToolStripMenuItem.Checked = false;
            viewAsTreeToolStripMenuItem.Checked = true;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            saveDecoded();
        }
        private void saveDecoded()
        {
            DialogResult msgFileSelectResult = this.sfDlg.ShowDialog();
            if (msgFileSelectResult == DialogResult.OK)
            {

                TextWriter tw = new StreamWriter(sfDlg.FileName, false);
                tw.WriteLine(txtBox.Text);
                tw.Close();
            }
        }

        private void treeView_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)                
                e.Effect = DragDropEffects.All;
        }

        private void treeView_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            // loop through the string array, adding each filename to the ListBox
            foreach (string file in files)
            {
                readFile(file);
            }
        }

        private void frmMain_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            // loop through the string array, adding each filename to the ListBox
            foreach (string file in files)
            {
                readFile(file);
            }
        }

        private void frmMain_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
                e.Effect = DragDropEffects.All;
        }

        private void copyContentToClipboardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(txtBox.Text);
        }


    }
}

       
    

