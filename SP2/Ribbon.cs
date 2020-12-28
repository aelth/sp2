using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using SP2.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SP2
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        
        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SP2.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public Bitmap GetButtonImage(IRibbonControl control)
        {
            return Resources.phishing;
        }

        public void ReportPhishing_Click(IRibbonControl control)
        {
            // first check whether any message has been selected
            var activeOutlookExp = Globals.ThisAddIn.Application.ActiveExplorer();
            int selectedMailsCount = 0;
            try
            {
                selectedMailsCount = activeOutlookExp.Selection.Count;
            }
            catch (System.Exception)
            {
                // no messages have been selected
                selectedMailsCount = 0;
            }

            if(selectedMailsCount > 0)
            {
                string noticeMsgPart = "selected suspicious email" + ((selectedMailsCount == 1) ? "" : "s");
                string noticeMsg = String.Format("Are you sure that you want to forward {0} to {1} for further analysis?", noticeMsgPart, Properties.Settings.Default.report_email);
                var msgForwardDialogResult = MessageBox.Show(noticeMsg, "SP2 forwarding", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (msgForwardDialogResult == DialogResult.Yes)
                {
                    try
                    {
                        ForwardPhishingMails(activeOutlookExp.Selection);
                    }
                    catch (System.Exception)
                    {
                        // TODO: handle
                    }
                }
            }
            else
            {
                MessageBox.Show("No emails selected!", "SP2", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }   
        }

        private void ForwardPhishingMails(Microsoft.Office.Interop.Outlook.Selection selection)
        {
            // prepare report email first
            Outlook.MailItem reportEmail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);

            reportEmail.To = Properties.Settings.Default.report_email;
            reportEmail.BodyFormat = (Properties.Settings.Default.body_plaintext) ? OlBodyFormat.olFormatPlain : OlBodyFormat.olFormatHTML;
            reportEmail.Importance = Outlook.OlImportance.olImportanceLow;
            reportEmail.DeleteAfterSubmit = false;
            reportEmail.OriginatorDeliveryReportRequested = false;
            reportEmail.ReadReceiptRequested = false;

            

            int reportedEmailsCount = 0;
            int errorsEncountered = 0;
            StringBuilder body = new StringBuilder(Properties.Settings.Default.body_start);
            foreach (object selectedItem in selection)
            {
                try
                {
                    if (selectedItem is Outlook.MailItem)
                    {
                        var phishEmail = (Outlook.MailItem)selectedItem;
                        reportEmail.Subject = String.Format("{0} {1}", Properties.Settings.Default.subject_prefix, phishEmail.Subject);

                        // extract headers - they will either be parsed, or pasted directly into the message
                        // see http://msdn.microsoft.com/en-us/library/ms530451.aspx and https://gist.github.com/ghotz/8b4c542254c5d440c1cd for header extraction
                        string rawHeaders = (string)phishEmail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E");
                        MailHeaderCollection headers = (Properties.Settings.Default.parse_headers) ? new MailHeaderCollection(rawHeaders) : null;

                        body.AppendLine("\n");
                        body.AppendFormat("Email {0}", ++reportedEmailsCount);
                        body.AppendLine();
                        body.AppendLine("---------------");

                        body.AppendFormat("Subject: {0}", phishEmail.Subject);
                        body.AppendLine();
                        body.AppendFormat("Sender: {0}", phishEmail.SenderEmailAddress);
                        body.AppendLine();
                        string recipients = string.Empty;
                        foreach (Outlook.Recipient recipient in phishEmail.Recipients)
                        {
                            recipients += String.Format("{0},", recipient.Address);
                        }
                        
                        body.AppendFormat((phishEmail.Recipients.Count >  1) ? "Recipients: {0}" : "Recipient: {0}", recipients.TrimEnd(new char[] { ',' }));
                        body.AppendLine();

                        if (phishEmail.CC != string.Empty)
                        {
                            body.AppendFormat("CC: {0}", phishEmail.CC);
                            body.AppendLine();
                        }

                        if(headers != null)
                        {
                            body.AppendFormat("Date sent: {0}", headers.GetHeaderString("date"));
                            body.AppendLine();
                        }

                        body.AppendLine();
                        if (Properties.Settings.Default.extract_links)
                        {
                            try
                            {
                                List<string> links = ExtractLinks(phishEmail.HTMLBody);
                                body.AppendFormat("Number of links: {0}", links.Count);
                                body.AppendLine();
                                foreach (var link in links)
                                {
                                    // defang - modify link to make it inactive
                                    string defangedLink = link.Replace("https://", "hxxps[://]").Replace("http://", "hxxp[://]");
                                    body.AppendFormat("\t- {0}", defangedLink);
                                    body.AppendLine();
                                }
                            }
                            catch (System.Exception)
                            {
                                // TODO: handle exception
                            }
                        }

                        body.AppendFormat("Number of attachments: {0}", phishEmail.Attachments.Count);
                        body.AppendLine();
                        foreach (Outlook.Attachment attach in phishEmail.Attachments)
                        {
                            body.AppendFormat("\t- {0}", attach.FileName);
                            body.AppendLine();
                        }

                        body.AppendLine("Mail headers:");
                        // append parsed headers, if parsing is enabled
                        if (headers != null)
                        {
                            try
                            {
                                int i = 1;
                                foreach (var receivedLine in GetOrderedReceivingLines(headers))
                                {
                                    body.AppendFormat("\t{0} {1}", i++, receivedLine);
                                    body.AppendLine();
                                }

                                body.AppendLine();

                                foreach (var interestingHeaderPair in GetInterestingHeaders(headers))
                                {
                                    body.AppendFormat("\t{0}: {1}", interestingHeaderPair.Item1, interestingHeaderPair.Item2);
                                    body.AppendLine();
                                }
                            }
                            catch (System.Exception)
                            {
                                // TODO: handle exception while parsing headers
                            }
                        }

                        body.AppendLine("-------");
                        body.AppendLine(rawHeaders);
                        body.AppendLine();

                        // although selected item can be added as attachment directly (reportEmail.Attachments.Add(selectedItem as Object)), this results in attachment type that is difficult to parse on our (CERT) side
                        // in order to get the attachments that can be (more or less) easily parsed, save the phishing message as MSG in temporary folder and add it as a file attachment
                        string tmpMsgName = Path.GetTempPath() + Guid.NewGuid().ToString() + ".msg";
                        phishEmail.SaveAs(tmpMsgName, Outlook.OlSaveAsType.olMSG);
                        //reportEmail.Attachments.Add(tmpMsgName, Outlook.OlAttachmentType.olEmbeddeditem);
                        reportEmail.Attachments.Add(tmpMsgName, Outlook.OlAttachmentType.olByValue);

                        // cleanup
                        try
                        {
                            // remove read-only flag if it was set and delete the file
                            File.SetAttributes(tmpMsgName, FileAttributes.Normal);
                            File.Delete(tmpMsgName);
                        }
                        catch (System.Exception ex)
                        {
                            // TODO: handle exception
                        }
                    }
                    else
                    {
                        // only emails are supported for forwarding
                        MessageBox.Show("Selected message does not appear to be an email - currently only forwarding of suspicious emails is supported. Support for other Outlook item types (invites, etc.) not yet implemented", "Invalid Outlook item type", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                catch (System.Exception)
                {
                    errorsEncountered++;
                }
            }

            if (reportedEmailsCount > 1)
            {
                reportEmail.Subject = String.Format("{0} Multiple suspicious emails", Properties.Settings.Default.subject_prefix);
            }

            reportEmail.Body = body.ToString();
            reportEmail.Send();

            if (errorsEncountered > 0)
            {
                MessageBox.Show("Some emails were not succesfully forwarded. Please retry.", "Forwarding error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private List<string> ExtractLinks(string mailBody)
        {
            List<string> links = new List<string>();
            Regex regex = new Regex("(?:href|src)=[\"|']?(.*?)[\"|'|>]+", RegexOptions.Singleline | RegexOptions.CultureInvariant);
            if (regex.IsMatch(mailBody))
            {
                foreach (Match match in regex.Matches(mailBody))
                {
                    links.Add(match.Groups[1].Value);
                }
            }

            return links;
        }

        private List<string> GetOrderedReceivingLines(MailHeaderCollection headers)
        {
            List<string> orderedReceivedList = new List<string>(10);
            Regex type1Rex = new Regex(@"from\s+(.*?)\s+by(.*?)(?:(?:with|via)(.*?)(?:\sid\s|$)|\sid\s|$)", RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            Regex type2Rex = new Regex(@"()by(.*?)(?:(?:with|via)(.*?)(?:\sid\s|$)|\sid\s)", RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

            List<string> recHeaders = headers.GetHeader("Received");
            for (int i = recHeaders.Count - 1; i >= 0; i--)
            {
                string curReceived = recHeaders[i];
                Match match = null;
                if(curReceived.StartsWith("from", StringComparison.InvariantCultureIgnoreCase))
                {
                    match = type1Rex.Match(curReceived);
                }
                else
                {
                    match = type2Rex.Match(curReceived);
                }

                if(match != null && match.Success)
                {
                    // format received line and remove multiple whitespaces and new lines
                    string parsedReceived = String.Format("{0} --> {1} - {2}", match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value).Replace("\r\n", " ");
                    parsedReceived = Regex.Replace(parsedReceived, @"\s+", " ");
                    orderedReceivedList.Add(parsedReceived);
                }
            }

            return orderedReceivedList;
        }

        private List<Tuple<string, string>> GetInterestingHeaders(MailHeaderCollection headers)
        {
            string[] headerKeys = new string[] {
                "Message-ID",
                "Return-Path",
                "Reply-To",
                "Content-Type",
                "X-Originating-IP",
                "X-Received",
                "X-Mailer",
                "X-Spam-Flag",
                "X-Spam-Status",
                "Received-SPF",
                "Authentication-Results",
                "DKIM-Signature",
                "ARC-Authentication-Results",
            };

            var headerList = new List<Tuple<string, string>>(15);

            foreach (var headerKey in headerKeys)
            {
                string headerValue = headers.GetHeaderString(headerKey);
                if(headerValue != string.Empty)
                {
                    headerValue = headerValue.Replace("\r\n", " ");
                    headerList.Add(new Tuple<string, string>(headerKey, headerValue));
                }
            }

            return headerList;
        }

        #endregion
    }
}
