using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Net;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Security.Policy;
using static System.Net.WebRequestMethods;

namespace Cybe_Security_Spam_Detector
{
    public partial class ThisAddIn
    { 
        const string ThreatSubjectAdd = "(Threat Detected!):";
        const string error_message = "Can't Activate Spam Detector ADD_IN\n";
        const string search_reference = "https://www.urlvoid.com/scan/";
        bool error_message_display = false;
        bool there_is_a_threat = false;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            NameSpace name_space = this.Application.GetNamespace("MAPI");
            MAPIFolder folder = name_space.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            Items current_item = folder.Items;
            //add event to new mail 
            current_item.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(spam_detector_addin);
        }

        private void spam_detector_addin(object item)
        {
            Microsoft.Office.Interop.Outlook.MailItem current_mail_item = (Outlook.MailItem)item;
            if(current_mail_item != null)
            {
                //in order to avoid duplication
                HashSet<string> URLs = extract_URLs(current_mail_item);
                //string::Current URL ; int:: The number of sources that identified him as a threat
                Dictionary<string, int> result_dictionary = new Dictionary<string, int>();

                try
                {
                    there_is_a_threat = false;
                    foreach (string url in URLs)
                        using (WebClient client = new WebClient())
                        {
                            //Web Security Protocol
                            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

                            //Get HTML RESULT
                            string httpResult = client.DownloadString(search_reference + url + "/");

                            string html_result_line = Regex.Match(httpResult, @"Detections Counts</span></td><td><span class=(.+?)<").ToString();
                            string rating_string = html_result_line.Substring("Detections Counts</span></td><td><span class=label label-dangerss".Length);

                            int rating = int.Parse(new Regex(@">(.*?)/").Match(rating_string).Groups[1].Value);

                            if (rating > 0)
                            {
                                there_is_a_threat = true;
                            }
                            result_dictionary[url] = rating;
                        }
                    changeMail(createContext(result_dictionary, current_mail_item.SenderEmailAddress, current_mail_item.ReceivedTime.ToString("R"), there_is_a_threat), there_is_a_threat, current_mail_item);
                }
                catch (WebException ex) //NETWORK ERROR
                {
                    if (error_message_display == false)
                    {
                        error_message_display = true;
                        MessageBox.Show(error_message + "An error occurred while connecting to the network\nplease make sure the device is connected to the network\n" + ex.GetType());
                    }
                }
                catch (System.Exception ex) //ANY ERROR
                {
                    if (error_message_display == false)
                    {
                        error_message_display = true;
                        MessageBox.Show(error_message + ex.GetType().Name.ToString());
                    }
                }
            }

        }
        private HashSet<string> extract_URLs(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            HashSet<string> urls = new HashSet<string>();

            string body = mailItem.Body;

            var linkParser = new Regex(@"\b(?:https?://|www\.)\S+\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            foreach (Match link in linkParser.Matches(body))
            {
                string currentUrl = Regex.Replace(link.ToString(), @"https?://|www.", string.Empty);
                urls.Add(currentUrl);
            }
            return urls;
        }

        private string createContext(Dictionary<string, int> scanResult, string sender, string receivingDate, bool threatDetected)
        {
            int numberOfUrls = scanResult.Count;
            string context = "<h1><b><u>Spam Scan Result:</u></b></h1><br/>" +
                "<b>Sender Address: </b> " + sender + "<br/>" +
                "<b>Receiving Date: </b>" + receivingDate + "<br/>";
            if (numberOfUrls == 0)
            {
                context += "<h4 style=color:green>No URLs references were found in the email content.</h4><br/> " +
                    "<b>If you recognize an unusual request in this email for certain information.<br/>" +
                    "Please make sure that the sender of the email is known to you as a verified source ,and his request is legitimate.</b>";
            }
            else
            {
                context += "<h4>" + numberOfUrls + " URLs were found in the email content!</h4><br/>";
                foreach (string key in scanResult.Keys)
                {
                    context += "<b>URL: " + key + "</b><br/>";
                    if (scanResult[key] > 0)
                    {
                        context += "<span style=color:red;><b>" + scanResult[key] + " scan sources flagged him as a threat!</b></span><br/><br/>";
                    }
                    else
                    {
                        context += "The scan sources found no evidence that there was a threat in the URL redirection<br/><br/>";
                    }
                }
                if (threatDetected)
                {
                    context += "<span style=color:red;><b>If you trust the sender of the email please ask them to send new links instead of the links that the system detected as a threat</b></span>";
                }
            }
            return context;
        }

        private void changeMail(string context, bool threatDetected, MailItem item)
        {
            //In order to convert string type to HTML text
            string body = "<span style=font-family: Arial, Helvetica, sans-serif;>" + item.Body.Replace("\r\n", "<br/>") + "</span>";

            if (threatDetected)
            {
                item.Subject = ThreatSubjectAdd + item.Subject;
                item.HTMLBody = "<html><body><h1 style=color:red; font-family: Arial, Helvetica, sans-serif;>Threat Detected</h1><br/>" + body + "<br/> " + context + "</body></html>";
            }
            else
            {
                item.HTMLBody = body + "<br/> " + context + "</body></html>";
            }
            item.Save();
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
