using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office;
using System.Xml;
using System.Web;
using System.Net;
using System.IO;
using System.Collections;
using System.Windows.Documents;


namespace HealthCheckWinApp
{
    public partial class Form1 : Form
    {
        urlClass HealthData = new urlClass();
        public List<ListOfURLs> tmpList = new List<ListOfURLs>();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                readXML();
                
                webBrowser1.DocumentText = HTMLdoc();
            }
            catch (Exception ex)
            {
                var msg = ex.Message;
            }
        }

        private void check_Click(object sender, EventArgs e)
        {
            sendMail();
        }

        private string HTMLdoc()
        {
            string tableStyle = " style='border:2px solid black;border-collapse: collapse'";
            string trStyle = " style=''";
            string thStyle = " style='border:2px solid black;'";
            string thOKStyle = " style='border:2px solid black;background-color:green;'";
            string thErrorStyle = " style='border:2px solid black;background-color:red;'";
            StringBuilder _html = new StringBuilder();
            _html.Append("<center>");
            _html.Append("<b><h3>" + HealthData.title + "</b></h3>");
            _html.Append("<table" + tableStyle + "><tr style='background-color:#eff0f1;'" + trStyle + "><th " + thStyle + ">URL</th><th" + thStyle + ">Status</th></tr>");
            for (int i = 0; i < tmpList.Count(); i++)
            {
                string StatusRes = checkURLstatus(tmpList.ToList()[i].url).ToString();
                string urlStatus = StatusRes == "OK" ? "<td" + thOKStyle + ">" + StatusRes + "</td>" : "<td" + thErrorStyle + " > " + StatusRes + "</td>";
                _html.Append("<tr" + trStyle + "><td" + thStyle + "><a href='"+ tmpList.ToList()[i].url + "'>" + tmpList.ToList()[i].name + "<a/></td>" + urlStatus + "</tr>");
            }
            _html.Append("</table></center>");
            return _html.ToString();
        }
        //private void readData()
        //{
        //    XmlTextReader reader = new XmlTextReader("books.xml");
        //    var readerName = "";
        //    var title = "";
        //    var cc = "";
        //    var to = "";
        //    var regards = "";
        //    while (reader.Read())
        //    {
        //        switch (reader.NodeType)
        //        {
        //            case XmlNodeType.Element: // The node is an element.
        //                readerName = @"<" + reader.Name + ">";
        //                break;

        //            case XmlNodeType.Text: //Display the text in each element.
        //                if (readerName == "<TITLE>")
        //                    title = title + reader.Value;
        //                else if (readerName == "<CC>")
        //                    cc = cc + reader.Value;
        //                else if (readerName == "<TO>")
        //                    to = to + reader.Value;
        //                else if (readerName == "<URL>")
        //                    regards = regards + reader.Value;
        //                break;
        //        }
        //    }

        //    HealthData.title = title;
        //    HealthData.regards = regards;
        //    HealthData.to = to;
        //    HealthData.cc = cc;

        //    //return HealthData;
        //}
        private void sendMail()
        {
            try
            {
                Outlook.Application app = new Outlook.Application();
                Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "This is the subject";
                mailItem.To = HealthData.to;
                mailItem.CC = HealthData.cc;
                mailItem.Body = "This is the message.";
                //mailItem.Attachments.Add(logPath);//logPath is a string holding path to the log.txt file
                mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
                //mailItem.Display(false);
                mailItem.Send();
            }
            catch (Exception k)
            {
                string m = k.Message;
            }
        }
        private string checkURLstatus(string url)
        {
            try
            {
                string status = "";
                // Creates an HttpWebRequest for the specified URL.
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                // Sends the HttpWebRequest and waits for a response.
                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                if (myHttpWebResponse.StatusCode == HttpStatusCode.OK)
                {
                    Console.WriteLine("\r\nResponse Status Code is OK and StatusDescription is: {0}", myHttpWebResponse.StatusDescription);
                    status = "OK";
                }
                // Releases the resources of the response.
                myHttpWebResponse.Close();
                return status;
            }
            catch (WebException e)
            {
                Console.WriteLine("\r\nWebException Raised. The following error occurred : {0}", e.Status);
                return e.Status.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nThe following Exception was raised : {0}", e.Message);
                return e.Message.ToString();
            }
        }
        private void readXML()
        {
            XmlDataDocument xmldoc = new XmlDataDocument();
            XmlNodeList xmlnode;
            FileStream fs = new FileStream("books.xml", FileMode.Open, FileAccess.Read);
            XmlTextReader reader = new XmlTextReader("books.xml");
            
            var readerName = "";

            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element: // The node is an element.
                        readerName = @"<" + reader.Name + ">";
                        break;

                    case XmlNodeType.Text: //Display the text in each element.
                        if (readerName == "<TITLE>")
                            HealthData.title = reader.Value;
                        else if (readerName == "<CC>")
                            HealthData.cc = reader.Value;
                        else if (readerName == "<TO>")
                            HealthData.to = reader.Value;
                        else if (readerName == "<REGARDS>")
                            HealthData.regards = reader.Value;

                        break;
                }
            }

            xmldoc.Load(fs);
            xmlnode = xmldoc.GetElementsByTagName("DATA");
            for (int i = 0; i <= xmlnode.Count - 1; i++)
            {
                tmpList.Add(new ListOfURLs(xmlnode[i].ChildNodes.Item(0).InnerText.Trim(), xmlnode[i].ChildNodes.Item(1).InnerText.Trim()));
            }
        }
    }


    public class urlClass
    {
        public string title { get; set; }
        public string regards { get; set; }
        public string cc { get; set; }
        public string to { get; set; }
    }

    public class ListOfURLs
    {

        public ListOfURLs(string v1, string v2)
        {
            this.url = v1;
            this.name = v2;
        }

        public string url { get; set; }
        public string name { get; set; }
    }
}
