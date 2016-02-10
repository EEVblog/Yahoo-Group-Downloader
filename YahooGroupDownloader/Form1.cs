using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;

namespace YahooGroupDownloader
{
   

    public partial class Form1 : Form
    {
        string DefaultURL = "https://groups.yahoo.com/neo/groups/${GROUP}/conversations/messages";
        string PostURL = "/${INDEX}";
        const int BufferSize = 65536;  // 64 Kilobytes
        string CurrentURL = "";
        string GroupName = "No name entered.";
        int Messages = 0;
        int PostIndex = 1;
        int Stage = 0;
        int LogIndex = 0;
        float Progress;
        string prev = "";

        Stopwatch ElapsedTimer = new Stopwatch();

        System.IO.StreamWriter File;

        public Form1()
        {
            InitializeComponent();
            prev = textBox2.Text;

            //Setup timers
            timeout.Interval = Int32.Parse(textBox2.Text) * 1000;
            antifreeze.Interval = Int32.Parse(textBox2.Text) * 10000;
        }

        void AddLog(string Log)
        {
            listBox1.Items.Add(LogIndex.ToString() + " :\t" + Log);

            int lines = listBox1.Height / listBox1.ItemHeight;
            while (listBox1.Items.Count > lines)
                listBox1.Items.RemoveAt(0);

            listBox1.TopIndex = listBox1.Items.Count - 1;
            LogIndex += 1;
        }

        void ToURL(string URL)
        {
            try
            {
                webBrowser1.Stop();
                webBrowser1.Navigate(new Uri(URL));
                AddLog("Navigating to\t" + URL);
            }
            catch (System.UriFormatException)
            {
                AddLog("Cannot navigate to " + URL);
                return;
            }
        }

        void AddPost(string Title, string Author, string Date, string Content)
        {
            string output = "";

            output += "<tr style = \"border: 2px solid #000;\">\n";
            output += "<td colspan=" + '"' + "1" + '"' + " style = \"border: 2px solid #000;\">" + Title + "</td>\n";
            output += "<td colspan=" + '"' + "1" + '"' + " style = \"border: 2px solid #000;\">" + Author + "</td>\n";
            output += "<td colspan=" + '"' + "1" + '"' + " style = \"border: 2px solid #000;\">" + Date + "</td>\n";
            output += "</tr>\n";
            output += "<tr style = \"border: 2px solid #000;\">\n";
            output += "<td colspan=" + '"' + "3" + '"' + " style = \"border: 2px solid #000;\">" + Content + "</td>\n";
            output += "</tr>\n";

            File.Write(output);
        }

        void ToMessage(int Index)
        {
            //Setup the timeout
            timeout.Stop();
            timeout.Enabled = false;
            
            CurrentURL = DefaultURL.Replace("${GROUP}", GroupName) + PostURL.Replace("${INDEX}", Index.ToString());
            ToURL(CurrentURL);

            Progress = (float)((float)(PostIndex) / (float)(Messages) * 100.0);

            float elapsed = ElapsedTimer.ElapsedMilliseconds;
            float avgtime = elapsed / PostIndex;

            int remaining = (int)(avgtime * (Messages - PostIndex));
            if (remaining < 0)
                remaining = 0;

            TimeSpan ts = TimeSpan.FromMilliseconds(remaining);
            string tream = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                        ts.Hours,
                        ts.Minutes,
                        ts.Seconds,
                        ts.Milliseconds);

            ts = TimeSpan.FromMilliseconds(elapsed);
            string elaps = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                        ts.Hours,
                        ts.Minutes,
                        ts.Seconds,
                        ts.Milliseconds);

            Text = "Complete: " + (Progress).ToString() + "%    Remaining Time: " + tream + "    Elapsed Time: " + elaps;
            progressBar1.Value = (int)Progress;

            //Setup the timeout
            timeout.Enabled = true;
            timeout.Start();
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //Only run if the entire page is loaded
            if (e.Url.AbsolutePath != (sender as WebBrowser).Url.AbsolutePath)
                return;

            if (webBrowser1.ReadyState == WebBrowserReadyState.Complete)
            {
                HtmlElement doc = webBrowser1.Document.Body;

                switch (Stage)
                {
                    case 0:
                        //Get the number of posts
                        string Temp = doc.Document.GetElementById("loading-msg").InnerText;
                        string[] stringSeparators = new string[] { "of total ", " messages" };
                        string[] result = Temp.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                        string CountStr = "";
                        PostIndex = 0;

                        if (result.Length > 1)
                        {
                            CountStr = result[1].Replace(",", "");
                            Messages = Int32.Parse(CountStr);
                            AddLog("Group has " + Messages.ToString() + " messages.");
                            Stage = 1;
                            PostIndex = 1;
                            ToMessage(PostIndex);
                        }
                        break;
                    case 1:
                        //Go through all posts
                        
                        if (PostIndex <= Messages)
                        {
                            //Get data from message
                            string[] TitleSep = new string[] { "</SPAN>" };
                            string PostTitle = doc.Document.GetElementById("yg-msg-subject").InnerHtml.Split(TitleSep, StringSplitOptions.RemoveEmptyEntries)[1];
                            string PostAuthor = "";
                            string PostDate = "";
                            string PostContent = "";

                            //Get document paramters
                            HtmlElement Content              = doc.Document.GetElementById("yg-msg-view");
                            HtmlElementCollection Collection = Content.GetElementsByTagName("div");
                            
                            //Get the date and author
                            foreach (HtmlElement Item in Collection)
                            {
                                string Attribute = Item.GetAttribute("classname");
                                if ( Attribute == "author fleft fw-600")
                                {
                                    PostAuthor = Item.InnerText;
                                    AddLog("Author found\t" + PostAuthor);
                                }
                                else if (Attribute == "tm single")
                                {
                                    PostDate = Item.InnerText;
                                    AddLog("Date found\t" + PostDate);
                                }
                                else if (Attribute == "msg-content undoreset")
                                {
                                    PostContent = Item.InnerHtml;
                                    AddLog("Content found\t" + PostContent);
                                }
                            }

                            //Output to a html file
                            AddPost(PostTitle, PostAuthor, PostDate, PostContent);

                            //Goto next message
                            PostIndex += 1;
                            ToMessage(PostIndex);
                        }
                        else
                        {
                            File.WriteLine("</table>");

                            //Setup the timeout
                            timeout.Stop();
                            timeout.Enabled = false;
                            
                            AddLog("Yahoo group download complete!");
                            Text = "YahooGroupDownloader";

                            string fullPath = ((System.IO.FileStream)(File.BaseStream)).Name;
                            AddLog("Posts output to file -> " + fullPath);
                            AddLog("Path has been added to your clipboard.");
                            Clipboard.SetText(fullPath, TextDataFormat.Text);

                            ElapsedTimer.Stop();
                            ElapsedTimer.Reset();

                            File.Close();
                            Stage = 2;
                        }
                        break;
                    case 2:

                        break;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                GroupName = textBox1.Text;
                CurrentURL = DefaultURL.Replace("${GROUP}", GroupName);
                Stage = 0;

                try
                {
                    if (File != null)
                        File.Close();

                    File = new System.IO.StreamWriter((GroupName + ".html"), false, Encoding.UTF8, BufferSize);
                    File.WriteLine("<table style = \"border: 2px solid #000;\">");
                    ToURL(CurrentURL);

                    ElapsedTimer.Start();
                }
                catch (System.UriFormatException)
                {
                    ToURL("File is likely in use, please close it.");
                }
            }
        }

        private void timeout_Tick(object sender, EventArgs e)
        {
            AddLog("The following post at index " + PostIndex.ToString() + " timed out.");

            //Goto next message
            PostIndex += 1;
            ToMessage(PostIndex);
        }


        string lastURL = "";
        private void antifreeze_Tick(object sender, EventArgs e)
        {
            if (webBrowser1.Url != null)
            {
                if (Stage != 2)
                {
                    string url = webBrowser1.Url.AbsolutePath;
                    if (url == lastURL)
                    {
                        AddLog("Antifreeze kicked in at " + PostIndex.ToString() + " timed out.");

                        //Goto next message
                        PostIndex += 1;
                        ToMessage(PostIndex);
                    }
                    lastURL = url;
                }
            }
            else
            {
                AddLog("To begin enter the group name then click run.");
            }
        }


        
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            // Filter out non-digit text input
            string txt = textBox2.Text;
            for (int i = 0; i < txt.Length; i++)
            {
                int c = txt[i];
                if (c < '0' || c > '9')
                    textBox2.Text = textBox2.Text.Remove(i, 1);
            }


            if (textBox2.Text.Length > 0 && (prev != textBox2.Text))
            {
                //Setup timers
                timeout.Interval = Int32.Parse(textBox2.Text) * 1000;
                antifreeze.Interval = Int32.Parse(textBox2.Text) * 10000;

                AddLog("Timeout set to " + (timeout.Interval / 1000).ToString() + " seconds");
                AddLog("Antifreeze set to " + (antifreeze.Interval / 1000).ToString() + " seconds");
            }

            prev = textBox2.Text;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(this, new EventArgs());
            }
        }
    }
}
