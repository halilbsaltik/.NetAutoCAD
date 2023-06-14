using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DNSCAD
{
    public partial class frmKomut : Form
    {
        public frmKomut()
        {
            InitializeComponent();
            this.MinimumSize = new System.Drawing.Size(1192, 732);
        }

        public static string JsonDataString;
        public static string JsonStatusMessage = "";
        public static bool FormVisibility;

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Visible = false;
            this.Opacity = 0.0;
            webBrowser1.IsWebBrowserContextMenuEnabled = false;
            webBrowser1.AllowWebBrowserDrop = false;
            pictureBox1.Visible = false;
            var result = Task.Run(async () => await GetFilesNcreateDataTask());
            webBrowser1.ScriptErrorsSuppressed = true;
            webBrowser1.ObjectForScripting = new ScriptInterface();
            webBrowser1.Visible = false;
        }

        void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //JSON İMAGE BASE64 DOSYA YOLU ADI
        }

        public string GetListData()
        {
            try
            {
                string json = "";
                if (File.Exists(@"X:\mysexe\DNSCAD\DataJson\Komut\KomutForm.json"))
                {
                    using (StreamReader r = new StreamReader(@"X:\mysexe\DNSCAD\DataJson\Komut\KomutForm.json"))
                    {
                        json = r.ReadToEnd();

                    }
                    return json;
                }
                else
                {
                    return "";
                }
            }
            catch
            {
                return "";
            }
        }



        public async Task GetFilesNcreateDataTask()
        {
            try
            {
                Invoke(new Action(() =>
                {
                    pictureBox1.Visible = true;
                }));

                var CheckData = GetListData();
                if (CheckData != "")
                {
                    JsonDataString = CheckData;
                }
                Invoke(new Action(() =>
                {
                    webBrowser1.Visible = true;
                    pictureBox1.Visible = false;
                    pictureBox1.Dispose(); //Destroy gif object
                }));

                webBrowser1.Navigate(new Uri(@"X:\mysexe\DNSCAD\web\index3.html"));

                Invoke(new Action(() =>
                {
                    this.Visible = true;
                }));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                JsonStatusMessage = ex.ToString();
                DnsSearch.frmArge.Dispose();
            }

        }





        [ComVisibleAttribute(true)]
        public class ScriptInterface
        {
            public void ChangeForm(string FormName)
            {
                if (FormName == "DİNAMİK")
                {
                    DnsSearch.hideForms();
                    DnsSearch.dnsDinamik();

                }
                else if (FormName == "ARGE")
                {
                    DnsSearch.hideForms();
                    DnsSearch.dnsArge();
                }
                else if (FormName == "KOMUT")
                {
                    DnsSearch.hideForms();
                    DnsSearch.dnsKomut();
                }

            }

            public string callMe()
            {
                string json = JsonDataString;

                return json;
            }

            public string callMe1()
            {
                return JsonStatusMessage;
            }

            public string GetVersion()
            {
                return "v" + FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileVersion.ToString();
            }


            public void Button_Clicked(string code)
            {
                code = (code.Split('\\').Last()).Split(new string[] { ".dwg" }, StringSplitOptions.None)[0];
                DnsSearch.RunCommand(code);
            }
        }

        private void frmKomut_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.Opacity < 1)
            {
                this.Opacity += 0.08;
            }
            else
            {
                timer1.Stop();
            }
        }

        private void Fade()
        {
            timer1.Start();
        }


        private void frmKomut_VisibleChanged(object sender, EventArgs e)
        {
            if (FormVisibility == false)
            {
                this.Opacity = 0;
                Fade();
                FormVisibility = false;
            }
            else
            {
                FormVisibility = true;
            }
        }

        private void frmKomut_SizeChanged(object sender, EventArgs e)
        {
            DnsSearch.formSize = this.Size;
        }
    }
}
