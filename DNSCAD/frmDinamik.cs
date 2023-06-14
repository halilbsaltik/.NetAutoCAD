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
    public partial class frmDinamik : Form
    {
        public frmDinamik()
        {
            InitializeComponent();
            this.MinimumSize = new System.Drawing.Size(1192, 732);
        }
        
        public static string JsonDataString;
        public static string JsonStatusMessage ="";
        public static bool FormVisibility;

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Visible = false;
            this.Opacity = 0.0;
            webBrowser1.IsWebBrowserContextMenuEnabled = false;
                webBrowser1.AllowWebBrowserDrop = false;
                pictureBox1.Visible = false;
                var result = Task.Run(async () => await GetFilesNcreateDataTask());
                //string upTwoDir = Path.GetFullPath(Path.Combine(System.AppContext.BaseDirectory, @"..\web\")) + "loading.html";
                /*string upTwoDir = @"X:\mysexe\DNSCAD\web\loading.html";
                webBrowser1.DocumentCompleted += webBrowser1_DocumentCompleted;
                webBrowser1.Navigate(new Uri(upTwoDir));*/
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
                if (File.Exists(@"X:\mysexe\DNSCAD\DataJson\Dinamik\DinamikForm.json"))
                {
                    using (StreamReader r = new StreamReader(@"X:\mysexe\DNSCAD\DataJson\Dinamik\DinamikForm.json"))
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

                webBrowser1.Navigate(new Uri(@"X:\mysexe\DNSCAD\web\index.html"));

                Invoke(new Action(() =>
                {
                    this.Visible = true;
                }));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                JsonStatusMessage = ex.ToString();
                DnsSearch.frmDinamik.Dispose();
            }

        }





        /*
        public async Task GetFilesNcreateDataTask()
        {
            try
            {
                
                var CheckData = GetListData();
                if (CheckData != "")
                {
                    JsonDataString = CheckData;
                }
                else
                {
                    Invoke(new Action(() =>
                    {
                        pictureBox1.Visible = true;
                    }));

                    List<Data> _data = new List<Data>();
                    string[] Directories = Directory.GetDirectories(@"X:\aytena\DYNAMİC", "*", SearchOption.AllDirectories);
                    foreach (var Directoryy in Directories)
                    {
                        string[] files = Directory.GetFiles(Directoryy, "*.dwg");
                        foreach (var file in files)
                        {
                            var filename = file.Split('\\').Last();
                            var imagePath = file.Replace(".dwg", ".png");
                            var Base64Image = ImageToBase64(imagePath);

                            _data.Add(new Data(
                                filename,
                                file,
                                imagePath,
                                Base64Image
                            ));
                        }
                    }
                    string json = JsonConvert.SerializeObject(_data.ToArray());
                    JsonDataString = json;
                    var TempPath = @"C:\ProgramData\DNSCAD\data.json";
                    File.WriteAllText(TempPath, json);
                }
                Invoke(new Action(() =>
                {
                    webBrowser1.Visible = true;
                    pictureBox1.Visible = false;
                    pictureBox1.Dispose(); //Destroy gif object
                }));
                webBrowser1.Navigate(new Uri(@"X:\mysexe\DNSCAD\web\index.html"));
                //webBrowser1.Navigate(new Uri("https://www.google.com"));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                JsonStatusMessage = ex.ToString();
                DnsSearch.frmDinamik.Dispose();
            }

        }
        */
        //Get Base64Image
        /*
        public string ImageToBase64(string inputImagePath)
        {
            try
            {
                byte[] imageArray = File.ReadAllBytes(inputImagePath);
                string base64ImageRepresentation = "data:image/png;base64," + Convert.ToBase64String(imageArray);
                return base64ImageRepresentation;
            }
            catch
            {
                return "";
            }
        }
        */

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
                return "v"+FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileVersion.ToString();
            }
            

            public void Button_Clicked(string path)
            {
                //DnsSearch.minimizeForm();
                DnsSearch.InsertDynamic(path);
            }
        }

        private void frmDinamik_FormClosing(object sender, FormClosingEventArgs e)
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

        private void frmDinamik_VisibleChanged(object sender, EventArgs e)
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

        private void frmDinamik_SizeChanged(object sender, EventArgs e)
        {
            DnsSearch.formSize = this.Size;
        }
    }
}
