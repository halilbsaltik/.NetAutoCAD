
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace DNSCAD
{
    public class DnsSearch : IExtensionApplication
    {
        public static string strFilePath = "";
        public static int RibbonMenu = 0;
        public static frmDinamik frmDinamik;
        public static frmArge frmArge;
        public static frmKomut frmKomut;
        public static bool formError = false;
        public static string ActiveFormName;
        public static System.Drawing.Size formSize = new System.Drawing.Size(1200, 750);
        static int lynE = 0;
        //public static Editor ed;
        //public static Document doc;


        public enum DnsCommand
        {
            dinamik,
            arge,
            komut,
            sesliasistan,
            yardim
        }

        const string appText = "DNS";
        const string appDesc = "DNSCAD Menü Paneli ";

        public void Initialize()
        {
            registrySave();
            GetSaveFilePathACAD();
            LoadFas();
            //batchScriptUpdate();

            //DNSMenu.Attach();
            //Application.Idle += LoadRibbonMenuOnIdle;
            //Application.Idle += LoadPartialMenuOnIdle;


            //Application.DocumentManager.MdiActiveDocument.Editor.PointMonitor += new PointMonitorEventHandler(ed_PointMonitor);
            //Application.DocumentManager.MdiActiveDocument.Editor.SelectionAdded += new SelectionAddedEventHandler(editor_Selection);

            ComponentManager.ApplicationMenu.Opening += new EventHandler<EventArgs>(ApplicationMenu_Opening);

            Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;

        }


        private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            LoadFas();
            //ThreadPool.QueueUserWorkItem(RemoveFas);
        }





        public void Terminate()
        {
            ThreadPool.QueueUserWorkItem(RemoveFas);
            ComponentManager.ApplicationMenu.Opening -= new EventHandler<EventArgs>(ApplicationMenu_Opening);
            //Application.Idle -= new System.EventHandler(Application_OnIdle);

        }

        //private void Application_OnIdle(object sender, System.EventArgs e)
        //{
        //    // Remove the event when it is fired
        //    Application.Idle -= new System.EventHandler(Application_OnIdle);
        //    // Add our Quick Access Toolbar item
        //    AddQuickAccessToolbarItem();
        //}
        public static bool MapControl()
        {
            if (Directory.Exists(@"X:\mysexe\DNSCAD"))
            {
                return true;
            }
            return false;
        }

        private void ApplicationMenu_Opening(object sender, EventArgs e)
        {
            if (Application.DocumentManager.MdiActiveDocument == null)
            {
                return;
            }

            // Remove the event when it is fired
            ComponentManager.ApplicationMenu.Opening -= new EventHandler<EventArgs>(ApplicationMenu_Opening);
            // Add our Application Menu
            if (MapControl())
            {
                AddApplicationMenu();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("\nDNSCAD Menu Oluşturulamıyor,Lütfen X Bağlantınızı Kontrol Ediniz Ve AutoCad Programını Yeniden Başlatınız");
            }
            //AddApplicationMenu();
        }

        private void AddApplicationMenu()
        {
            ApplicationMenu menu = ComponentManager.ApplicationMenu;
            if (menu != null && menu.MenuContent != null)
            {
                // Create our Application Menu Item
                ApplicationMenuItem mi = new ApplicationMenuItem();
                mi.Text = appText;
                mi.Description = appDesc;
                mi.Name = "DNS";
                mi.LargeImage = GetIcon(@"X:\mysexe\DNSCAD\img\ai.ico");
                mi.Items.Add(new RibbonMenuItem()
                {
                    Text = "Dinamik",
                    LargeImage = GetIcon(@"X:\mysexe\DNSCAD\img\dynamic.ico"),
                    CommandHandler = new ribbonCommand((_) => ribbonButton_Clicked(DnsCommand.dinamik), (_) => true)
                });
                mi.Items.Add(new RibbonMenuItem()
                {
                    Text = "AR-GE",
                    LargeImage = GetIcon(@"X:\mysexe\DNSCAD\img\rd.ico"),
                    CommandHandler = new ribbonCommand((_) =>
                  ribbonButton_Clicked(DnsCommand.arge), (_) => true)
                });
                mi.Items.Add(new RibbonMenuItem()
                {
                    Text = "Komut",
                    LargeImage = GetIcon(@"X:\mysexe\DNSCAD\img\command.ico"),
                    CommandHandler = new ribbonCommand((_) =>
                  ribbonButton_Clicked(DnsCommand.komut), (_) => true)
                });

                mi.Items.Add(new RibbonMenuItem()
                {
                    Text = "Sesli Asistan",
                    LargeImage = GetIcon(@"X:\mysexe\DNSCAD\img\voicecommand.ico"),
                    CommandHandler = new ribbonCommand((_) =>
                  ribbonButton_Clicked(DnsCommand.sesliasistan), (_) => true)
                });


                mi.Items.Add(new RibbonSeparator());
                mi.Items.Add(new RibbonMenuItem()
                {
                    Text = "Yardım",
                    LargeImage = GetIcon(@"X:\mysexe\DNSCAD\img\help.ico"),
                    CommandHandler = new ribbonCommand((_) =>
                  ribbonButton_Clicked(DnsCommand.yardim), (_) => true)
                });
                menu.MenuContent.Items.Add(mi);

            }
        }


        private ImageSource GetIcon(string fileName)
        {
            // Get access to it via a stream
            Stream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            using (fs)
            {
                // Decode the contents and return them
                IconBitmapDecoder dec = new IconBitmapDecoder(fs, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
                return dec.Frames[0];
            }
        }

        //private void AddQuickAccessToolbarItem()
        //{
        //    Autodesk.Windows.ToolBars.QuickAccessToolBarSource qat =
        //ComponentManager.QuickAccessToolBar;
        //    if (qat != null)
        //    {
        //        // Create our Ribbon Button
        //        RibbonButton rb = new RibbonButton();
        //        rb.Text = appText;
        //        rb.Description = appDesc;
        //        //rb.Image = GetIcon(smallFile);
        //        rb.Image = new BitmapImage(new System.Uri(@"\\192.168.10.1\users\halil\001_AICAD\AutoCad\AddIn\cbb\res\asearch16.png"));
        //        // Attach the handler to fire out command
        //        rb.CommandHandler = new ribbonCommand((_) => ribbonButton_Clicked(), (_) => true);
        //        // Add it to the Quick Access Toolbar
        //        qat.AddStandardItem(rb);
        //    }
        //}

        [CommandMethod("nrgstr")]
        public void UnregisterDNS()
        {
            try
            {
                // Get the AutoCAD Applications key
                string sProdKey = HostApplicationServices.Current.UserRegistryProductRootKey;
                string sAppName = "DnsCad";
                RegistryKey regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey);
                RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);
                // Delete the key for the application
                regAcadAppKey.DeleteSubKeyTree(sAppName);
                regAcadAppKey.Close();
            }
            catch
            {

            }

        }

        [CommandMethod("dns")]
        public void OpenApplicationMenu()
        {
            ComponentManager.ApplicationMenu.IsOpen = true;
        }

        //private void LoadRibbonMenuOnIdle(object sender, System.EventArgs e)
        //{
        //    createButton();
        //    Application.Idle -= LoadRibbonMenuOnIdle;
        //}

        /*private void LoadPartialMenuOnIdle(object sender, System.EventArgs e)
        {
            try
            {
                //DNSMenu.Attach();
                Application.Idle -= LoadPartialMenuOnIdle;
            }
            catch (Exception ex)
            {

            }
        }*/

        void registrySave()
        {
            try
            {
                //Get the AutoCad Applications Key
                string sProdKey = HostApplicationServices.Current.UserRegistryProductRootKey;
                string sAppName = "DnsCad";
                RegistryKey reqAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey);
                RegistryKey reqAcadAppKey = reqAcadProdKey.OpenSubKey("Applications", true);
                // Check to see if the "MyApp" key exists
                string[] subKeys = reqAcadProdKey.GetSubKeyNames();
                foreach (string subKey in subKeys)
                {
                    //If the application is already registered,exit
                    if (subKey.Equals(sAppName))
                    {
                        reqAcadAppKey.Close();
                        return;
                    }
                }
                //Get the location of this module
                string sAssemblyPath = Assembly.GetExecutingAssembly().Location;
                // Register the application
                RegistryKey regAppAddInKey = reqAcadAppKey.CreateSubKey(sAppName);
                regAppAddInKey.SetValue("DESCRIPTION", sAppName, RegistryValueKind.String);
                regAppAddInKey.SetValue("LOADCTRLS", 14, RegistryValueKind.DWord);
                regAppAddInKey.SetValue("LOADER", sAssemblyPath, RegistryValueKind.String);
                regAppAddInKey.SetValue("MANAGED", 1, RegistryValueKind.DWord);
                reqAcadAppKey.Close();
            }
            catch
            {
            }

        }

        //void createButton()
        //{
        //    try
        //    {
        //        RibbonControl ribbon = ComponentManager.Ribbon;
        //        ribbon = ComponentManager.Ribbon;
        //        RibbonTab rtab = new RibbonTab();
        //        rtab = new RibbonTab();
        //        rtab.Title = "DNS4.0";
        //        rtab.Id = "DNS40";
        //        //Add the Tab
        //        ribbon.Tabs.Add(rtab);
        //        addContent(rtab);

        //    }
        //    catch
        //    {
        //        System.Windows.Forms.MessageBox.Show("");
        //    }

        //}

        //void addContent(RibbonTab rtab)
        //{
        //    try
        //    {
        //        rtab.Panels.Add(AddOnePanel());
        //    }
        //    catch
        //    {
        //        //System.Windows.Forms.MessageBox.Show("DNS 4.0 Menüsü Oluşturulamadı");
        //    }
        //}

        void ribbonButton_Clicked(DnsCommand cmd)
        {
            if (cmd == DnsCommand.dinamik)
            {
                if (MapControl())
                {
                    dnsDinamik();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Lütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
                }

            }
            else if (cmd == DnsCommand.arge)
            {
                //System.Windows.Forms.MessageBox.Show("Bu Alan Geliştirme Aşamasındadır.");
                if (MapControl())
                {
                    dnsArge();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Lütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
                }

            }
            else if (cmd == DnsCommand.komut)
            {
                if (MapControl())
                {
                    dnsKomut();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Lütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
                }
                //System.Windows.Forms.MessageBox.Show("Geliştirme Aşamasındadır.");
            }
            else if (cmd == DnsCommand.sesliasistan)
            {
                /*System.Windows.Forms.MessageBox.Show("Lütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
                if (MapControl())
                {
                    dnsKomut();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("\nLütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
                }*/
                System.Windows.Forms.MessageBox.Show("Geliştirme Aşamasındadır.");
            }
            else if (cmd == DnsCommand.yardim)
            {
                System.Diagnostics.Process.Start("https://mysdns4.com:9093/ygm/EducationVideo/CAD/dnsCAD.mp4");

                System.Diagnostics.Process.Start("https://mysdns4.com:9093/ygm/EducationVideo/CAD/CommandInformation.pdf");

            }
        }

        //RibbonPanel AddOnePanel()
        //{
        //    RibbonButton rb = new RibbonButton();
        //    RibbonPanelSource rps = new RibbonPanelSource();
        //    rps.Title = "DNS4.0";

        //    RibbonPanel rp = new RibbonPanel();
        //    rp.Source = rps;

        //    rb.Name = "SearchButton";
        //    rb.ShowText = false;
        //    rb.ShowImage = true;
        //    //rb.Text = "DNS4.0 Search";
        //    rb.Size = RibbonItemSize.Large;
        //    rb.ToolTip = "DNS4.0 Search System";
        //    try
        //    {
        //        rb.LargeImage = new BitmapImage(new System.Uri(@"\\192.168.10.1\users\halil\001_AICAD\AutoCad\AddIn\cbb\res\asearch.png"));
        //    }
        //    catch
        //    {
        //        //System.Windows.Forms.MessageBox.Show("Lütfen X Ağ Bağlantınızı Kontrol Ediniz");
        //    }
        //    rb.CommandHandler = new ribbonCommand((_) => ribbonButton_Clicked(), (_) => true);
        //    rps.DialogLauncher = rb;
        //    rps.Items.Add(rb);
        //    RibbonMenu = 1;
        //    return rp;
        //}

        internal static void minimizeForm()
        {
            frmDinamik.WindowState = System.Windows.Forms.FormWindowState.Minimized;
        }

        [CommandMethod("dnsDinamik")]
        public static void dnsDinamik()
        {
            if (MapControl())
            {
                if (frmDinamik == null)
                {
                    frmDinamik = new frmDinamik();
                }

                if (ActiveFormName != null)
                {
                    if (ActiveFormName == "ARGE")
                    {
                        frmDinamik.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        frmDinamik.Location = frmArge.Location;
                    }
                    else if (ActiveFormName == "KOMUT")
                    {
                        frmDinamik.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        frmDinamik.Location = frmKomut.Location;
                    }
                }

                frmDinamik.Size = formSize;
                frmDinamik.Show();
                ActiveFormName = "DİNAMİK";
                frmDinamik.WindowState = System.Windows.Forms.FormWindowState.Normal;
                frmDinamik.BringToFront();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Lütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
            }
        }

        [CommandMethod("dnsArge")]
        public static void dnsArge()
        {
            if (MapControl())
            {
                if (frmArge == null)
                {
                    frmArge = new frmArge();
                }

                if (ActiveFormName != null)
                {
                    if (ActiveFormName == "DİNAMİK")
                    {
                        frmArge.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        frmArge.Location = frmDinamik.Location;
                    }
                    else if (ActiveFormName == "KOMUT")
                    {
                        frmArge.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        frmArge.Location = frmKomut.Location;
                    }
                }

                frmArge.Size = formSize;
                frmArge.Show();
                ActiveFormName = "ARGE";
                frmArge.WindowState = System.Windows.Forms.FormWindowState.Normal;
                frmArge.BringToFront();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Lütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
            }
        }

        [CommandMethod("dnsKomut")]
        public static void dnsKomut()
        {
            if (MapControl())
            {
                if (frmKomut == null)
                {
                    frmKomut = new frmKomut();
                }

                if (ActiveFormName != null)
                {
                    if (ActiveFormName == "DİNAMİK")
                    {
                        frmKomut.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        frmKomut.Location = frmDinamik.Location;
                    }
                    else if (ActiveFormName == "ARGE")
                    {
                        frmKomut.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
                        frmKomut.Location = frmArge.Location;
                    }
                }

                frmKomut.Size = formSize;
                frmKomut.Show();
                ActiveFormName = "KOMUT";
                frmKomut.WindowState = System.Windows.Forms.FormWindowState.Normal;
                frmKomut.BringToFront();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Lütfen baglan.bat Dosyasını Çalıştırınız ve X Bağlantınızı Kontrol Ediniz! Sistemi Yeniden Başlatınız.");
            }
        }

        public static bool layoutControl(string pathx)
        {
            return pathx.Contains("\\LY-");
        }

        public static void InsertLayout(string pathx)
        {
            if (File.Exists(pathx))
            {
                try
                {
                    string layoutName = "Layout2";
                    string filename = pathx;
                    // Get the current document and database
                    Document acDoc = Application.DocumentManager.MdiActiveDocument;
                    Database acCurDb = acDoc.Database;
                    acDoc.LockDocument();
                    // Create a new database object and open the drawing into memory
                    Database acExDb = new Database(false, true);
                    acExDb.ReadDwgFile(filename, FileOpenMode.OpenForReadAndAllShare, true, "");
                    // Create a transaction for the external drawing
                    using (Transaction acTransEx = acExDb.TransactionManager.StartTransaction())
                    {
                        // Get the layouts dictionary
                        DBDictionary layoutsEx = acTransEx.GetObject(acExDb.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                        string ly1 = "";
                        foreach (var item in layoutsEx)
                        {
                            ly1 = item.Key;
                            layoutName = ly1 + "-" + lynE;
                            LayoutManager.Current.CreateLayout(layoutName);
                            lynE++;
                            break;
                        }
                        // Get the layout and block objects from the external drawing
                        Layout layEx = layoutsEx.GetAt(ly1).GetObject(OpenMode.ForRead) as Layout;
                        BlockTableRecord blkBlkRecEx = acTransEx.GetObject(layEx.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                        // Get the objects from the block associated with the layout
                        ObjectIdCollection idCol = new ObjectIdCollection();
                        foreach (ObjectId id in blkBlkRecEx)
                        {
                            idCol.Add(id);
                        }
                        // Create a transaction for the current drawing
                        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                        {
                            // Get the block table and create a new block
                            // then copy the objects between drawings
                            BlockTable blkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite) as BlockTable;
                            using (BlockTableRecord blkBlkRec = new BlockTableRecord())
                            {
                                int layoutCount = layoutsEx.Count - 1;
                                blkBlkRec.Name = "*Paper_Space" + layoutCount.ToString();
                                blkTbl.Add(blkBlkRec);
                                acTrans.AddNewlyCreatedDBObject(blkBlkRec, true);
                                acExDb.WblockCloneObjects(idCol,
                                                          blkBlkRec.ObjectId,
                                                          new IdMapping(),
                                                          DuplicateRecordCloning.Ignore,
                                                          false);
                                // Create a new layout and then copy properties between drawings
                                DBDictionary layouts = acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                                using (Layout lay = new Layout())
                                {
                                    lay.LayoutName = layoutName;
                                    lay.AddToLayoutDictionary(acCurDb, blkBlkRec.ObjectId);
                                    acTrans.AddNewlyCreatedDBObject(lay, true);
                                    lay.CopyFrom(layEx);
                                    DBDictionary plSets = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, OpenMode.ForRead) as DBDictionary;
                                    // Check to see if a named page setup was assigned to the layout,
                                    // if so then copy the page setup settings
                                    if (lay.PlotSettingsName != "")
                                    {
                                        // Check to see if the page setup exists
                                        if (plSets.Contains(lay.PlotSettingsName) == false)
                                        {
                                            acTrans.GetObject(acCurDb.PlotSettingsDictionaryId, OpenMode.ForWrite);
                                            using (PlotSettings plSet = new PlotSettings(lay.ModelType))
                                            {
                                                plSet.PlotSettingsName = lay.PlotSettingsName;
                                                plSet.AddToPlotSettingsDictionary(acCurDb);
                                                acTrans.AddNewlyCreatedDBObject(plSet, true);
                                                DBDictionary plSetsEx =
                                                    acTransEx.GetObject(acExDb.PlotSettingsDictionaryId,
                                                                        OpenMode.ForRead) as DBDictionary;
                                                PlotSettings plSetEx = plSetsEx.GetAt(lay.PlotSettingsName).GetObject(OpenMode.ForRead) as PlotSettings;

                                                plSet.CopyFrom(plSetEx);
                                            }
                                        }
                                    }
                                }
                                LayoutManager.Current.CurrentLayout = layoutName;
                            }
                            // Regen the drawing to get the layout tab to display
                            acDoc.Editor.Regen();
                            // Save the changes made
                            acTrans.Commit();
                        }
                        // Discard the changes made to the external drawing file
                        acTransEx.Abort();
                    }
                    // Close the external drawing file
                    acExDb.Dispose();
                    frmDinamik.Enabled = true;
                    CommandLog(layoutName, "InsertLayout", "InsertLayout");

                }
                catch (System.Exception ex)
                {
                    ErrorLog("InsertLayout", ex.ToString());
                    System.Windows.Forms.MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Belirtilen Dosyaya Yetkiniz Bulunmamaktadır.");
            }
            // Specify the layout name and drawing file to work with
        }

        public static void InsertDynamic(string pathx)
        {
            if (File.Exists(pathx))
            {
                if (pathx != "" && pathx != null)
                {
                    frmDinamik.Hide();
                    frmDinamik.Enabled = false;
                    if (layoutControl(pathx))
                    {
                        InsertLayout(pathx);
                        frmDinamik.Size = formSize;
                        frmDinamik.Show();
                        ActiveFormName = "DİNAMİK";
                        frmDinamik.Enabled = true;
                        return;
                    }
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    Editor ed = doc.Editor;
                    //Transaction tr = doc.TransactionManager.StartTransaction();

                    using (Transaction tr = doc.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            PromptPointOptions ppo = new PromptPointOptions("\nYerleştirmek İçin Mouse Tıklama veya X,Y Değerlerini Giriniz.");
                            PromptPointResult ppr = ed.GetPoint(ppo);

                            if (ppr.Status.ToString() != "Cancel")
                            {
                                string dwgName = HostApplicationServices.Current.FindFile(pathx, Application.DocumentManager.MdiActiveDocument.Database, FindFileHint.Default);
                                using (doc.LockDocument())
                                {
                                    using (Database db = new Database(false, false))
                                    {
                                        db.ReadDwgFile(dwgName, FileShare.Read, false, null);
                                        ObjectId BlkId;
                                        BlkId = doc.Database.Insert(dwgName, db, false);
                                        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead, true);
                                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                                        BlockReference bref = new BlockReference(new Point3d(ppr.Value.X, ppr.Value.Y, 0), BlkId);
                                        btr.AppendEntity(bref);
                                        tr.AddNewlyCreatedDBObject(bref, true);
                                        bref.ExplodeToOwnerSpace();
                                        bref.Erase();
                                        tr.Commit();
                                    }
                                }
                            }
                            CommandLog(pathx, "InsertDynamic", "InsertDynamic");
                        }
                        catch (System.Exception ex)
                        {
                            ErrorLog("InsertDynamic", ex.ToString());
                            System.Windows.Forms.MessageBox.Show(ex.ToString());
                            ed.WriteMessage(ex.ToString());
                        }
                        finally
                        {
                            frmDinamik.Show();
                            frmDinamik.Enabled = true;
                        }
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Belirtilen Dosyaya Yetkiniz Bulunmamaktadır.");
            }

        }

        public static void InsertRD(string pathx)
        {
            bool contu = false;
            pathx = pathx.Split('\\').Last();
            string pathd = @"C:\ProgramData\DNSCAD\AICAD\" + pathx;
            pathx = @"X:\mysexe\DNSCAD\CryptoFile\" + pathx;
            byte[] rest = null;
            try
            {
                rest = File.ReadAllBytes(pathx);
                contu = true;
            }
            catch (System.Exception)
            {
                System.Windows.Forms.MessageBox.Show("Temp DWG Oluşturulamadı");
            }
            
            pathx = pathd;
            string blkName;
            if (pathx != "" && pathx != null && contu == true)
            {
                frmArge.Hide();
                frmArge.Enabled = false;
                if (layoutControl(pathx))
                {
                    InsertLayout(pathx);
                    frmArge.Size = formSize;
                    frmArge.Show();
                    ActiveFormName = "ARGE";
                    frmArge.Enabled = true;
                    return;
                }
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;
                //Transaction tr = doc.TransactionManager.StartTransaction();
                using (Transaction tr = doc.TransactionManager.StartTransaction())
                {
                    try
                    {
                        PromptPointOptions ppo = new PromptPointOptions("\nYerleştirmek İçin Mouse Tıklama veya X,Y Değerlerini Giriniz.");
                        PromptPointResult ppr = ed.GetPoint(ppo);
                        if (ppr.Status.ToString() != "Cancel")
                        {

                            DecAlgorithm(rest, pathd);
                            if (File.Exists(pathd))
                            {
                                string dwgName = HostApplicationServices.Current.FindFile(pathx, Application.DocumentManager.MdiActiveDocument.Database, FindFileHint.Default);
                                using (doc.LockDocument())
                                {
                                    using (Database db = new Database(false, false))
                                    {

                                        blkName = Path.GetFileNameWithoutExtension(pathx);
                                        db.ReadDwgFile(dwgName, FileShare.Read, false, null);

                                        ObjectId BlkId;
                                        //BlkId = doc.Database.Insert(dwgName, db, false);
                                        BlkId = doc.Database.Insert(blkName, db, false);

                                        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead, false, true);
                                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                                        //BlockReference bref = new BlockReference(new Point3d(ppr.Value.X, ppr.Value.Y, 0), BlkId);

                                        BlockReference bref = new BlockReference(new Point3d(ppr.Value.X, ppr.Value.Y, 0), bt[blkName]);

                                        btr.AppendEntity(bref);
                                        tr.AddNewlyCreatedDBObject(bref, true);
                                        //bref.ExplodeToOwnerSpace();
                                        //bref.Erase();
                                        tr.Commit();
                                    }
                                }
                                CommandLog(pathx, "InsertRD", "InsertRD");
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("Belirtilen Dosyaya Yetkiniz Bulunmamaktadır.");
                            }





                        }
                    }
                    catch (System.Exception ex)
                    {
                        ErrorLog("InsertRD", ex.ToString());
                        System.Windows.Forms.MessageBox.Show(ex.ToString());
                        ed.WriteMessage(ex.ToString());
                    }
                    finally
                    {
                        frmArge.Show();
                        frmArge.Enabled = true;
                        File.Delete(pathx);
                    }

                }

            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Temp DWG Oluşturulamadı");
            }
        }

        /*async public static void RunCommand(string code)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            doc.SendStringToExecute(code, true, false, false);
        }*/

        public static void RunCommand(string code)
        {
            try
            {
                string Cmd = code + "\n";
                Document doc = Application.DocumentManager.MdiActiveDocument;
                doc.SendStringToExecute(Cmd, true, false, false);
                CommandLog(code, "RunCommand", "Runcommand");
            }
            catch (System.Exception ex)
            {
                ErrorLog("RunCommand", ex.ToString());
            }
        }

        public static void WriteLog(System.Exception input)
        {
            using (var file = new StreamWriter(@"C:\ProgramData\DNSCAD\log.txt", true))
            {
                file.WriteLine(input.ToString());
                file.Close();
            }
        }

        public static void CommandLog(string code, string logType, string LogDesc)
        {
            List<DNSCADLOGDATACOMMAND> _command = new List<DNSCADLOGDATACOMMAND>() { new DNSCADLOGDATACOMMAND(code, DNSCADDS.GetUserInformation(), DNSCADDS.GetTimeStamp(), logType, LogDesc) };
            var jsonCommand = JArray.Parse(JsonConvert.SerializeObject(_command));
            DNSCADDS.WriteDB("DNSCAD_USERCOMMANDLOG_INSERT", jsonCommand);
        }

        public static void ErrorLog(string area, string emes)
        {
            List<DNSCADLOGDATAERROR> _command = new List<DNSCADLOGDATAERROR>() { new DNSCADLOGDATAERROR(DNSCADDS.GetUserInformation(), DNSCADDS.GetTimeStamp(), area, emes) };
            var jsonCommand = JArray.Parse(JsonConvert.SerializeObject(_command));
            DNSCADDS.WriteDB("DNSCAD_USERERRORLOG_INSERT", jsonCommand);
        }

        public static void GetSaveFilePathACAD()
        {
            var productKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            var groups = Regex.Match(productKey, @"ACAD-([0-9A-F])\d(\d{2}):([0-9A-F]{3})").Groups;
            string Val = "";
            string input = productKey + @"\Profiles\<<Unnamed Profile>>\Editor Configuration";
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(input))
            {
                if (key != null)
                {
                    Object o = key.GetValue("SaveFilePath");
                    if (o != null)
                    {
                        Val = o.ToString();
                    }
                }
            }
            List<DNSCADLOGUSERINFO> _command = new List<DNSCADLOGUSERINFO>() { new DNSCADLOGUSERINFO(DNSCADDS.GetUserInformation(), DNSCADDS.GetTimeStamp(), "USERINFOLOG", "USERINFOLOG", DNSCADDS.GetOS(), DNSCADDS.GetAcadVersionInfo("release"), DNSCADDS.GetAcadVersionInfo("productId"), DNSCADDS.GetAcadVersionInfo("localeId"), Val, "1") };
            var jsonCommand = JArray.Parse(JsonConvert.SerializeObject(_command));
            DNSCADDS.WriteDB("DNSCAD_USERALLINFOLOG_INSERT", jsonCommand);

            //File.WriteAllText(@"C:\ProgramData\DNSCAD\AICAD\\SaveFilePath.txt", Val);
        }

        public static void RemoveFas(Object stateInfor)
        {
            Thread.Sleep(20000);
            string lispPath = "C:/ProgramData/DNSCAD/AICAD/System.fas";
            File.Delete(lispPath);
        }

        public static bool DecAlgorithm(byte[] revDatam, string targetpath)
        {
            if (revDatam != null)
            {
                try
                {
                    //string target = @"C:\ProgramData\DNSCAD\AICAD\System.fas";
                    string target = targetpath;

                    byte[] revData = new byte[revDatam.Length / 2];
                    byte[] temprevData = new byte[revDatam.Length / 2];
                    int j = 0;
                    for (int i = 0; i < revData.Length; i++)
                    {
                        revData[i] = (byte)(revDatam[j] - revDatam[j + 1]);
                        j += 2;
                    }
                    int count = 0;
                    for (int i = revData.Length - 1; i >= 0; i--)
                    {
                        temprevData[count] = revData[i];
                        count++;
                    }


                    //File.WriteAllBytes(target, temprevData);
                    using (FileStream fs = new FileStream(target, FileMode.Create))
                    {
                        for (int i = 0; i < temprevData.Length; i++)
                        {
                            fs.WriteByte(temprevData[i]);
                        }
                        fs.Seek(0, SeekOrigin.Begin);
                        for (int i = 0; i < fs.Length; i++)
                        {
                            if (temprevData[i] != fs.ReadByte())
                            {
                                fs.Close();
                                fs.Dispose();
                                return false;
                            }
                        }
                        fs.Close();
                        fs.Dispose();
                    }
                    return true;
                }
                catch
                {
                    return false;
                }

            }
            else
            {
                return false;
            }

        }

        public static void LoadFas()
        {

            string source = @"X:\mysexe\DNSCAD\AICAD\SystemEnc.enc";
            string master = @"C:\ProgramData\DNSCAD\AICAD\System.fas";
            string lispPath = "C:/ProgramData/DNSCAD/AICAD/System.fas";
            byte[] sourceData = File.ReadAllBytes(source);
            bool restC = DecAlgorithm(sourceData, master);

            if (restC)
            {
                if (File.Exists(lispPath))
                {
                    try
                    {
                        Document doc = Application.DocumentManager.MdiActiveDocument;
                        if (doc != null)
                        {
                            string loadStr = String.Format("(load \"{0}\") "/*space after closing paren!!!*/, lispPath);
                            doc.SendStringToExecute(loadStr, true, false, false);


                        }
                    }
                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("Fas Dosyası Yüklenemedi" + " " + ex.ToString());
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Fas Dosyası Yüklenemedi");
                }

            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Fas Dosyası Yüklenemedi");
            }


        }

        public static void batchScriptUpdate()
        {
            string users = Environment.UserName;
            string fileName = "baglan.bat";
            string sourcePath = @"X:\mysexe\DNSCAD";
            string copyt = "C:\\Users\\" + users + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup";
            string copyt1 = "C:\\Users\\" + users + "\\Desktop";


            // Use Path class to manipulate file and directory paths.
            string sourceFile = Path.Combine(sourcePath, fileName);
            string destFile = Path.Combine(copyt, fileName);
            string destFile1 = Path.Combine(copyt1, fileName);
            if (Directory.Exists(@"X:\mysexe\DNSCAD"))
            {
                File.Copy(sourceFile, destFile, true);
                File.Copy(sourceFile, destFile1, true);
            }
        }

        public static void hideForms()
        {
            if (frmDinamik != null)
            {
                frmDinamik.Hide();
            }
            if (frmArge != null)
            {
                frmArge.Hide();
            }
            if (frmKomut != null)
            {
                frmKomut.Hide();
            }
        }


        ////END CLASSS
    }

    public class ribbonCommand : System.Windows.Input.ICommand
    {
        readonly Action<object> execute;
        readonly Predicate<object> canExecute;

        public ribbonCommand(Action<object> execute, Predicate<object> canExecute)
        {
            this.execute = execute;
            this.canExecute = canExecute;
        }

        public void Execute(object parameter)
        {
            execute(parameter);
        }

        public bool CanExecute(object parameter)
        {
            return canExecute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { System.Windows.Input.CommandManager.RequerySuggested += value; }
            remove { System.Windows.Input.CommandManager.RequerySuggested -= value; }
        }
    }
}
