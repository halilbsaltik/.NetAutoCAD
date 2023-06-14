using System.Collections.Specialized;

namespace DNSCAD
{
    public class DNSMenu
    {
        /*public static void Attach()
        {
            string myCuiFilex = @"C:\DNS4\dns4.cuix";

            CustomizationSection cs = new CustomizationSection(Application.GetSystemVariable("MENUNAME") + ".cuix");

            if (!System.IO.File.Exists(myCuiFilex))
            {
                cs.RemovePartialMenu(myCuiFilex, "DNS");

                cs = new CustomizationSection();
                cs.MenuGroupName = "DNS";

                MacroGroup mg = new MacroGroup("DNS", cs.MenuGroup);
                new MenuMacro(mg, "Dinamik", "^C^CdnsDinamik", "ID_dnsDinamik");
                new MenuMacro(mg, "Arge", "^C^CdnsArge", "ID_dnsArge");

                CreateDNSMenu(cs, mg);
                //CreateDNSRibbon(cs, mg);

                cs.SaveAs(myCuiFilex);
                LoadCui(myCuiFilex);
            }
            else if (!cs.PartialCuiFiles.Contains(myCuiFilex))
            {
                LoadCui(myCuiFilex);
            }
        }

        private static void CreateDNSMenu(CustomizationSection cs, MacroGroup mg)
        {
            PopMenu pm = new PopMenu("DNS", new StringCollection() { "POP15" }, "ID_dnsPop", cs.MenuGroup);
            new PopMenuItem(mg.MenuMacros[0], "Dinamik", pm, -1);
            new PopMenuItem(mg.MenuMacros[1], "Arge", pm, -1);
        }

        private static void CreateDNSRibbon(CustomizationSection cs, MacroGroup mg)
        {
            RibbonTabSource rts = new RibbonTabSource(cs.MenuGroup.RibbonRoot);
            rts.Text = "DNS";
            rts.Id = "ID_dnsRibbonTab";

            RibbonPanelSource rps = new RibbonPanelSource(cs.MenuGroup.RibbonRoot);
            rps.Name = "DNS Ribbon Panel";
            rps.ElementID = rps.Id = "ID_dnsRibbonPanel";
            
            RibbonPanelSourceReference rpsr = new RibbonPanelSourceReference(rts);
            rpsr.PanelId = rps.ElementID;
            rts.Items.Add(rpsr);

            RibbonRow rr = new RibbonRow(rpsr);
            RibbonCommandButton rcb1 = new RibbonCommandButton(rr);
            rcb1.Id = "ID_dnsRibbonCommandButton1";
            rcb1.MacroID = mg.MenuMacros[0].ElementID;

            RibbonCommandButton rcb2 = new RibbonCommandButton(rr);
            rcb2.Id = "ID_dnsRibbonCommandButton2";
            rcb2.MacroID = mg.MenuMacros[1].ElementID;

            rr.Items.Add(rcb1);
            rr.Items.Add(rcb2);
            rps.Items.Add(rr);

            cs.MenuGroup.RibbonRoot.RibbonPanelSources.Add(rps);
            cs.MenuGroup.RibbonRoot.RibbonTabSources.Add(rts);
        }

        private static void LoadCui(string cuiFile)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;

            object oldCmdEcho = Application.GetSystemVariable("CMDECHO");
            object oldFileDia = Application.GetSystemVariable("FILEDIA");
            Application.SetSystemVariable("CMDECHO", 0);
            Application.SetSystemVariable("FILEDIA", 0);

            doc.SendStringToExecute("_.cuiload " + cuiFile + " ", false, false, false);

            doc.SendStringToExecute("(setvar \"FILEDIA\" " + oldFileDia.ToString() + ")(princ) ", false, false, false);

            doc.SendStringToExecute(
              "(setvar \"CMDECHO\" "
              + oldCmdEcho.ToString()
              + ")(princ) ",
              false, false, false
            );
        }*/
    }
}
