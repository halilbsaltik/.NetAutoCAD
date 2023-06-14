using Autodesk.AutoCAD.DatabaseServices;
using Microsoft.VisualBasic.Devices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace DNSCAD
{
    // DNSCAD Data Science


    /// <summary>
    /// 
    /// FunctionList:
    ///     GetUserInformation : Returns Current Username
    ///     GetTimeStamp : Returns TimeStamp
    ///     GetLocalIpAddress : Returns Ip Address
    ///     GetMacAddress : Returns MacAddress
    ///     GetOS : Returns Operating System Name and Version
    ///     GetAcadVersionInfo : Return the Autocad Information
    ///     WriteDB : Writes the input JARRAY data to SqliteDB
    /// 
    /// </summary>
    /// <returns></returns>
    /// 
    public class DNSCADLOGDATACOMMAND
    {
        public string CommandName { get; set; }
        public string User { get; set; }
        public string TimeStamp { get; set; }
        public string LogType { get; set; }
        public string LogDescription { get; set; }

        public DNSCADLOGDATACOMMAND(string commandname, string user, string timestamp, string logtype, string logdescription)
        {
            CommandName = commandname;
            User = user;
            TimeStamp = timestamp;
            LogType = logtype;
            LogDescription = logdescription;
        }
    }
    public class DNSCADLOGDATAERROR
    {

        public string User { get; set; }
        public string TimeStamp { get; set; }
        public string LogType { get; set; }
        public string LogDescription { get; set; }


        public DNSCADLOGDATAERROR(string user, string timestamp, string logtype, string logdescription)
        {
            User = user;
            TimeStamp = timestamp;
            LogType = logtype;
            LogDescription = logdescription;
        }
    }

    public class DNSCADLOGUSERINFO
    {

        public string User { get; set; }
        public string TimeStamp { get; set; }
        public string LogType { get; set; }
        public string LogDescription { get; set; }
        public string OperatingSystem { get; set; }
        public string AcadVersion { get; set; }
        public string AcadProduct { get; set; }
        public string AcadLocale { get; set; }
        public string RecoveryPath { get; set; }
        public string Active { get; set; }

        public DNSCADLOGUSERINFO(string user, string timestamp, string logtype, string logdescription, string operatingsystem, string acadversion, string acadproduct, string acadlocale, string recoverypath, string active)
        {
            User = user;
            TimeStamp = timestamp;
            LogType = logtype;
            LogDescription = logdescription;
            OperatingSystem = operatingsystem;
            AcadVersion = acadversion;
            AcadProduct = acadproduct;
            AcadLocale = acadlocale;
            RecoveryPath = recoverypath;
            Active = active;
        }
    }

    public class DNSCADDS
    {
        public static string GetUserInformation()
        {
            return System.Security.Principal.WindowsIdentity.GetCurrent().Name; // Get Curent User
        }

        public static string GetTimeStamp()
        {
            return Convert.ToString((int)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds); // Get TimeStamp
            /*using (WebClient wc = new WebClient())
            {
                var rest = (JObject)JsonConvert.DeserializeObject(wc.DownloadString("https://mysdns4.com:9094/getserverdt"));
                return rest["E_SERVERDT"].Value<string>();
            }*/
        }

        public static string GetLocalIPAddress()
        {
            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        return ip.ToString();
                    }
                }
                return "";
            }
            catch
            {
                return "";
            }
        }

        public static string GetMacAddress()
        {
            try
            {
                foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
                {
                    // Only consider Ethernet network interfaces
                    if (nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet &&
                        nic.OperationalStatus == OperationalStatus.Up)
                    {
                        return nic.GetPhysicalAddress().ToString();
                    }
                }
                return null;
            }
            catch
            {
                return "";
            }
        }

        public static string GetOS()
        {
            try
            {
                return new ComputerInfo().OSFullName;
            }
            catch
            {
                return "";
            }
        }

        public static string GetAcadVersionInfo(string infoType)
        {
            var productKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            var groups = Regex.Match(productKey, @"ACAD-([0-9A-F])\d(\d{2}):([0-9A-F]{3})").Groups;
            string release, localeId, productId;
            switch (groups[1].Value)
            {
                case "5": release = "2007"; break;
                case "6": release = "2008"; break;
                case "7": release = "2009"; break;
                case "8": release = "2010"; break;
                case "9": release = "2011"; break;
                case "A": release = "2012"; break;
                case "B": release = "2013"; break;
                case "D": release = "2014"; break;
                case "E": release = "2015"; break;
                case "F": release = "2016"; break;
                case "0": release = "2017"; break;
                case "1": release = "2018"; break;
                case "2": release = "2019"; break;
                case "3": release = "2020"; break;
                case "4": release = "2021"; break;
                default: release = "unknown"; break;
            }
            switch (groups[2].Value)
            {
                case "00": productId = "Autodesk Civil 3d"; break;
                case "01": productId = "AutoCAD"; break;
                case "0A": productId = "AutoCAD OEM"; break;
                case "02": productId = "AutoCAD Map"; break;
                case "04": productId = "AutoCAD Architecture"; break;
                case "05": productId = "AutoCAD Mechanical"; break;
                case "06": productId = "AutoCAD MEP"; break;
                case "07": productId = "AutoCAD Electrical"; break;
                case "16": productId = "AutoCAD P & ID"; break;
                case "17": productId = "AutoCAD Plant 3d"; break;
                case "29": productId = "AutoCAD ecscad"; break;
                case "30": productId = "AutoCAD Structural Detailing"; break;
                default: productId = "unknown"; break;
            }
            switch (groups[3].Value)
            {
                case "409": localeId = "English"; break;
                case "407": localeId = "German"; break;
                case "40C": localeId = "French"; break;
                case "410": localeId = "Italian"; break;
                case "40A": localeId = "Spanish"; break;
                case "415": localeId = "Polish"; break;
                case "40E": localeId = "Hungarian"; break;
                case "405": localeId = "Czech"; break;
                case "416": localeId = "Brasilian Portuguese"; break;
                case "804": localeId = "Simplified Chinese"; break;
                case "404": localeId = "Traditional Chinese"; break;
                case "412": localeId = "Korean"; break;
                case "411": localeId = "Japanese"; break;
                default: localeId = "unknown"; break;
            }

            // return the requested info
            switch (infoType)
            {
                case "release": return release;
                case "productId": return productId;
                case "localeId": return localeId;
                default: return "unknown request type";
            }
        }

        public static JObject WriteDB(string inputParam, JArray inputJ) // SAMPLE WRİTE DB PARAMETERS --> DNSCAD_LOG_INSERT
        {
            try
            {
                var temp = JObject.Parse(inputJ[0].ToString());
                IList<string> keys = (temp).Properties().Select(p => p.Name).ToList();

                JObject JsonResult = new JObject();
                var param = inputParam.Split('_');

                using (var connection = new SQLiteConnection(@"URI=file:X:\mysexe\DNSCAD\LOG\" + param[0] + ".db"))
                {
                    connection.Open();
                    string sql = "create table if not exists " + param[1] + " (";
                    var command = connection.CreateCommand();
                    foreach (var key in keys)
                    {
                        sql += key.ToString() + " TEXT,";
                    }
                    var result = sql.Remove(sql.Length - 1) + ")";
                    command.CommandText = result;
                    command.ExecuteNonQuery();
                }



                if (param[2] == "INSERT")
                {
                    using (var connection = new SQLiteConnection(@"URI=file:X:\mysexe\DNSCAD\LOG\" + param[0] + ".db"))
                    {
                        string stm = "";
                        bool goton = false;

                        connection.Open();
                        if (param[1] == "USERALLINFOLOG")
                        {
                            int rowCount = 0;
                            stm = @"SELECT * FROM " + param[1] + " WHERE User=" + "'" + temp["User"] + "'" + " AND OperatingSystem=" + "'" + temp["OperatingSystem"] + "'" + " AND AcadVersion=" + "'" + temp["AcadVersion"] + "'" + " AND AcadProduct=" + "'" + temp["AcadProduct"] + "'" + " AND AcadLocale=" + "'" + temp["AcadLocale"] + "'" + " AND RecoveryPath=" + "'" + temp["RecoveryPath"] + "'" + " AND Active=" + "'" + temp["Active"] + "'";
                            var cmc = new SQLiteCommand(stm, connection);
                            rowCount = Convert.ToInt32(cmc.ExecuteScalar());
                            if (rowCount > 0)
                            {
                                goton = true;
                            }
                            else
                            {
                                string update = "UPDATE " + param[1] + " SET Active = 0 WHERE User=" + "'" + temp["User"] + "'";
                                var cmdUpdate = new SQLiteCommand(update, connection);
                                cmdUpdate.ExecuteNonQuery();
                            }



                            /*SQLiteDataReader rdr = cmc.ExecuteReader();
                            while (rdr.Read())
                            {
                                goton = true;
                            }*/

                            if (goton == false)
                            {

                                foreach (var item in inputJ)
                                {
                                    var command = connection.CreateCommand();

                                    var x = "INSERT INTO " + param[1] + " (";
                                    foreach (var key in keys)
                                    {
                                        x += key + ",";
                                    }
                                    var MidString = x.Remove(x.Length - 1) + ") VALUES (";

                                    foreach (var key in keys)
                                    {
                                        MidString += "@" + key + ",";
                                        command.Parameters.AddWithValue(key.ToString(), item[key]);
                                    }

                                    var FinalString = MidString.Remove(MidString.Length - 1) + ")";
                                    command.CommandText = FinalString;
                                    command.ExecuteNonQuery();
                                }

                            }


                        }
                        else
                        {
                            foreach (var item in inputJ)
                            {
                                var command = connection.CreateCommand();

                                var x = "INSERT INTO " + param[1] + " (";
                                foreach (var key in keys)
                                {
                                    x += key + ",";
                                }
                                var MidString = x.Remove(x.Length - 1) + ") VALUES (";

                                foreach (var key in keys)
                                {
                                    MidString += "@" + key + ",";
                                    command.Parameters.AddWithValue(key.ToString(), item[key]);
                                }

                                var FinalString = MidString.Remove(MidString.Length - 1) + ")";
                                command.CommandText = FinalString;
                                command.ExecuteNonQuery();
                            }
                        }
                    }

                    JsonResult["SUCCESSFUL"] = "Y"; // Operation Succed.
                    return JsonResult;
                }

                JsonResult["SUCCESSFUL"] = "N"; // Operation Succed.
                return JsonResult;
            }
            catch (Exception ex)
            {
                JObject JsonResult = new JObject();
                JsonResult["SUCCESSFUL"] = "N"; // Operation Succed.
                JsonResult["ERRORINFO"] = ex.ToString(); // Operation Succed.
                return JsonResult;
            }

        }

    }
}
