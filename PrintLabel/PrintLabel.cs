using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using Oracle.DataAccess.Client;
using Print_General;


namespace PrintLabel
{

    [ComVisible(true)]
    [ProgId("PrintLabel.PrintLabel")]
    public class PrintLabel : IWorkflowExtension
    {
        INautilusServiceProvider sp;
        private IDataLayer dal;
        private const string Type = "1";
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                #region params
                string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];


                var workstationId = Parameters["WORKSTATION_ID"];
                #endregion
                ////////////יוצר קונקשן//////////////////////////
                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);
                /////////////////////////////           
                dal = new DataLayer();
                dal.Connect();
                var pg = new PrintOperation(dal, workstationId, Type);
                var sampleID = "";
                var sampleName = "";
                var sampleDscription = "";
                if (tableName == "RESULT")
                {
                    Test test = dal.GetTestById(Convert.ToInt32(rs.Fields["TEST_ID"].Value));
                    sampleID = test.Aliquot.SampleId.ToString();
                    sampleName = test.Aliquot.Sample.Name;
                    sampleDscription = test.Aliquot.Sample.Description;
                }
                else
                {
                    sampleName = rs.Fields["NAME"].Value;
                    string description = rs.Fields["DESCRIPTION"].Value;
                    if (description != null)
                    {
                        sampleDscription = pg.ReverseString(description);
                    }
                    else
                    {
                        sampleDscription = "";
                    }
                    var temp = rs.Fields["SAMPLE_ID"].Value;
                    sampleID = temp.ToString();
                }
                pg.ManipulateHebrew(sampleDscription);

                string ZPLTestLabel =
            "^XA" +

              "^CI28^FT86,68^A@N,62,61,TT0003M_^FD^FS^PQ1" +//שורה שמטפלת בעברית

            "^LH0,0" + //מיקום התחלה
            "^FO20,10" + //מיקום יחסי למיקום התחלה
            "^A@N30,30" +
           string.Format("^FD{0}^FS", sampleName) + //שם
            "^FO20,80" + //מיקום
            "^A@N20,20" +

            string.Format("^FD{0}^FS", sampleDscription) +

        //   "^FO260,0" + "^BQN,4,3" +
           "^FO320,0" + "^BQN,4,3" +  //מיקום של הברקוד + 
                    //string.Format("^FD   {0}^FS", itxt) + //ברקוד
               string.Format("^FDLA,{0}^FS", sampleID) + //ברקוד
           "^XZ";
                pg.Print(ZPLTestLabel);
            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת מדבקה");
                Logger.WriteLogFile(ex);
            }
        }
        //private string removeBadChar(string ip)
        //{
        //    string ret = "";
        //    foreach (var c in ip)
        //    {
        //        int ascii = (int)c;
        //        if ((ascii >= 48 && ascii <= 57) || ascii == 44 || ascii == 46)
        //            ret += c;
        //    }
        //    return ret;
        //}
        //public string GetIp(string printerName)
        //{
        //    string query = string.Format("SELECT * from Win32_Printer WHERE Name LIKE '%{0}'", printerName);
        //    string ret = "";
        //    var searcher = new ManagementObjectSearcher(query);
        //    var coll = searcher.Get();
        //    foreach (ManagementObject printer in coll)
        //    {
        //        foreach (PropertyData property in printer.Properties)
        //        {
        //            if (property.Name == "PortName")
        //            {
        //                ret = property.Value.ToString();
        //            }
        //        }
        //    }
        //    return ret;
        //}
        //private static string ReverseString(string s)
        //{
        //    var str = s;
        //    string[] strsubs = s.Split(Convert.ToChar(" "));
        //    var newstr = "";
        //    string substr = "";
        //    int i;
        //    int c = strsubs.Count();
        //    for (i = 0; i < c; ++i)
        //    {
        //        substr = strsubs[i];
        //        if (HasHebrewChar(strsubs[i]))
        //        {
        //            substr = Reverse(substr);
        //        }

        //        newstr += substr + " ";
        //    }
        //    return newstr;
        //}

        //private static string Reverse(string s)
        //{
        //    char[] arr = s.ToCharArray();
        //    Array.Reverse(arr);
        //    return new string(arr);
        //}

        //public static bool HasHebrewChar(string value)
        //{
        //    return value.ToCharArray().Any(x => (x <= 'ת' && x >= 'א'));
        //}

        //private void Print(string name, string description, string ID, string ip)
        //{
        //    string ipAddress = ip;
        //    // ZPL Command(s)
        //    string ntxt = name;
        //    string dtxt = "";
        //    if (HasHebrewChar(description))
        //    {
        //        var split = ReverseString(description).Split(' ');
        //        split.Reverse();
        //        foreach (string s in split)
        //        {
        //            dtxt = s + " " + dtxt;
        //        }
        //    }
        //    else
        //    {
        //        dtxt = description;
        //    }
        //    //            dtxt = split.ToString();
        //    string itxt = ID;
        //    string ZPLString =
        //      //  "^XA^CI28^FT86,68^A@N,62,61,TT0003M_^FD^FS^PQ1^XZ";
        //    "^XA" +
        //    "^CI28^FT86,68^A@N,62,61,TT0003M_^FD^FS^PQ1" +
        //    "^LH0,0" +
        //    "^FO13,15" +
        //    "^A@N30,30" +
        //   string.Format("^FD{0}^FS", ntxt) +
        //    "^FO10,150" +
        //    "^A@N30,30" +
        //    string.Format("^FD{0}^FS", dtxt) +
        //    "^FO260,25" + "^BQN,4,5" +
        //        //string.Format("^FD   {0}^FS", itxt) +
        //     string.Format("^FDLA,{0}^FS", itxt) + //ברקוד
        //     "^XZ";
        //    try
        //    {
        //        //MessageBox.Show(ntxt + " name1");
        //        //MessageBox.Show(dtxt + " description");
        //        //MessageBox.Show(itxt + " code");
        //        // Open connection
        //        System.Net.Sockets.TcpClient client = new System.Net.Sockets.TcpClient();
        //        client.Connect(ipAddress, _port);
        //        // Write ZPL String to connection
        //        StreamWriter writer = new StreamWriter(client.GetStream());
        //        writer.Write(ZPLString);
        //        writer.Flush();
        //        // Close Connection
        //        writer.Close();
        //        client.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.InnerException.Message);
        //    }
        //}

    }
}
