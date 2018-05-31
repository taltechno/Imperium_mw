using Microsoft.Vbe.Interop;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch sw = Stopwatch.StartNew();

            //Read Excel file into array
            Console.WriteLine("{0}: Opening and reading Excel File...", sw.Elapsed);
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\ZoHo\\data.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Console.WriteLine("{0}: Excel File Opened...", sw.Elapsed);

            object[,] records = (object[,])xlRange.Value2;
            int rowCount = records.GetLength(0);
            Console.WriteLine("{0}: Excel File Read Into Memory with {1} records...", sw.Elapsed, rowCount);

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("{0}: Quit Excel...", sw.Elapsed);


            var client = new RestClient("https://creator.zoho.eu/api/xml/write");

            XElement xmlRoot = new XElement("ZohoCreator");
            XElement xmlAppList = new XElement("applicationlist");
            xmlRoot.Add(xmlAppList);
            XElement xmlApp = new XElement("application", new XAttribute("name", "imperium"));
            xmlAppList.Add(xmlApp);
            XElement xmlFrmList = new XElement("formlist");
            xmlApp.Add(xmlFrmList);
            XElement xmlFrm = new XElement("form", new XAttribute("name", "Delivery_Test"));
            xmlFrmList.Add(xmlFrm);

            Console.WriteLine("{0}: Started creating XML", sw.Elapsed);
            int count = 0;
            //for (int i = 2; i < rowCount; i++)

            for (int i = 2; i <= rowCount; i++)
                {
                if (records[i, 17].ToString() == "CDLVF33") // && records[i, 34] != null)
                {
                    XElement rec = new XElement("add",
                                    new XElement("field", new XAttribute("name", "Entrepren_r_selskap"), new XElement("value", fixObject(records[i, 1]))),
                                    new XElement("field", new XAttribute("name", "Fylke_Id"), new XElement("value", fixObject(records[i, 2]))),
                                    new XElement("field", new XAttribute("name", "Arbeids_Ordre_Id"), new XElement("value", fixObject(records[i, 4]))),
                                    new XElement("field", new XAttribute("name", "Delprosjekt"), new XElement("value", fixObject(records[i, 19]))),
                                    new XElement("field", new XAttribute("name", "Bestilt"), new XElement("value", fixDate(records[i, 20]))),
                                    new XElement("field", new XAttribute("name", "Hovedprosjekt"), new XElement("value", fixObject(records[i, 22]))),
                                    new XElement("field", new XAttribute("name", "Avtalt_til_dato"), new XElement("value", fixDate(records[i, 34]))),
                                    new XElement("field", new XAttribute("name", "Avvikskode"), new XElement("value", fixObject(records[i, 35]))),
                                    new XElement("field", new XAttribute("name", "Utf_rt_dato"), new XElement("value", fixDate(records[i, 36]))));
                    
                                    xmlFrm.Add(rec);
                    count++;
                    
                }
                if(count == 300 || i == rowCount)
                {
                    Console.WriteLine("{0}: XML created with {1} elements", sw.Elapsed, count);
                    //Console.WriteLine(xmlRoot);
                    //xmlRoot.Save("c:\\ZoHo\\test.xml");
                    //Console.ReadKey();

                    
                    Console.WriteLine("{0}: Sending XML POST with {1} elements", sw.Elapsed, count);
                    var request = new RestRequest("", Method.POST);
                    request.AddParameter("authtoken", Imperium_mw.secrets.ResourceManager.GetString("authtoken"));
                    request.AddParameter("scope", "creatorapi");
                    request.AddParameter("zc_ownername", "taltech");
                    request.AddParameter("XMLString", xmlRoot);
                    IRestResponse response = client.Execute(request);
                    var content = response.Content; // raw content as string
                    //Console.WriteLine("{0}: Content: {1}", sw.Elapsed, content);
                    
                    count = 0;
                    xmlFrm.RemoveNodes();
                    
                }
            }
            
            
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static String fixDate(Object o)
        {
            if(o == null)
            {
                return "";
            }else
            {
                String newDate = DateTime.FromOADate((double)o).ToString("dd/MM/yy");
                return newDate;
            } 
        }
        static String fixObject(Object o)
        {
            if (o == null)
            {
                return "";
            }
            else
            {
                return o.ToString();
            }
        }


    }
}
