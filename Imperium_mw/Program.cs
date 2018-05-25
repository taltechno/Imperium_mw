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
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            //Read Excel file into array
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\ZoHo\data.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            object[,] records = (object[,])xlRange.Value2;

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

            var client = new RestClient("https://creator.zoho.com/api/taltech/json/imperium/form/Delivery_Test/record/add");
            Stopwatch sw = Stopwatch.StartNew();
           

            //for (int i = 2; i <= rowCount; i++)
            try
            {
                Parallel.For(1, rowCount, (int i, ParallelLoopState state) =>
                    {
                        Console.WriteLine("{0} [{1}] Started iteration", sw.Elapsed, i);
                        Console.WriteLine("{0} [{1}] State: {2}", sw.Elapsed, i, Thread.CurrentThread.ThreadState);
                        
                        /*
                        var request = new RestRequest("", Method.POST);
                        //Authtoken
                      
                        request.AddParameter("authtoken", Imperium_mw.secrets.ResourceManager.GetString("authtoken"));
                        //Fields
                        request.AddParameter("Entrepren_r_selskap", records[i, 1]);
                        request.AddParameter("Fylke_Id", records[i, 2]);
                        request.AddParameter("Fylke_navn", records[i, 3]);
                        request.AddParameter("Arbeids_Ordre_Id", records[i, 4]);
                        request.AddParameter("Kunde_Ordrenr", records[i, 5]);
                        request.AddParameter("Kundenr", records[i, 6]);
                        request.AddParameter("Fornavn", records[i, 7]);
                        request.AddParameter("Etternavn", records[i, 8]);
                        request.AddParameter("Adresse", records[i, 9]);
                        request.AddParameter("Husnr", records[i, 10]);
                        request.AddParameter("Postnr", records[i, 11]);
                        request.AddParameter("Poststed", records[i, 12]);
                        request.AddParameter("Boligtype", records[i, 13]);
                        request.AddParameter("Borettslag", records[i, 14]);
                        request.AddParameter("Kommune", records[i, 15]);
                        request.AddParameter("Ordretype", records[i, 16]);
                        request.AddParameter("Produktkode_Entr.", records[i, 17]);
                        request.AddParameter("Produkt_Navn", records[i, 18]);
                        request.AddParameter("Delprosjekt", records[i, 19]);
                        */
                        if (records[i,20] == null)
                        {
                            //request.AddParameter("Bestilt", null);
                            Console.WriteLine("{0} [{1}] is null", sw, i);
                            Console.WriteLine("{0} [{1}] State: {2}", sw.Elapsed, i, Thread.CurrentThread.ThreadState);
                        }
                        else
                        {
                            Object o = 4321d;
                            Console.WriteLine("o= {0} : Type={1}", o, o.GetType());
                            Double d = (Double)o;
                           
                            Console.WriteLine("d= {0}",d);
                            //Console.WriteLine(DateTime.FromOADate(d)));
                            ////Console.WriteLine("{0} [{1}] State: {2}", sw.Elapsed, i, Thread.CurrentThread.ThreadState);
                            /*
                            Console.WriteLine("{0} [{1}] d={2}",sw.Elapsed, i, d);
                            Object o = DateTime.FromOADate(d.Value);
                            //request.AddParameter("Bestilt", o);
                            Console.WriteLine("{0} [{1}] Bestilt after FromOADate: {2}",sw.Elapsed, i, o);
                            //
                            request.AddParameter("Bestilt", DateTime.FromOADate((double)records[i, 20]));
                            */
                        }
                        //request.AddParameter("Bestilt", records[i, 20] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 20]));
                        /*
                        request.AddParameter("Tidligste_inst._Dato", records[i, 21] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 21]));
                        request.AddParameter("Hovedprosjekt", records[i, 22]);
                        request.AddParameter("AD", records[i, 23] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 23]));
                        request.AddParameter("Mont_rnavn", records[i, 24]);
                        request.AddParameter("Forfall_booking", records[i, 25] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 25]));
                        request.AddParameter("F_rste_kontakt_fors_k", records[i, 26] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 26]));
                        request.AddParameter("Siste_kontakt_fors_k", records[i, 27] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 27]));
                        request.AddParameter("Antall_kontakt_fors_k", records[i, 28]);
                        request.AddParameter("OIS_Status", records[i, 29]);
                        request.AddParameter("Satt_OIS_Status", records[i, 30] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 30]));
                        request.AddParameter("Endret_avvikskode", records[i, 31]);
                        request.AddParameter("Endret_dato", records[i, 32] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 32]));
                        request.AddParameter("Forfalls_dato", records[i, 33] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 33]));
                        request.AddParameter("Avtalt_til-dato", records[i, 34] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 34]));
                        request.AddParameter("Avvikskode", records[i, 35]);
                        request.AddParameter("Utf_rt_dato", records[i, 36] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 36]));
                        request.AddParameter("Satt_til_UTF", records[i, 37] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 37]));
                        request.AddParameter("Satt_til_Annullert", records[i, 38] == null ? (DateTime?)null : DateTime.FromOADate((double)records[i, 38]));
                        request.AddParameter("Sporingsnummer", records[i, 39]);
                        */
                        if (state.IsExceptional)
                        {
                            
                            Console.WriteLine("{0} [{1}] Exception inline: code:{2}", sw.Elapsed, i, Thread.CurrentThread.ThreadState);
                        }

                        // execute the request
                        //IRestResponse response = client.Execute(request);
                        //var content = response.Content; // raw content as string

                        Console.WriteLine("{0} [{1}] Ended iteration",sw.Elapsed, i);
                        Console.WriteLine("{0} [{1}] State: {2}", sw.Elapsed, i, Thread.CurrentThread.ThreadState);
                    }
                 );
            }
            catch (AggregateException ae)
            {
                ae.Handle((inner) =>
                {
                    Console.WriteLine(inner.Message);
                    return true;
                });
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();

            

        }
    }
}
