using Microsoft.Office.Interop.Excel;
using System;
using System.Threading;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Timers;
using System.Drawing.Text;

namespace Simu_Gen
{
    public partial class Simu_Gen : Form
    {
        private Stopwatch stopwatch;
        public Simu_Gen()
        {
            InitializeComponent();
            stopwatch = new Stopwatch();
            timer1.Interval = 1000;
            timer1.Tick += Timer_Tick;
            timer1.Start();
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            timelabel.Text = "Elapsed Time: " + stopwatch.Elapsed.ToString("hh\\:mm\\:ss");
        }
       
            

        private void Form1_Load(object sender, EventArgs e)
        {
            flowLayoutPanel1.Enabled = false;
            Output.Enabled = false;
            Deselect_All.Enabled = false;
            Select_All.Enabled = false;
        }
        private void UpdateRichTextBox(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(UpdateRichTextBox), message);
            }
            else
            {
                richTextBox1.Text += message + Environment.NewLine;
            }
        }
        private void UpdateElapsedTime(string time)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(UpdateElapsedTime), time);
            }
            else
            {
                stopwatch.Start();
                timelabel.Text = "Elapsed Time: " + time;
            }

        }
     



        private void Output_Click(object sender, EventArgs e)
        {
            stopwatch.Start();

            this.Output.Enabled = false;

            if (!File.Exists(File_Entry.Text))
            {
                MessageBox.Show("File Does not exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(99);
            }
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook wbApp = xlApp.Workbooks.Open(File_Entry.Text);



            FileInfo PAR_File = new FileInfo(Output_Folder.Text + @"\PAR.xml");
            if (PAR_File.Exists == true)
            {
                PAR_File.Delete();
            }

            

            Excel._Worksheet ws = wbApp.Sheets["CT2_Track Circuit"];
            Excel.Range xlRange = ws.UsedRange;
            int r_count = xlRange.Count;
            int c_count = xlRange.Count;

            Excel._Worksheet ws_TD_111 = wbApp.Sheets["CT111_Traffic Direction Out TBZ"];
            Excel.Range xlRange_TD_111 = ws_TD_111.UsedRange;
            int r_count_TD_111 = xlRange_TD_111.Count;

            Excel._Worksheet ws_TD_112 = wbApp.Sheets["CT112_Traffic Direction In TBZ"];
            Excel.Range xlRange_TD_112 = ws_TD_112.UsedRange;
            int r_count_TD_112 = xlRange_TD_112.Count;

            Excel._Worksheet ws_CC_3 = wbApp.Sheets["CT3_Cycle Control "];
            Excel.Range xlRange_CC_3 = ws_CC_3.UsedRange;
            int r_count_CC_3 = xlRange_CC_3.Count;

            Excel._Worksheet ws_MB_4 = wbApp.Sheets["CT4_Maintenance Block"];
            Excel.Range xlRange_MB_4 = ws_MB_4.UsedRange;
            int r_count_MB_4 = xlRange_MB_4.Count;

            Excel._Worksheet ws_MA_9 = wbApp.Sheets["CT9_Manual Authorization Point"];
            Excel.Range xlRange_MA_9 = ws_MA_9.UsedRange;
            int r_count_MA_9 = xlRange_MA_9.Count;

            Excel._Worksheet ws_PL_10 = wbApp.Sheets["CT10_Point Locking"];
            Excel.Range xlRange_PL_10 = ws_PL_10.UsedRange;
            int r_count_PL_10 = xlRange_PL_10.Count;

            Excel._Worksheet ws_SR_13 = wbApp.Sheets["CT13_Subroute Released"];
            Excel.Range xlRange_SR_13 = ws_SR_13.UsedRange;
            int r_count_SR_13 = xlRange_SR_13.Count;

            Excel._Worksheet ws1 = wbApp.Sheets["CT12_Subroute Locked"];
            Excel.Range xlRange1 = ws1.UsedRange;
            int r1_count = xlRange.Count;
            int c1_count = xlRange.Count;

            Excel._Worksheet ws_OC_141 = wbApp.Sheets["CT141_Overlap Calling Releasing"];
            Excel.Range xlRange_OC_141 = ws_OC_141.UsedRange;
            int r_count_OC_141 = xlRange_OC_141.Count;

            Excel._Worksheet ws_OE_142 = wbApp.Sheets["CT142_Overlap Establishment"];
            Excel.Range xlRange_OE_142 = ws_OE_142.UsedRange;
            int r_count_OE_142 = xlRange_OE_142.Count;

            Excel._Worksheet ws_MS_151 = wbApp.Sheets["CT151_Main Signal Proceed Asp."];
            Excel.Range xlRange_MS_151 = ws_MS_151.UsedRange;
            int r_MS_15_count = xlRange_MS_151.Count;

            Excel._Worksheet ws_SS_152 = wbApp.Sheets["CT152_Shunt Signal Proceed Asp."];
            Excel.Range xlRange_SS_152 = ws_SS_152.UsedRange;
            int r_SS_15_count = xlRange_SS_152.Count;

            Excel._Worksheet ws_RR_16 = wbApp.Sheets["CT16_Route Released"];
            Excel.Range xlRange_RR_16 = ws_RR_16.UsedRange;
            int r_RR_16_count = xlRange_RR_16.Count;

            Excel._Worksheet ws_RI_5 = wbApp.Sheets["CT5_Control Route by Individual"];
            Excel.Range xlRange_RI_5 = ws_RI_5.UsedRange;
            int r_RI_5_count = xlRange_RI_5.Count;

            Excel._Worksheet ws_RF_6 = wbApp.Sheets["CT6_Control Route by Fleet Mode"];
            Excel.Range xlRange_RF_6 = ws_RF_6.UsedRange;
            int r_RF_6_count = xlRange_RF_6.Count;

            Excel._Worksheet ws_RI_17 = wbApp.Sheets["CT17_Route Indicator"];
            Excel.Range xlRange_RI_17 = ws_RI_17.UsedRange;
            int r_RI_17_count = xlRange_RI_17.Count;

            Excel._Worksheet ws_DI = wbApp.Sheets["CT17_Route Indicator"];
            Excel.Range xlRange_DI = ws_DI.UsedRange;
            int r_count_DI = xlRange_DI.Rows.Count;

            /*Thread workerThread_TC = new Thread(() =>
            {
                Track_circuit(r_count, xlRange, PAR_File);
            });
            workerThread_TC.Start();
            Thread workerThread_TD = new Thread(() =>
            {
                Traffic_direction(r_count_TD_111, xlRange_TD_111, PAR_File, r_count_TD_112, xlRange_TD_112);
            });
            workerThread_TD.Start();
            Thread workerThread_MBL = new Thread(() =>
            {
                Mbl(r_count_MB_4, xlRange_MB_4, PAR_File);
            });
            workerThread_MBL.Start();
            Thread workerThread_SB = new Thread(() =>
            {
                Subroutes(r1_count, xlRange1, r_count_SR_13, xlRange_SR_13, PAR_File);
            });
            workerThread_SB.Start();
            Thread workerThread_Route = new Thread(() =>
            {
                Routes(r_RF_6_count, r1_count, xlRange1, r_MS_15_count, r_SS_15_count, xlRange_MS_151, xlRange_SS_152, r_RR_16_count, xlRange_RR_16, r_RI_5_count, xlRange_RI_5, xlRange_RF_6, PAR_File);
            });
            workerThread_Route.Start();*/

            //workerThread_TC.Join();
            //workerThread_TD.Join();
            //workerThread_MBL.Join();
            //workerThread_SB.Join();
            //workerThread_Route.Join();
            Thread workerThread_TC = new Thread(() =>
            {
                Track_circuit(r_count, xlRange, PAR_File);
                Traffic_direction(r_count_TD_111, xlRange_TD_111, PAR_File, r_count_TD_112, xlRange_TD_112);
                Mbl(r_count_MB_4, xlRange_MB_4, PAR_File);
                Subroutes(r1_count, xlRange1, r_count_SR_13, xlRange_SR_13, PAR_File);
                Routes(r_RF_6_count, r1_count, xlRange1, r_MS_15_count, r_SS_15_count, xlRange_MS_151, xlRange_SS_152, r_RR_16_count, xlRange_RR_16, r_RI_5_count, xlRange_RI_5, xlRange_RF_6, PAR_File);
                Points(r_count_MA_9, xlRange_MA_9, r_count_PL_10, xlRange_PL_10, PAR_File);
                Three_Aspect_Signals(r_RI_17_count, xlRange_RI_17, r_MS_15_count, xlRange_MS_151, PAR_File);
                Cycles(r_count_CC_3, xlRange_CC_3, PAR_File);
                Shunt_Signals(r_SS_15_count, xlRange_SS_152, PAR_File);
                Overlaps(r_count_OC_141, xlRange_OC_141, r_count_OE_142, xlRange_OE_142, PAR_File);
                Mtk(r_count_DI, xlRange_DI);
            });
            workerThread_TC.Start();
            

            MessageBox.Show("Program Completed!", "INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            wbApp.Close(false);
            xlApp.Quit();
            Marshal.ReleaseComObject(wbApp);
            Marshal.ReleaseComObject(xlApp);
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(ws_TD_111);
            Marshal.ReleaseComObject(xlRange_TD_111);
            Marshal.ReleaseComObject(ws_TD_112);
            Marshal.ReleaseComObject(xlRange_TD_112);
            Marshal.ReleaseComObject(ws_CC_3);
            Marshal.ReleaseComObject(xlRange_CC_3);
            Marshal.ReleaseComObject(ws_MB_4);
            Marshal.ReleaseComObject(xlRange_MB_4);
            Marshal.ReleaseComObject(ws_MA_9);
            Marshal.ReleaseComObject(xlRange_MA_9);
            Marshal.ReleaseComObject(ws_PL_10);
            Marshal.ReleaseComObject(xlRange_PL_10);
            Marshal.ReleaseComObject(ws_SR_13);
            Marshal.ReleaseComObject(xlRange_SR_13);
            Marshal.ReleaseComObject(ws1);
            Marshal.ReleaseComObject(xlRange1);
            Marshal.ReleaseComObject(ws_OC_141);
            Marshal.ReleaseComObject(xlRange_OC_141);
            Marshal.ReleaseComObject(ws_OE_142);
            Marshal.ReleaseComObject(xlRange_OE_142);
            Marshal.ReleaseComObject(ws_MS_151);
            Marshal.ReleaseComObject(xlRange_MS_151);
            Marshal.ReleaseComObject(ws_SS_152);
            Marshal.ReleaseComObject(xlRange_SS_152);
            Marshal.ReleaseComObject(ws_RR_16);
            Marshal.ReleaseComObject(xlRange_RR_16);
            Marshal.ReleaseComObject(ws_RI_5);
            Marshal.ReleaseComObject(xlRange_RI_5);
            Marshal.ReleaseComObject(ws_RF_6);
            Marshal.ReleaseComObject(xlRange_RF_6);
            Marshal.ReleaseComObject(ws_RI_17);
            Marshal.ReleaseComObject(xlRange_RI_17);
            Marshal.ReleaseComObject(xlRange_DI);
            Marshal.ReleaseComObject(ws_DI);
            // This example simply waits for 5 seconds
            //Thread.Sleep(5000);
            // Update the UI with the results of the operation
            //Simu_Gen_FormClosing();
            stopwatch.Stop();
            this.Invoke((MethodInvoker)delegate
            {
                this.Output.Enabled = true;
            });

        }
        

            

            


            

            int flag = 0;
            int idx = 11;
            int ind = 11;
            int refe = 11;
            List<string> list_TC = new List<string>();
            List<string> list_OL = new List<string>();
            List<string> list_U = new List<string>();
        private static object _lock = new object();

        //string itineraires = "";

        /* void WriteToFile()
         {
             // Write a line of text to the file
             lock (_lock)
             {
                 using (StreamWriter sw = PAR_File.AppendText())
                 {
                     sw.WriteLine($"Thread {Thread.CurrentThread.ManagedThreadId} wrote this line");
                 }
             }
         }*/
        private void Track_circuit(int r_count, Excel.Range xlRange, FileInfo PAR_File)
        {
            UpdateElapsedTime(stopwatch.Elapsed.ToString("hh\\:mm\\:ss"));
            FileInfo TC_File = new FileInfo(Output_Folder.Text + "//TC_File.xml");
            bool tcEnabled = Track_Circuit.Checked;

            if (tcEnabled)
            {
                //Track Circuit
                for (int i = 11; i <= r_count; i++)
                {
                    if (xlRange.Cells[i, 1].Value2 == null)
                    {
                        break;
                    }

                    lock (_lock)
                    {
                        try
                        {
                            using (StreamWriter sw = TC_File.AppendText())
                            {
                                sw.WriteLine("TC" + xlRange.Cells[i, 1].Value2.ToString() + "\r\nzdv_quai: 0\r\ndist_arret: 50\r\n%\n");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                    }
                }

                UpdateRichTextBox("Track Circuit Completed\n");
            }
        }

        //GC.Collect();
        //GC.WaitForPendingFinalizers();
        //
        //Marshal.ReleaseComObject(xlRange);
        //Marshal.ReleaseComObject(ws);
        private void Traffic_direction(int r_count_TD_111, Excel.Range xlRange_TD_111,FileInfo PAR_File,int r_count_TD_112,Excel.Range xlRange_TD_112) 
        {
            UpdateElapsedTime(stopwatch.Elapsed.ToString("hh\\:mm\\:ss"));
            if (Traffic_Direction.Checked == true)
            {
                //Thread.Sleep(1000);
                //Traffic Direction
                //ws = wbApp.Sheets["CT111_Traffic Direction Out TBZ"];
                //xlRange = ws.UsedRange;
                //r_count = xlRange.Count;
                //c_count = xlRange.Count;
                //Console.WriteLine(r_count);

                for (int i = 11; i <= r_count_TD_111; i++)
                {
                    if (xlRange_TD_111.Cells[i, 1].Value2 == null)
                    {
                        continue;
                    }
                    if (xlRange_TD_111.Cells[i, 1].Value2.ToString().Contains("("))
                    {
                        //Console.WriteLine(i);
                        break;
                    }
                    lock (_lock)
                    {
                        using (StreamWriter sw = PAR_File.AppendText())
                        {
                            try
                            {
                                sw.WriteLine("TD" + xlRange_TD_111.Cells[i, 1].Value2.ToString() + "\r\ntransit_iti: U" + xlRange_TD_111.Cells[i, 1].Value2.ToString() + " .\r\n%\n");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex?.ToString());
                            }
                        }
                    }
                }

                //ws = wbApp.Sheets["CT112_Traffic Direction In TBZ"];
                //xlRange = ws.UsedRange;
                //r_count = xlRange.Count;
                //c_count = xlRange.Count;
                //Console.WriteLine(r_count);

                for (int i = 11; i <= r_count_TD_112; i++)
                {
                    if (xlRange_TD_112.Cells[i, 1].Value2 == null)
                    {
                        continue;
                    }
                    if (xlRange_TD_112.Cells[i, 1].Value2.ToString().Contains("("))
                    {
                        //Console.WriteLine(i);
                        break;
                    }
                    using (StreamWriter sw = PAR_File.AppendText())
                    {
                        try
                        {
                            sw.WriteLine("TD" + xlRange_TD_112.Cells[i, 1].Value2.ToString() + "\r\ntransit_iti: U" + xlRange_TD_112.Cells[i, 1].Value2.ToString() + " .\r\n%\n");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex?.ToString());
                        }
                    }
                }
                //richTextBox1.Text += "Traffic Direction Completed\n";
                UpdateRichTextBox("Traffic Direction Completed\n");
            }
        }
        //GC.Collect();
        //GC.WaitForPendingFinalizers();
        //
        //Marshal.ReleaseComObject(xlRange);
        //Marshal.ReleaseComObject(ws);

        //MBL
        private void Mbl(int r_count_MB_4, Excel.Range xlRange_MB_4, FileInfo PAR_File)
        {
            UpdateElapsedTime(stopwatch.Elapsed.ToString("hh\\:mm\\:ss"));
            bool mblEnabled = MBL.Checked;

            if (mblEnabled)
            {
                Thread.Sleep(500);

                for (int i = 11; i <= r_count_MB_4; i++)
                {
                    if (xlRange_MB_4.Cells[i, 1].Value2 == null)
                    {
                        continue;
                    }
                    if (xlRange_MB_4.Cells[i, 1].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                    lock (_lock)
                    {
                        try
                        {
                            using (StreamWriter sw = PAR_File.AppendText())
                            {
                                sw.WriteLine("MBL" + xlRange_MB_4.Cells[i, 1].Value2.ToString() + "\r\nsetting_release_maintenance_block: MBLSR" + Sector_Entry.Text + " .\r\n%\n");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex?.ToString());
                        }
                    }
                }

                UpdateRichTextBox("MBL Completed\n");
            }
        }
        //GC.Collect();
        //GC.WaitForPendingFinalizers();
        //
        //Marshal.ReleaseComObject(xlRange);
        //Marshal.ReleaseComObject(ws);


        //Subroutes
        private void Subroutes(int r1_count,Excel.Range xlRange1,int r_count_SR_13,Excel.Range xlRange_SR_13,FileInfo PAR_File) 
        {
            UpdateElapsedTime(stopwatch.Elapsed.ToString("hh\\:mm\\:ss"));
            FileInfo SB_File = new FileInfo(Output_Folder.Text+"\\SB.xml");
            if (Subroute.Checked == true)
            {
                //Thread.Sleep(500);
                //Console.WriteLine(r_count);
                string zdv = "zdv: ";
                string transit_amont = "transit_amont: ";
                //List<string> Subroute = new List<string> { "zdv: ", "itis: ", "transit_amont: " };

                Dictionary<string, List<string>> itis = new Dictionary<string, List<string>>();
                //List<string> list_Sub = new List<string>();
                lock (_lock)
                {
                    for (int k = 11; k <= r1_count; k++)
                    {
                        if (xlRange1.Cells[k, 1].Value2 != null && xlRange1.Cells[k, 1].Value2.ToString().Contains("("))
                        {
                            break;
                        }
                        if (xlRange1.Cells[k, 1].Value2 != null)
                        {
                            ind = k;
                        }

                        if (xlRange1.Cells[k, 2].Value2 != null && itis.ContainsKey(xlRange1.Cells[k, 2].Value2.ToString()))
                        {
                            itis[xlRange1.Cells[k, 2].Value2.ToString()].Add("R" + xlRange1.Cells[ind, 1].Value2.ToString());
                        }
                        if (xlRange1.Cells[k, 2].Value2 != null && !itis.ContainsKey(xlRange1.Cells[k, 2].Value2.ToString()))
                        {
                            itis[xlRange1.Cells[k, 2].Value2.ToString()] = new List<string>() { "R" + xlRange1.Cells[ind, 1].Value2.ToString() };
                        }

                    }

                    for (int i = 11; i <= r_count_SR_13; i++)
                    {

                        if (xlRange_SR_13.Cells[i, 9].Value2 != null && xlRange_SR_13.Cells[i, 9].Value2.ToString() != "Nil")
                        {
                            zdv += "TC" + xlRange_SR_13.Cells[i, 9].Value2.ToString() + " ";
                        }
                        if (xlRange_SR_13.Cells[i, 3].Value2 != null && xlRange_SR_13.Cells[i, 3].Value2.ToString() != "Nil")
                        {
                            transit_amont += "U" + xlRange_SR_13.Cells[i, 3].Value2.ToString() + " ";
                        }
                        if (xlRange_SR_13.Cells[i, 1].Value2 != null && xlRange_SR_13.Cells[i, 1].Value2.ToString().Contains("("))
                        {
                            //Console.WriteLine(i);
                            break;
                        }
                        if (xlRange_SR_13.Cells[i + 1, 1].Value2 != null)
                        {

                            //Console.WriteLine(zdv, transit_amont);



                            using (StreamWriter sw = SB_File.AppendText())
                            {
                                sw.WriteLine("U" + xlRange_SR_13.Cells[idx, 1].Value2.ToString());
                                sw.WriteLine(zdv + ".");
                                if (itis.ContainsKey(xlRange_SR_13.Cells[idx, 1].Value2.ToString()))
                                {
                                    sw.WriteLine("itis: " + string.Join(" ", itis[xlRange_SR_13.Cells[idx, 1].Value2.ToString()]) + " .");
                                }
                                else
                                {
                                    sw.WriteLine("itis: .");
                                }
                                sw.WriteLine(transit_amont + ".");
                                sw.WriteLine("%\n");
                            }

                            idx = i + 1;
                            flag = 0;
                            zdv = "zdv: ";
                            //itis = "itis: ";
                            transit_amont = "transit_amont: ";
                        }

                    }
                }
                //richTextBox1.Text += "Subroute Completed\n";
                UpdateRichTextBox("Subroute Completed\n");
            }
        }
        //GC.Collect();
        //GC.WaitForPendingFinalizers();
        //
        //Marshal.ReleaseComObject(xlRange);
        //Marshal.ReleaseComObject(ws);

        //Route
        private void Routes(int r_RF_6_count,int r1_count,Excel.Range xlRange1,int r_MS_15_count,int r_SS_15_count,Excel.Range xlRange_MS_151,Excel.Range xlRange_SS_152,int r_RR_16_count,Excel.Range xlRange_RR_16,int r_RI_5_count,Excel.Range xlRange_RI_5,Excel.Range xlRange_RF_6,FileInfo PAR_File) 
        {
            UpdateElapsedTime(stopwatch.Elapsed.ToString("hh\\:mm\\:ss"));
            FileInfo Route_File = new FileInfo(Output_Folder.Text + @"\Route.xml");
            if (Route.Checked == true)
            {
                //Thread.Sleep(500);
                //ws = wbApp.Sheets["CT5_Control Route by Individual"];
                //xlRange = ws.UsedRange;
                //r_count = xlRange.Count;
                //c_count = xlRange.Count;



                //ws1 = wbApp.Sheets["CT12_Subroute Locked"];
                //xlRange1 = ws1.UsedRange;
                //r1_count = xlRange1.Count;

                //Console.WriteLine(r_count);
                //string transit_iti = "transits_iti: ";
                Dictionary<string, List<string>> transit_iti = new Dictionary<string, List<string>>();
                //string transits_over = "transits_over: ";
                Dictionary<string, List<string>> transit_over = new Dictionary<string, List<string>>();
                string transits_over_opp = "transits_over_opp: ";
                string maintenance_block = "maintenance_block: ";
                string transits_iti_incomp = "transits_iti_incomp: ";
                string itis_incomp = "itis_incomp: ";
                string cycs_incomp = "cycs_incomp: ";
                string aigs_N = "aigs_N: ";
                string aigs_R = "aigs_R: ";
                string zdv_TORR_Ta = "zdv_TORR_Ta: ";
                string zdv_TORR_Tb = "zdv_TORR_Tb: ";
                string zdv_TORR_Tc = "zdv_TORR_Tc: ";
                string trace_perm = "trace_perm: ";
                Dictionary<string, List<string>> zdvs_green_aspect = new Dictionary<string, List<string>> { };
                Dictionary<string, List<string>> zdvs_blue_aspect = new Dictionary<string, List<string>> { };
                Dictionary<string, List<string>> zdvs_approch = new Dictionary<string, List<string>> { };
                //string zdvs_blue_aspect = "zdvs_blue_aspect: ";
                //string zdvs_approch = "zdvs_approch: ";
                string zdvs_libres_cycle = "zdvs_libres_cycle: ";
                string blocking_unblocking_iti = "blocking_unblocking_iti: ";
                string duree_lib_approch = "duree_lib_approch: ";
                string foul_zdv = "foul_zdv: ";
                //List<string> Subroute = new List<string> { "zdv: ", "itis: ", "transit_amont: " };
                idx = 11;

                // To find Fleet Route or not
                string Fleet_Check(string fleet_name, Excel.Range fleet_c)
                {
                    for (int i = 11; i < r_RF_6_count; i++)
                    {
                        if (fleet_c.Cells[i, 1].Value2 != null && fleet_c.Cells[i, 1].Value2.ToString().Contains("("))
                        {
                            break;
                        }

                        if (fleet_c.Cells[i, 1].Value2 != null && (fleet_name == fleet_c.Cells[i, 1].Value2.ToString()))
                        {
                            return "Y";
                        }
                    }
                    return "N";
                }

                // For transit_iti Dictionary from "CT12_Subroute Locked" sheet

                flag = 0;
                int flag_ack = 0;
                //int flag_ack_Shunt = 0;
                //int flag_Shunt = 0;
                for (int k = 11; k <= r1_count; k++)
                {
                    if (xlRange1.Cells[k, 1].Value2 != null && xlRange1.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                    if (xlRange1.Cells[k, 1].Value2 != null)
                    {
                        if (refe != k)
                        {
                            transit_iti[xlRange1.Cells[refe, 1].Value2.ToString()] = new List<string>(list_TC);
                        }
                        refe = k;
                        list_TC.Clear();

                    }
                    if (xlRange1.Cells[k, 2].Value2 != null)
                    {
                        list_TC.Add("U" + xlRange1.Cells[k, 2].Value2.ToString());
                    }
                    else
                    {
                        transit_iti[xlRange1.Cells[refe, 1].Value2.ToString()] = new List<string>(list_TC);
                        list_TC.Clear();
                        break;
                    }
                }

                // For transit_over and Signal Aspects track circuits from "CT151_Main Signal Proceed Asp." sheet

                refe = 11;
                int refe_shunt = 11;
                int flag_break = 0;
                if (r_MS_15_count > r_SS_15_count)
                {
                    flag_break = 0;
                }
                else
                {
                    flag_break |= 1;
                }
                //List<string> list_OL = new List<string>();
                List<string> list_TC_Shunt = new List<string>();
                for (int k = 11; k < Math.Max(r_MS_15_count, r_SS_15_count); k++)
                {

                    if (xlRange_MS_151.Cells[k, 2].Value2 != null)
                    {
                        if (refe != k || xlRange_MS_151.Cells[k, 9].Value2.ToString() == "(9)")
                        {
                            if (!transit_over.ContainsKey(xlRange_MS_151.Cells[refe, 2].Value2.ToString()))
                            {
                                transit_over[xlRange_MS_151.Cells[refe, 2].Value2.ToString()] = new List<string>(list_OL);
                            }
                        }

                        if (xlRange_MS_151.Cells[k, 14].Value2.ToString() == "Green" || xlRange_MS_151.Cells[k, 14].Value2.ToString() == "(14)")
                        {
                            flag = 0;
                        }
                        else if (xlRange_MS_151.Cells[k, 14].Value2.ToString() == "Violet")
                        {
                            flag = 1;
                        }
                        if (flag != flag_ack && flag == 1)
                        {
                            zdvs_green_aspect[xlRange_MS_151.Cells[refe, 2].Value2.ToString()] = new List<string>(list_TC);
                        }
                        if ((flag != flag_ack && flag == 0) || xlRange_MS_151.Cells[refe, 2].Value2.ToString().Contains("("))
                        {
                            zdvs_blue_aspect[xlRange_MS_151.Cells[refe, 2].Value2.ToString()] = new List<string>(list_TC);
                        }
                        refe = k;
                        flag_ack = flag;
                        list_TC.Clear();
                        list_OL.Clear();
                    }
                    if (xlRange_SS_152.Cells[k, 2].Value2 != null)
                    {

                        /*if (xlRange_SS_152.Cells[k, 10].Value2.ToString() == "Green" || xlRange_SS_152.Cells[k, 10].Value2.ToString() == "(10)")
                        {
                            flag_Shunt = 0;
                        }
                        else if (xlRange_SS_152.Cells[k, 10].Value2.ToString() == "Violet")
                        {
                        flag_Shunt = 1;
                        }*/
                        /*if (flag_Shunt != flag_ack_Shunt && flag == 1)
                        {
                            zdvs_green_aspect[xlRange_SS_152.Cells[refe_shunt, 2].Value2.ToString()] = new List<string>(list_TC_Shunt);
                        }*?
                        /*if ((flag_Shunt != flag_ack_Shunt && flag == 0) || xlRange_SS_152.Cells[refe_shunt, 2].Value2.ToString().Contains("("))
                        {
                            zdvs_blue_aspect[xlRange_SS_152.Cells[refe_shunt, 2].Value2.ToString()] = new List<string>(list_TC_Shunt);
                        }*/
                        zdvs_green_aspect[xlRange_SS_152.Cells[refe_shunt, 2].Value2.ToString()] = new List<string>(list_TC_Shunt);
                        refe_shunt = k;
                        //flag_ack_Shunt = flag_Shunt;
                        list_TC_Shunt.Clear();
                    }
                    if (xlRange_MS_151.Cells[k, 6].Value2 != null && xlRange_MS_151.Cells[k, 6].Value2.ToString() != "Nil")
                    {
                        list_TC.Add("TC" + xlRange_MS_151.Cells[k, 6].Value2.ToString());
                    }
                    if (xlRange_MS_151.Cells[k, 9].Value2 != null && xlRange_MS_151.Cells[k, 9].Value2.ToString() != "Nil")
                    {
                        list_OL.Add("OL" + xlRange_MS_151.Cells[k, 9].Value2.ToString());
                    }
                    if (xlRange_SS_152.Cells[k, 6].Value2 != null && xlRange_SS_152.Cells[k, 6].Value2.ToString() != "Nil")
                    {
                        list_TC_Shunt.Add("TC" + xlRange_SS_152.Cells[k, 6].Value2.ToString());
                    }
                    if (flag_break == 0 && xlRange_MS_151.Cells[k, 1].Value2 != null && xlRange_MS_151.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                    if (flag_break == 1 && xlRange_SS_152.Cells[k, 1].Value2 != null && xlRange_SS_152.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                    //Console.WriteLine(k);
                }

                // For zdvs_approch from "CT16_Route Released" sheet

                refe = 11;
                for (int k = 11; k <= r_RR_16_count; k++)
                {

                    if (xlRange_RR_16.Cells[k, 1].Value2 != null)
                    {
                        if (refe != k)
                        {
                            zdvs_approch[xlRange_RR_16.Cells[refe, 1].Value2.ToString()] = new List<string>(list_TC);
                        }
                        refe = k;
                        list_TC.Clear();
                    }
                    if (xlRange_RR_16.Cells[k, 3].Value2 != null)
                    {
                        list_TC.Add("TC" + xlRange_RR_16.Cells[k, 3].Value2.ToString());
                    }
                    if (xlRange_RR_16.Cells[k, 1].Value2 != null && xlRange_RR_16.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                }

                // remaining route property loop

                for (int i = 11; i <= r_RI_5_count; i++)
                {
                    if (xlRange_RI_5.Cells[i, 11].Value2 != null && xlRange_RI_5.Cells[i, 11].Value2.ToString() != "Nil")
                    {
                        transits_over_opp += "OL" + xlRange_RI_5.Cells[i, 11].Value2.ToString() + " ";
                    }

                    if (xlRange_RI_5.Cells[i, 9].Value2 != null && xlRange_RI_5.Cells[i, 9].Value2.ToString() != "Nil")
                    {
                        maintenance_block += "MBL" + xlRange_RI_5.Cells[i, 9].Value2.ToString() + " ";
                    }

                    if (xlRange_RI_5.Cells[i, 10].Value2 != null && xlRange_RI_5.Cells[i, 10].Value2.ToString() != "Nil")
                    {
                        transits_iti_incomp += "U" + xlRange_RI_5.Cells[i, 10].Value2.ToString() + " ";
                    }

                    if (xlRange_RI_5.Cells[i, 4].Value2 != null && xlRange_RI_5.Cells[i, 4].Value2.ToString() != "Nil")
                    {
                        itis_incomp += "R" + xlRange_RI_5.Cells[i, 4].Value2.ToString() + " ";
                    }

                    if (xlRange_RI_5.Cells[i, 12].Value2 != null && xlRange_RI_5.Cells[i, 12].Value2.ToString() != "Nil")
                    {
                        cycs_incomp += "CY" + xlRange_RI_5.Cells[i, 12].Value2.ToString() + " ";
                    }

                    if (xlRange_RI_5.Cells[i, 2].Value2 != null && xlRange_RI_5.Cells[i, 2].Value2.ToString() != "Nil")
                    {
                        if (xlRange_RI_5.Cells[i, 2].Value2.ToString().Length > 4)
                        {
                            aigs_N += "P" + xlRange_RI_5.Cells[i, 2].Value2.ToString().Split('_')[0] + " P" + xlRange_RI_5.Cells[i, 2].Value2.ToString().Split('_')[1] + " ";
                        }
                        else
                        {
                            aigs_N += "P" + xlRange_RI_5.Cells[i, 2].Value2.ToString() + " ";
                        }
                    }

                    if (xlRange_RI_5.Cells[i, 3].Value2 != null && xlRange_RI_5.Cells[i, 3].Value2.ToString() != "Nil")
                    {
                        if (xlRange_RI_5.Cells[i, 3].Value2.ToString().Length > 4)
                        {
                            aigs_R += "P" + xlRange_RI_5.Cells[i, 3].Value2.ToString().Split('_')[0] + " P" + xlRange_RI_5.Cells[i, 3].Value2.ToString().Split('_')[1] + " ";
                        }
                        else
                        {
                            aigs_R += "P" + xlRange_RI_5.Cells[i, 3].Value2.ToString() + " ";
                        }
                    }

                    /*if (xlRange_RI_5.Cells[i, 9].Value2 != null && xlRange_RI_5.Cells[i, 9].Value2.ToString() != "Nil")
                    {
                        zdv += "TC" + xlRange_RI_5.Cells[i, 9].Value2.ToString() + " ";
                    }
                    if (xlRange_RI_5.Cells[i, 3].Value2 != null && xlRange_RI_5.Cells[i, 3].Value2.ToString() != "Nil")
                    {
                        transit_amont += "U" + xlRange_RI_5.Cells[i, 3].Value2.ToString() + " ";
                    }*/
                    if (xlRange_RI_5.Cells[i, 1].Value2 != null && xlRange_RI_5.Cells[i, 1].Value2.ToString().Contains("("))
                    {
                        //Console.WriteLine(i);
                        break;
                    }

                    // Main logic for Route and writing file

                    if (xlRange_RI_5.Cells[i + 1, 1].Value2 != null)
                    {


                        //for ZDV_TORR_Ta,Tb,TC
                        for (int j = 11; j <= r_RR_16_count; j++)
                        {
                            if (xlRange_RR_16.Cells[j, 1].Value2 != null && xlRange_RR_16.Cells[j, 1].Value2.ToString().Contains("("))
                            {
                                break;
                            }

                            if (xlRange_RR_16.Cells[j, 1].Value2 != null && xlRange_RI_5.Cells[idx, 1].Value2.ToString() == xlRange_RR_16.Cells[j, 1].Value2.ToString())
                            {
                                zdv_TORR_Ta += "TC" + xlRange_RR_16.Cells[j, 14].Value2.ToString() + " ";
                                zdv_TORR_Tb += "TC" + xlRange_RR_16.Cells[j, 15].Value2.ToString() + " ";
                                zdv_TORR_Tc += "TC" + xlRange_RR_16.Cells[j, 17].Value2.ToString() + " ";
                            }
                        }

                        //For Fleet check

                        trace_perm += Fleet_Check(xlRange_RI_5.Cells[idx, 1].Value2.ToString(), xlRange_RF_6);


                        using (StreamWriter sw = Route_File.AppendText())
                        {

                            sw.WriteLine("R" + xlRange_RI_5.Cells[idx, 1].Value2.ToString());
                            sw.WriteLine("feu_orig: S" + xlRange_RI_5.Cells[idx, 1].Value2.ToString().Split('_')[0]);
                            if (transit_iti.ContainsKey(xlRange_RI_5.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("transits_iti: " + string.Join(" ", transit_iti[xlRange_RI_5.Cells[idx, 1].Value2.ToString()]) + " .");
                            }
                            else
                            {
                                sw.WriteLine("transits_iti: .");
                            }
                            if (transit_over.ContainsKey(xlRange_RI_5.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("transits_over: " + string.Join(" ", transit_over[xlRange_RI_5.Cells[idx, 1].Value2.ToString()]) + ".");
                            }
                            else
                            {
                                sw.WriteLine("transits_over: .");
                            }
                            sw.WriteLine(transits_over_opp + ".");
                            sw.WriteLine(maintenance_block + ".");
                            sw.WriteLine(transits_iti_incomp + ".");
                            sw.WriteLine(itis_incomp + ".");
                            sw.WriteLine(cycs_incomp + ".");
                            sw.WriteLine(aigs_N + ".");
                            sw.WriteLine(aigs_R + ".");
                            sw.WriteLine(zdv_TORR_Ta + ".");
                            sw.WriteLine(zdv_TORR_Tb + ".");
                            sw.WriteLine(zdv_TORR_Tc + ".");
                            sw.WriteLine(trace_perm);
                            //Console.WriteLine(zdvs_green_aspect[xlRange.Cells[idx, 1].Value2.ToString()]);
                            //Console.WriteLine(string.Join(" ", zdvs_green_aspect[xlRange.Cells[idx, 1].Value2.ToString()]));
                            if (zdvs_green_aspect.ContainsKey(xlRange_RI_5.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("zdvs_green_aspect: " + string.Join(" ", zdvs_green_aspect[xlRange_RI_5.Cells[idx, 1].Value2.ToString()]) + " .");
                            }
                            else
                            {
                                sw.WriteLine("zdvs_green_aspect: .");
                            }
                            if (zdvs_blue_aspect.ContainsKey(xlRange_RI_5.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("zdvs_blue_aspect: " + string.Join(" ", zdvs_blue_aspect[xlRange_RI_5.Cells[idx, 1].Value2.ToString()]) + " .");
                            }
                            else
                            {
                                sw.WriteLine("zdvs_blue_aspect: .");
                            }
                            if (zdvs_approch.ContainsKey(xlRange_RI_5.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("zdvs_approch: " + string.Join(" ", zdvs_approch[xlRange_RI_5.Cells[idx, 1].Value2.ToString()]) + " .");
                            }

                            else
                            {
                                sw.WriteLine("zdvs_approch: .");
                            }
                            sw.WriteLine(zdvs_libres_cycle + ".");
                            sw.WriteLine(blocking_unblocking_iti + "RBU" + Sector_Entry.Text + " .");
                            sw.WriteLine(duree_lib_approch + "90");
                            sw.WriteLine(foul_zdv + ".");
                            sw.WriteLine("%\n");
                        }

                        idx = i + 1;
                        flag = 0;
                        //transit_iti = "transits_iti: ";
                        //transits_over = "transits_over: ";
                        transits_over_opp = "transits_over_opp: ";
                        maintenance_block = "maintenance_block: ";
                        transits_iti_incomp = "transits_iti_incomp: ";
                        itis_incomp = "itis_incomp: ";
                        cycs_incomp = "cycs_incomp: ";
                        aigs_N = "aigs_N: ";
                        aigs_R = "aigs_R: ";
                        zdv_TORR_Ta = "zdv_TORR_Ta: ";
                        zdv_TORR_Tb = "zdv_TORR_Tb: ";
                        zdv_TORR_Tc = "zdv_TORR_Tc: ";
                        trace_perm = "trace_perm: ";
                        //zdvs_green_aspect = "zdvs_green_aspect: ";
                        //zdvs_blue_aspect = "zdvs_blue_aspect: ";
                        //zdvs_approch = "zdvs_approch: ";
                        zdvs_libres_cycle = "zdvs_libres_cycle: ";
                        blocking_unblocking_iti = "blocking_unblocking_iti: ";
                        duree_lib_approch = "duree_lib_approch: ";
                        foul_zdv = "foul_zdv: ";
                    }

                }
                //richTextBox1.Text += "Route Completed\n";
                UpdateRichTextBox("Route Completed\n");
            }
        }
        //GC.Collect();
        //GC.WaitForPendingFinalizers();
        //
        //Marshal.ReleaseComObject(xlRange);
        //Marshal.ReleaseComObject(ws);
        //
        //Marshal.ReleaseComObject(xlRange1);
        //Marshal.ReleaseComObject(ws1);
        //
        //Marshal.ReleaseComObject(xlRange_RR_16);
        //Marshal.ReleaseComObject(ws_RR_16);
        //
        //Marshal.ReleaseComObject(xlRange_MS_151);
        //Marshal.ReleaseComObject(ws_MS_151);
        //
        //Marshal.ReleaseComObject(xlRange_RF_6);
        //Marshal.ReleaseComObject(ws_RF_6);


        //Point
        private void Points(int r_count_MA_9,Excel.Range xlRange_MA_9,int r_count_PL_10,Excel.Range xlRange_PL_10,FileInfo PAR_File) {
            if (Point.Checked == true)
            {
                Thread.Sleep(500);
                string transit_N = "";
                string transit_R = "";
                Dictionary<string, List<string>> signaux_selon_aig = new Dictionary<string, List<string>>();
                string zdvs = "";
                string overlaps_N = "";
                string overlaps_R = "";
                flag = 0;
                refe = 11;
                idx = 11;

                string auto_normal_check(string name, Range sheet, int id)
                {
                    if (name == sheet.Cells[id, 15].Value2.ToString())
                    {
                        return "Y";
                    }
                    else
                    {
                        return "N";
                    }
                }
                refe = 11;
                for (int k = 11; k < r_count_MA_9; k++)
                {
                    if (xlRange_MA_9.Cells[k, 1].Value2 != null)
                    {
                        if (refe != k)
                        {
                            if (xlRange_MA_9.Cells[refe, 1].Value2.ToString().Split('_').Length == 2)
                            {
                                signaux_selon_aig[xlRange_MA_9.Cells[refe, 1].Value2.ToString().Split('_')[0]] = new List<string>(list_TC);
                                signaux_selon_aig[xlRange_MA_9.Cells[refe, 1].Value2.ToString().Split('_')[1]] = new List<string>(list_TC);
                            }
                            else
                            {
                                signaux_selon_aig[xlRange_MA_9.Cells[refe, 1].Value2.ToString()] = new List<string>(list_TC);
                            }
                        }
                        refe = k;
                        list_TC.Clear();
                    }
                    if (xlRange_MA_9.Cells[k, 2].Value2 != null && xlRange_MA_9.Cells[k, 2].Value2.ToString() != "Nil")
                    {
                        list_TC.Add("S" + xlRange_MA_9.Cells[k, 2].Value2.ToString());
                    }
                    if (xlRange_MA_9.Cells[k, 1].Value2 != null && xlRange_MA_9.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        //Console.WriteLine(i);
                        break;
                    }
                }
                refe = 11;
                int count = 0;
                for (int k = 11; k <= r_count_PL_10; k++)
                {

                    if (xlRange_PL_10.Cells[k, 2].Value2 != null && xlRange_PL_10.Cells[k, 2].Value2.ToString() == "nr")
                    {
                        flag = 1;
                        ++count;
                    }
                    if (xlRange_PL_10.Cells[k, 2].Value2 != null && xlRange_PL_10.Cells[k, 2].Value2.ToString() == "rn")
                    {
                        flag = 0;
                        ++count;
                    }


                    if (xlRange_PL_10.Cells[k, 3].Value2 != null && xlRange_PL_10.Cells[k, 3].Value2.ToString() != "Nil" && flag == 1)
                    {
                        zdvs += "TC" + xlRange_PL_10.Cells[k, 3].Value2.ToString() + " ";
                    }
                    if (xlRange_PL_10.Cells[k, 13].Value2 != null && xlRange_PL_10.Cells[k, 13].Value2.ToString() != "Nil" && flag == 1)
                    {
                        overlaps_N += "OL" + xlRange_PL_10.Cells[k, 13].Value2.ToString() + " ";
                    }

                    if (xlRange_PL_10.Cells[k, 13].Value2 != null && xlRange_PL_10.Cells[k, 13].Value2.ToString() != "Nil" && flag == 0)
                    {
                        overlaps_R += "OL" + xlRange_PL_10.Cells[k, 13].Value2.ToString() + " ";
                    }
                    if (xlRange_PL_10.Cells[k, 12].Value2 != null && xlRange_PL_10.Cells[k, 12].Value2.ToString() != "Nil" && flag == 1)
                    {
                        transit_N += "U" + xlRange_PL_10.Cells[k, 12].Value2.ToString() + " ";
                    }
                    if (xlRange_PL_10.Cells[k, 12].Value2 != null && xlRange_PL_10.Cells[k, 12].Value2.ToString() != "Nil" && flag == 0)
                    {
                        transit_R += "U" + xlRange_PL_10.Cells[k, 12].Value2.ToString() + " ";
                    }
                    if (xlRange_PL_10.Cells[k + 1, 1].Value2 != null)
                    {
                        if (count == 2)
                        {
                            if (xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_').Length == 1)
                            {
                                using (StreamWriter sw = PAR_File.AppendText())
                                {
                                    sw.WriteLine("P" + xlRange_PL_10.Cells[idx, 1].Value2.ToString());
                                    sw.WriteLine("transit_N: " + transit_N + " .");
                                    sw.WriteLine("transit_R: " + transit_R + " .");
                                    sw.WriteLine("pointKm: ");
                                    sw.WriteLine("blocking_unblocking_aiguille: PBU" + Sector_Entry.Text);
                                    sw.WriteLine("auto_normalisation: " + auto_normal_check(xlRange_PL_10.Cells[idx, 1].Value2.ToString(), xlRange_PL_10, idx) + " .");
                                    sw.WriteLine("duree_auto_normalisation: 0,5 .");
                                    sw.WriteLine("auto_normalisation_activee: Y .");
                                    sw.WriteLine("position_auto_normalisation_N: Y .");
                                    sw.WriteLine("aiguille_conjuguee: .");
                                    sw.WriteLine("zdvs: " + zdvs + " .");
                                    sw.WriteLine("overlaps_N: " + overlaps_N + " .");
                                    sw.WriteLine("overlaps_R: " + overlaps_R + " .");
                                    sw.WriteLine("aiguille_trailing: .");
                                    if (signaux_selon_aig.ContainsKey(xlRange_PL_10.Cells[idx, 1].Value2.ToString()))
                                    {
                                        sw.WriteLine("signaux_selon_aig: " + string.Join(" ", signaux_selon_aig[xlRange_PL_10.Cells[idx, 1].Value2.ToString()]) + " .");
                                    }
                                    else
                                    {
                                        sw.WriteLine("signaux_selon_aig: .");
                                    }
                                    sw.WriteLine("%\n");
                                }
                            }
                            else
                            {
                                using (StreamWriter sw = PAR_File.AppendText())
                                {
                                    sw.WriteLine("P" + xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[0]);
                                    sw.WriteLine("transit_N: " + transit_N + " .");
                                    sw.WriteLine("transit_R: " + transit_R + " .");
                                    sw.WriteLine("pointKm: ");
                                    sw.WriteLine("blocking_unblocking_aiguille: PBU" + Sector_Entry.Text);
                                    sw.WriteLine("auto_normalisation: " + auto_normal_check(xlRange_PL_10.Cells[idx, 1].Value2.ToString(), xlRange_PL_10, idx) + " .");
                                    sw.WriteLine("duree_auto_normalisation: 0,5 .");
                                    sw.WriteLine("auto_normalisation_activee: Y .");
                                    sw.WriteLine("position_auto_normalisation_N: Y .");
                                    sw.WriteLine("aiguille_conjuguee: P" + xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[1] + " .");
                                    sw.WriteLine("zdvs: " + zdvs + " .");
                                    sw.WriteLine("overlaps_N: " + overlaps_N + " .");
                                    sw.WriteLine("overlaps_R: " + overlaps_R + " .");
                                    sw.WriteLine("aiguille_trailing: .");
                                    if (signaux_selon_aig.ContainsKey(xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[0]))
                                    {
                                        sw.WriteLine("signaux_selon_aig: " + string.Join(" ", signaux_selon_aig[xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[0]]) + " .");
                                    }
                                    else
                                    {
                                        sw.WriteLine("signaux_selon_aig: .");
                                    }
                                    sw.WriteLine("%\n");

                                    sw.WriteLine("P" + xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[1]);
                                    sw.WriteLine("transit_N: " + transit_N + " .");
                                    sw.WriteLine("transit_R: " + transit_R + " .");
                                    sw.WriteLine("pointKm: ");
                                    sw.WriteLine("blocking_unblocking_aiguille: PBU" + Sector_Entry.Text);
                                    sw.WriteLine("auto_normalisation: " + auto_normal_check(xlRange_PL_10.Cells[idx, 1].Value2.ToString(), xlRange_PL_10, idx) + " .");
                                    sw.WriteLine("duree_auto_normalisation: 0,5 .");
                                    sw.WriteLine("auto_normalisation_activee: Y .");
                                    sw.WriteLine("position_auto_normalisation_N: Y .");
                                    sw.WriteLine("aiguille_conjuguee: P" + xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[0] + " .");
                                    sw.WriteLine("zdvs: " + zdvs + " .");
                                    sw.WriteLine("overlaps_N: " + overlaps_N + " .");
                                    sw.WriteLine("overlaps_R: " + overlaps_R + " .");
                                    sw.WriteLine("aiguille_trailing: .");
                                    if (signaux_selon_aig.ContainsKey(xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[1]))
                                    {
                                        sw.WriteLine("signaux_selon_aig: " + string.Join(" ", signaux_selon_aig[xlRange_PL_10.Cells[idx, 1].Value2.ToString().Split('_')[1]]) + " .");
                                    }
                                    else
                                    {
                                        sw.WriteLine("signaux_selon_aig: .");
                                    }
                                    sw.WriteLine("%\n");
                                }
                            }
                            count = 0;
                            transit_N = "";
                            transit_R = "";
                            overlaps_N = "";
                            overlaps_R = "";
                            zdvs = "";
                        }

                        idx = k + 1;


                    }

                    if (xlRange_PL_10.Cells[k, 1].Value2 != null && xlRange_PL_10.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        //Console.WriteLine(i);
                        break;
                    }



                }
                //richTextBox1.Text += "Point Completed\n";
                UpdateRichTextBox("Point Completed\n");
            }
        }

        // 3 Aspect Signal
        private void Three_Aspect_Signals(int r_RI_17_count, Excel.Range xlRange_RI_17, int r_MS_15_count, Excel.Range xlRange_MS_151, FileInfo PAR_File)
        {
            if (Three_Aspect_Signal.Checked == true)
            {
                //string aiguilles = "";
                Dictionary<string, List<string>> itinerairesRI = new Dictionary<string, List<string>>();
                Dictionary<string, HashSet<string>> itineraires = new Dictionary<string, HashSet<string>>();
                Dictionary<string, HashSet<string>> aiguilles = new Dictionary<string, HashSet<string>>();
                HashSet<string> hashset_route = new HashSet<string>();
                HashSet<string> hashset_point = new HashSet<string>();
                HashSet<string> hashset_signal = new HashSet<string>();


                refe = 11;
                flag = 0;
                list_TC.Clear();
                for (int k = 11; k < r_RI_17_count; k++)
                {
                    if (xlRange_RI_17.Cells[k, 1].Value2 != null && xlRange_RI_17.Cells[k, 1].Value2.ToString().Split('_').Length == 2 && xlRange_RI_17.Cells[k, 1].Value2.ToString().Split('_')[1] == "M")
                    {
                        flag = 0;
                    }

                    if (xlRange_RI_17.Cells[k, 1].Value2 != null && xlRange_RI_17.Cells[k, 1].Value2.ToString().Split('_').Length == 2 && xlRange_RI_17.Cells[k, 1].Value2.ToString().Split('_')[1] != "M")
                    {
                        flag = 1;
                    }
                    if (xlRange_RI_17.Cells[k, 1].Value2 != null && xlRange_RI_17.Cells[k, 1].Value2.ToString().Split('_').Length == 2)
                    {
                        if (refe != k && xlRange_RI_17.Cells[k, 1].Value2.ToString().Split('_')[0] != xlRange_RI_17.Cells[refe, 1].Value2.ToString().Split('_')[0])
                        {
                            itinerairesRI[xlRange_RI_17.Cells[refe, 1].Value2.ToString().Split('_')[0]] = new List<string>(list_TC);
                            list_TC.Clear();
                        }
                        refe = k;


                    }
                    /*if (xlRange_RI_17.Cells[k, 1].Value2 != null && ! itinerairesRI.ContainsKey(xlRange_RI_17.Cells[k, 1].Value2.ToString().Split('_')[0]))
                    {
                        list_TC.Clear();
                    }*/
                    if (xlRange_RI_17.Cells[k, 2].Value2 != null && flag == 1 && !xlRange_RI_17.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        list_TC.Add("R" + xlRange_RI_17.Cells[k, 2].Value2.ToString());
                    }
                    if (xlRange_RI_17.Cells[k, 1].Value2 != null && xlRange_RI_17.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        itinerairesRI[xlRange_RI_17.Cells[refe, 1].Value2.ToString().Split('_')[0]] = new List<string>(list_TC);
                        list_TC.Clear();
                        break;
                    }
                }
                idx = 11;
                flag = 0;
                int flag_a = 0;
                for (int k = 11; k < r_MS_15_count; k++)
                {
                    if (xlRange_MS_151.Cells[k, 2].Value2 != null && xlRange_MS_151.Cells[k, 2].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                    if (xlRange_MS_151.Cells[k, 2].Value2 != null && xlRange_MS_151.Cells[k, 2].Value2.ToString() != "Nil")
                    {
                        //itineraires[xlRange_MS_151.Cells[idx,1].Value2].add( "R" + xlRange_MS_151.Cells[k, 2].Value2.ToString() + " ");
                        hashset_signal.Add(xlRange_MS_151.Cells[k, 1].Value2.ToString());
                        if (!itineraires.ContainsKey(xlRange_MS_151.Cells[k, 1].Value2.ToString()))
                        {
                            hashset_route.Add("R" + xlRange_MS_151.Cells[k, 2].Value2.ToString() + " ");
                            flag = 0;
                        }
                        else
                        {
                            itineraires[xlRange_MS_151.Cells[idx, 1].Value2].Add("R" + xlRange_MS_151.Cells[k, 2].Value2.ToString() + " ");
                            flag = 1;
                        }
                    }
                    if (xlRange_MS_151.Cells[k, 3].Value2 != null && xlRange_MS_151.Cells[k, 3].Value2.ToString() != "Nil")
                    {
                        if (xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_').Length == 1)
                        {
                            //aiguilles += "P" + xlRange_MS_151.Cells[k, 3].Value2.ToString() + " ";
                            if (!aiguilles.ContainsKey(xlRange_MS_151.Cells[idx, 1].Value2.ToString()))
                            {
                                hashset_point.Add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString() + " ");
                                flag_a = 0;
                            }
                            else
                            {
                                aiguilles[xlRange_MS_151.Cells[idx, 1].Value2].Add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString() + " ");
                                flag_a = 1;
                            }
                        }

                        if (xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_').Length == 2)
                        {
                            //aiguilles[xlRange_MS_151.Cells[idx,1].Value2].add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_')[0] + " ");
                            //aiguilles[xlRange_MS_151.Cells[idx,1].Value2].add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_')[1] + " ");
                            if (!aiguilles.ContainsKey(xlRange_MS_151.Cells[idx, 1].Value2.ToString()))
                            {
                                hashset_point.Add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_')[0] + " ");
                                hashset_point.Add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_')[1] + " ");
                                flag_a = 0;
                            }
                            else
                            {
                                aiguilles[xlRange_MS_151.Cells[idx, 1].Value2].add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_')[0] + " ");
                                aiguilles[xlRange_MS_151.Cells[idx, 1].Value2].add("P" + xlRange_MS_151.Cells[k, 3].Value2.ToString().Split('_')[1] + " ");
                                flag_a = 1;
                            }
                        }
                    }
                    if (xlRange_MS_151.Cells[k + 1, 1].Value2 != null && xlRange_MS_151.Cells[k + 1, 1].Value2 != xlRange_MS_151.Cells[idx, 1].Value2)
                    {
                        if (flag == 0)
                        {
                            itineraires[xlRange_MS_151.Cells[idx, 1].Value2.ToString()] = new HashSet<string>(hashset_route);
                        }
                        if (flag_a == 0)
                        {
                            aiguilles[xlRange_MS_151.Cells[idx, 1].Value2.ToString()] = new HashSet<string>(hashset_point);
                        }
                        hashset_route.Clear();
                        hashset_point.Clear();

                        idx = k + 1;



                    }

                }
                foreach (string k in hashset_signal)
                {
                    using (StreamWriter sw = PAR_File.AppendText())
                    {
                        sw.WriteLine("S" + k);
                        sw.WriteLine("global_blocking_unblocking_signal: SBU" + Sector_Entry.Text);
                        Console.WriteLine(k);
                        sw.WriteLine("itineraires: " + string.Join(" ", itineraires[k]) + ".");
                        if (itinerairesRI.ContainsKey(k))
                        {
                            sw.WriteLine("itinerairesRI: " + string.Join(" ", itinerairesRI[k]) + " .");
                        }
                        else
                        {
                            sw.WriteLine("itinerairesRI: .");
                        }
                        sw.WriteLine("indicateur_de_direction: .");
                        sw.WriteLine("aiguilles: " + string.Join(" ", aiguilles[k]) + " .");
                        sw.WriteLine("%\n");
                    }


                }
                //richTextBox1.Text += "3 Aspect Signal Completed\n";
                UpdateRichTextBox("3 Aspect Signal Completed\n");
            }
        }

            // Cycle
            private void Cycles(int r_count_CC_3,Excel.Range xlRange_CC_3,FileInfo PAR_File) 
            {
                if (Cycle.Checked == true)
                {
                    Thread.Sleep(500);
                    string itineraires_cycle = "";
                    string aiguilles_cycle = "";
                    string itineraires_incompatibles = "";
                    string transits_Iti_incompatibles = "";
                    idx = 11;
                    for (int k = 11; k < r_count_CC_3; k++)
                    {
                        if (xlRange_CC_3.Cells[k, 2].Value2 != null && xlRange_CC_3.Cells[k, 2].Value2.ToString() != "Nil")
                        {
                            itineraires_cycle += "R" + xlRange_CC_3.Cells[k, 2].Value2.ToString() + " ";
                        }

                        if (xlRange_CC_3.Cells[k, 7].Value2 != null && xlRange_CC_3.Cells[k, 7].Value2.ToString() != "Nil")
                        {
                            if (xlRange_CC_3.Cells[k, 7].Value2.ToString().Split('_').Length == 1)
                            {
                                aiguilles_cycle += "P" + xlRange_CC_3.Cells[k, 7].Value2.ToString() + " ";
                            }
                            else
                            {
                                aiguilles_cycle += "P" + xlRange_CC_3.Cells[k, 7].Value2.ToString().Split('_')[0] + " " + "P" + xlRange_CC_3.Cells[k, 7].Value2.ToString().Split('_')[1] + " ";
                            }

                        }

                        if (xlRange_CC_3.Cells[k, 3].Value2 != null && xlRange_CC_3.Cells[k, 3].Value2.ToString() != "Nil")
                        {
                            itineraires_incompatibles += "R" + xlRange_CC_3.Cells[k, 3].Value2.ToString() + " ";
                        }

                        if (xlRange_CC_3.Cells[k, 9].Value2 != null && xlRange_CC_3.Cells[k, 9].Value2.ToString() != "Nil")
                        {
                            transits_Iti_incompatibles += "U" + xlRange_CC_3.Cells[k, 9].Value2.ToString() + " ";
                        }

                        if (xlRange_CC_3.Cells[k + 1, 1].Value2 != null)
                        {
                            using (StreamWriter sw = PAR_File.AppendText())
                            {
                                sw.WriteLine("CY" + xlRange_CC_3.Cells[idx, 1].Value2.ToString());
                                sw.WriteLine("itineraires_cycle: " + itineraires_cycle + ".");
                                sw.WriteLine("aiguilles_cycle: " + aiguilles_cycle + ".");
                                sw.WriteLine("cycles_incompatibles: .");
                                sw.WriteLine("itineraires_incompatibles: " + itineraires_incompatibles + ".");
                                sw.WriteLine("transits_Iti_incompatibles: " + transits_Iti_incompatibles + ".");
                                sw.WriteLine("%\n");
                            }
                            idx = k + 1;
                            itineraires_cycle = "";
                            aiguilles_cycle = "";
                            itineraires_incompatibles = "";
                            transits_Iti_incompatibles = "";

                        }
                        if (xlRange_CC_3.Cells[k + 1, 1].Value2 != null && xlRange_CC_3.Cells[k + 1, 1].Value2.ToString().Contains("("))
                        {
                            break;
                        }

                    }
                    //richTextBox1.Text += "Cycle Completed\n";
                    UpdateRichTextBox("Cycle Completed\n");
                } 
            }

        // Shunt Signal
        private void Shunt_Signals(int r_SS_15_count,Excel.Range xlRange_SS_152,FileInfo PAR_File) 
        {
            flag = 0;
            if (Shunt_Signal.Checked == true)
            {
                Thread.Sleep(500);
                Dictionary<string, HashSet<string>> itineraires = new Dictionary<string, HashSet<string>>();
                HashSet<string> hashset_route = new HashSet<string>();
                HashSet<string> hashset_signal = new HashSet<string>();
                //itineraires = "";
                idx = 11;
                for (int k = 11; k < r_SS_15_count; k++)
                {
                    if (xlRange_SS_152.Cells[k, 1].Value2 != null && xlRange_SS_152.Cells[k, 1].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                    if (xlRange_SS_152.Cells[k, 2].Value2 != null)
                    {
                        //itineraires[xlRange_SS_152.Cells[idx, 1].Value2].add("R" + xlRange_SS_152.Cells[k, 2].Value2.ToString() + " ");
                        hashset_signal.Add(xlRange_SS_152.Cells[k, 1].Value2.ToString());
                        if (!itineraires.ContainsKey(xlRange_SS_152.Cells[k, 1].Value2.ToString()))
                        {
                            hashset_route.Add("R" + xlRange_SS_152.Cells[k, 2].Value2.ToString() + " ");
                            flag = 0;
                        }
                        else
                        {

                            itineraires[xlRange_SS_152.Cells[idx, 1].Value2.ToString()].Add("R" + xlRange_SS_152.Cells[k, 2].Value2.ToString() + " ");
                            flag = 1;
                        }
                    }
                    if (xlRange_SS_152.Cells[k + 1, 1].Value2 != null && xlRange_SS_152.Cells[idx, 1].Value2 != xlRange_SS_152.Cells[k + 1, 1].Value2)
                    {
                        if (flag == 0)
                        {
                            itineraires[xlRange_SS_152.Cells[idx, 1].Value2.ToString()] = new HashSet<string>(hashset_route);
                        }
                        hashset_route.Clear();
                        idx = k + 1;
                        //itineraires = "";
                    }

                }
                foreach (string k in hashset_signal)
                {
                    using (StreamWriter sw = PAR_File.AppendText())
                    {
                        //Console.WriteLine(xlRange_SS_152.Cells[idx, 1].Value2.ToString());   
                        sw.WriteLine("S" + k);
                        sw.WriteLine("itineraires: " + string.Join(" ", itineraires[k]) + ".");
                        sw.WriteLine("global_blocking_unblocking_signal: SBU" + Sector_Entry.Text + " .");
                        sw.WriteLine("deux_aspects: Y");
                        sw.WriteLine("%\n");
                    }



                }
                UpdateRichTextBox("Shunt Signal Completed\n");
            }
        }

        // Overlap
        private void Overlaps(int r_count_OC_141,Excel.Range xlRange_OC_141,int r_count_OE_142,Excel.Range xlRange_OE_142,FileInfo PAR_File)
        {
            if (Overlap.Checked == true)
            {
                Dictionary<string, List<string>> transit_iti_fin_strike = new Dictionary<string, List<string>>();
                Dictionary<string, List<string>> zdvs_strike = new Dictionary<string, List<string>>();
                Dictionary<string, List<string>> zdv_fin_strike = new Dictionary<string, List<string>>();
                string aiguillesN = "";
                refe = 11;
                for (int k = 11; k < r_count_OC_141; k++)
                {
                    if (xlRange_OC_141.Cells[k, 1].Value2 != null)
                    {
                        if (refe != k)
                        {
                            zdvs_strike[xlRange_OC_141.Cells[refe, 1].Value2.ToString()] = new List<string>(list_TC);
                            transit_iti_fin_strike[xlRange_OC_141.Cells[refe, 1].Value2.ToString()] = new List<string>(list_U);
                            zdv_fin_strike[xlRange_OC_141.Cells[refe, 1].Value2.ToString()] = new List<string>(list_OL);
                        }
                        refe = k;
                        list_TC.Clear();
                        list_U.Clear();
                        list_OL.Clear();
                    }
                    if (xlRange_OC_141.Cells[k, 2].Value2 != null && xlRange_OC_141.Cells[k, 2].Value2.ToString() != "Nil")
                    {
                        list_TC.Add("TC" + xlRange_OC_141.Cells[k, 2].Value2.ToString());
                    }
                    if (xlRange_OC_141.Cells[k, 6].Value2 != null && xlRange_OC_141.Cells[k, 6].Value2.ToString() != "Nil")
                    {
                        list_U.Add("U" + xlRange_OC_141.Cells[k, 6].Value2.ToString());
                    }
                    if (xlRange_OC_141.Cells[k, 9].Value2 != null && xlRange_OC_141.Cells[k, 9].Value2.ToString() != "Nil")
                    {
                        list_OL.Add("TC" + xlRange_OC_141.Cells[k, 9].Value2.ToString());
                    }
                }
                idx = 11;
                for (int k = 11; k < r_count_OE_142; k++)
                {
                    if (xlRange_OE_142.Cells[k, 3].Value2 != null && xlRange_OE_142.Cells[k, 3].Value2.ToString() != "Nil")
                    {
                        if (xlRange_OE_142.Cells[k, 3].Value2.ToString().Split('_').Length == 1)
                        {
                            aiguillesN += "P" + xlRange_OE_142.Cells[k, 3].Value2.ToString() + " ";
                        }
                        else if (xlRange_OE_142.Cells[k, 3].Value2.ToString().Split('_').Length == 2)
                        {
                            aiguillesN += "P" + xlRange_OE_142.Cells[k, 3].Value2.ToString().Split('_')[0] + " P" + xlRange_OE_142.Cells[k, 3].Value2.ToString().Split('_')[1] + " ";
                        }
                    }
                    if (xlRange_OE_142.Cells[k + 1, 1].Value2 != null)
                    {
                        using (StreamWriter sw = PAR_File.AppendText())
                        {
                            sw.WriteLine("OL" + xlRange_OE_142.Cells[idx, 1].Value2.ToString());
                            sw.WriteLine("duree_liberation_overlap: 33");
                            sw.WriteLine("zdv: TC" + xlRange_OE_142.Cells[idx, 1].Value2.ToString().Split('_')[0] + ".");
                            if (zdv_fin_strike.ContainsKey(xlRange_OE_142.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("zdv_fin_strike: " + string.Join(" ", zdv_fin_strike[xlRange_OE_142.Cells[idx, 1].Value2.ToString()]) + ".");
                            }
                            else
                            {
                                sw.WriteLine("zdv_fin_strike: .");
                            }
                            if (transit_iti_fin_strike.ContainsKey(xlRange_OE_142.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("transit_iti_fin_strike: " + string.Join(" ", transit_iti_fin_strike[xlRange_OE_142.Cells[idx, 1].Value2.ToString()]) + ".");
                            }
                            else
                            {
                                sw.WriteLine("transit_iti_fin_strike: .");
                            }
                            if (zdvs_strike.ContainsKey(xlRange_OE_142.Cells[idx, 1].Value2.ToString()))
                            {
                                sw.WriteLine("zdvs_strike: " + string.Join(" ", zdvs_strike[xlRange_OE_142.Cells[idx, 1].Value2.ToString()]) + ".");
                            }
                            else
                            {
                                sw.WriteLine("zdvs_strike: .");
                            }
                            sw.WriteLine("aiguillesN: " + aiguillesN + ".");
                            sw.WriteLine("aiguillesR: .");
                            sw.WriteLine("itineraire_suivant: .");
                            sw.WriteLine("%\n");
                        }
                        aiguillesN = "";
                        idx = k + 1;
                    }
                    if (xlRange_OE_142.Cells[k + 1, 1].Value2 != null && xlRange_OE_142.Cells[k + 1, 1].Value2.ToString().Contains("("))
                    {
                        break;
                    }
                }
                UpdateRichTextBox("Overlap Completed\n");
            }
        }






        private void Mtk(int r_count_DI,Excel.Range xlRange_DI)
        {
            if (MTK.Checked == true)
            {
                Thread.Sleep(500);
                FileInfo MTK_File = new FileInfo(Output_Folder.Text + @"\MTK.xml");
                if (MTK_File.Exists == true)
                {
                    MTK_File.Delete();
                }
                List<string> mtk_f = new List<string>();
                using (StreamReader sw = new StreamReader(MTK_Entry.Text))
                {
                    string line;
                    while ((line = sw.ReadLine()) != null)
                    {
                        line = line.Trim();
                        mtk_f.Add(line);
                    }
                }
                /*foreach(string a in mtk_f)
                {
                    if (a.Contains(";"))
                    {
                        if(a.Split(';')[0] == "")
                        { 
                            Console.WriteLine(a.Split(';')[0]+"-->"+a.Split(';').Length);
                        }
                    }
                }*/
                Dictionary<string, string> DI = new Dictionary<string, string>();


                Dictionary<string, string> RM = new Dictionary<string, string>();

                RM.Add("TC", "TC");
                RM.Add("DPN", "P");
                RM.Add("DPR", "P");
                RM.Add("DPNL", "P");
                RM.Add("DPRL", "P");
                RM.Add("PPN", "P");
                RM.Add("PPR", "P");
                RM.Add("PREP_MPLA", "PBU");
                RM.Add("PREP_MPLDA", "PBU");
                RM.Add("MPL", "P");
                RM.Add("MAGP", "P");
                RM.Add("PMCK", "P");
                RM.Add("S_ON", "S");
                RM.Add("ALS", "S");
                RM.Add("S_OFF_V", "S");
                RM.Add("S_OFF_B", "S");
                RM.Add("LDS_G", "S");
                RM.Add("LDS_VR", "S");
                RM.Add("LDBS", "S");
                RM.Add("LDRI", "S");
                RM.Add("PREP_SBLA", "S");
                RM.Add("PREP_SUBLA", "S");
                RM.Add("SBL", "S");
                RM.Add("SS_OFF", "S");
                RM.Add("U", "U");
                RM.Add("TD", "TD");
                RM.Add("RL_S", "R");
                RM.Add("PREP_RBLA", "RBU");
                RM.Add("RBL", "R");
                RM.Add("FR", "R");
                RM.Add("MPS1", "AL");
                RM.Add("MPS2", "AL");
                RM.Add("UPS1", "AL");
                RM.Add("UPS2", "AL");
                RM.Add("AU_ACC", "AU_ACC");
                RM.Add("PSC", "AL");
                RM.Add("PREP_ISNA", "PSNAI");
                RM.Add("PREP_ASNA", "PSNAI");
                RM.Add("AMC", "AMC");
                RM.Add("VMC", "VMC");
                RM.Add("ISN", "P");
                RM.Add("PREP_RUBLA", "RBU");
                RM.Add("CY", "CY");
                RM.Add("OL", "OL");
                RM.Add("ESP", "ESP");
                RM.Add("UPS", "AL");
                RM.Add("MPS", "AL");
                RM.Add("MBL", "MBL");
                RM.Add("PREP_MBLA", "MBLSR");
                RM.Add("PREP_MBULA", "MBLSR");
                RM.Add("PSAPR", "PSAPR");
                RM.Add("LDSS_1", "S");
                RM.Add("LDSS_2", "S");

                flag = 0;
                Console.WriteLine(r_count_DI);
                for (int k = 1; k <= r_count_DI; k++)
                {
                    if (xlRange_DI.Cells[k, 1].Value2 != null && RM.ContainsKey(xlRange_DI.Cells[k, 1].Value2.ToString()))
                    {
                        flag = 0;
                        if (xlRange_DI.Cells[k, 1].Value2.ToString() == "S_ON")
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#" + xlRange_DI.Cells[k, 2].Value2.ToString().Split('_')[0]] = xlRange_DI.Cells[k, 3].Value2.ToString();
                            flag = 1;
                        }
                        if (xlRange_DI.Cells[k, 1].Value2.ToString() == "SS_OFF")
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#" + xlRange_DI.Cells[k, 2].Value2.ToString().Split('_')[0].Substring(1)] = xlRange_DI.Cells[k, 3].Value2.ToString();
                            flag = 1;
                        }
                        if (xlRange_DI.Cells[k, 1].Value2.ToString() == "LDSS_1" || xlRange_DI.Cells[k, 1].Value2.ToString() == "LDSS_2")
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#" + xlRange_DI.Cells[k, 2].Value2.ToString().Split('_')[0].Substring(3)] = xlRange_DI.Cells[k, 3].Value2.ToString();
                            flag = 1;
                        }
                        if (xlRange_DI.Cells[k, 1].Value2.ToString() == "S_OFF_V" || xlRange_DI.Cells[k, 1].Value2.ToString() == "S_OFF_B")
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#" + xlRange_DI.Cells[k, 2].Value2.ToString().Split('_')[0]] = xlRange_DI.Cells[k, 3].Value2.ToString();
                            flag = 1;
                        }
                        if (xlRange_DI.Cells[k, 1].Value2.ToString() == "UPS" || xlRange_DI.Cells[k, 1].Value2.ToString() == "MPS" || xlRange_DI.Cells[k, 1].Value2.ToString() == "UPS1" || xlRange_DI.Cells[k, 1].Value2.ToString() == "UPS2" || xlRange_DI.Cells[k, 1].Value2.ToString() == "MPS1" || xlRange_DI.Cells[k, 1].Value2.ToString() == "MPS2" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PSC")
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#AL" + Sector_Entry.Text] = xlRange_DI.Cells[k, 3].Value2.ToString();
                            flag = 1;
                        }
                        if (xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_MPLA" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_MPLDA" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_ISNA" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_ASNA" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_RBLA" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_RUBLA" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_MBLA" || xlRange_DI.Cells[k, 1].Value2.ToString() == "PREP_MBULA")
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#" + RM[xlRange_DI.Cells[k, 1].Value2.ToString()] + Sector_Entry.Text] = xlRange_DI.Cells[k, 3].Value2.ToString();
                            flag = 1;
                        }
                        if (xlRange_DI.Cells[k, 1].Value2.ToString() == "RL_S")
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#" + xlRange_DI.Cells[k, 2].Value2.ToString().Split('_')[0] + "_" + xlRange_DI.Cells[k, 2].Value2.ToString().Split('_')[1]] = xlRange_DI.Cells[k, 3].Value2.ToString();
                            flag = 1;
                        }
                        if (flag == 0)
                        {
                            DI[xlRange_DI.Cells[k, 1].Value2.ToString() + "#" + RM[xlRange_DI.Cells[k, 1].Value2.ToString()] + xlRange_DI.Cells[k, 2].Value2.ToString().Substring(xlRange_DI.Cells[k, 1].Value2.ToString().Length)] = xlRange_DI.Cells[k, 3].Value2.ToString();
                        }

                        //DI.Add(xlRange_DI.Cells[k,1].Value2.ToString()+"#"+ xlRange_DI.Cells[k, 1].Value2.ToString().Substring(xlRange_DI.Cells[k, 1].Value2.ToString().Length), xlRange_DI.Cells[k, 3].Value2.ToString());
                    }
                    else if (xlRange_DI.Cells[k, 1].Value2 != null)
                    {
                        //Console.WriteLine(xlRange_DI.Cells[k, 1].Value2.ToString());
                    }
                }
                string bit = "";
                for (int i = 0; i < mtk_f.Count; i++)
                {
                    if (mtk_f[i].Contains(";") && DI.ContainsKey(mtk_f[i].Split(';')[1]))
                    {
                        bit = DI[mtk_f[i].Split(';')[1]];
                        mtk_f[i] = bit + ";" + mtk_f[i].Split(';')[1];
                        Console.WriteLine(mtk_f[i]);

                    }

                }
                using (StreamWriter mw = MTK_File.AppendText())
                {
                    foreach (string l in mtk_f)
                    {
                        mw.WriteLine(l);
                    }
                }
                UpdateRichTextBox("MTK Completed\n");

                //Marshal.ReleaseComObject(xlApp);




            }
        }

        static void WriteToFile(string line,FileInfo PAR_File)
        {
            // Write a line of text to the file with a format based on the thread ID
            lock (_lock)
            {

                using (StreamWriter writer = PAR_File.AppendText())
                {
                    writer.WriteLine(line);
                }
            }
        }
        private void Browse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            DialogResult result_folder = folder.ShowDialog();
            if(result_folder == DialogResult.OK)
            {
                string folder_name = folder.SelectedPath;
                try
                {
                    Output_Folder.Text = folder_name.ToString();
                }
                catch(IOException ex)
                { 
                    MessageBox.Show(ex.ToString(),"Error",MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        


        private void Search_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            DialogResult result = file.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file_name = file.FileName;
                try
                {
                    File_Entry.Text = file_name.ToString();
                }
                catch(IOException ex)
                {
                    MessageBox.Show(ex.ToString(),"Error",MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if(Sector_Entry.Text != null)
            {
                flowLayoutPanel1.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
        }
        private void MTK_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Deselect_All.Enabled = true;
            }
        }

        /*private void PAR_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true)
            {
                Output.Enabled = true;
            }
        }*/

        private void Track_Circuit_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Traffic_Direction_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void File_Entry_TextChanged(object sender, EventArgs e)
        {

        }

        private void Output_Folder_TextChanged(object sender, EventArgs e)
        {

        }

        private void MBL_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void Subroute_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void Route_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void Point_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void Cycle_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void Three_Aspect_Signal_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void Shunt_Signal_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }
        private void Overlap_CheckedChanged(object sender, EventArgs e)
        {
            if (MTK.Checked == true || Track_Circuit.Checked == true || Traffic_Direction.Checked == true || MBL.Checked == true || Subroute.Checked == true || Route.Checked == true || Point.Checked == true || Three_Aspect_Signal.Checked == true || Shunt_Signal.Checked == true || Cycle.Checked == true || Overlap.Checked == true || MTK.Checked == true)
            {
                Output.Enabled = true;
                Select_All.Enabled = true;
                Deselect_All.Enabled = true;
            }
            else
            {
                Output.Enabled = false;
            }
        }

        private void flowLayoutPanel1_Paint_1(object sender, PaintEventArgs e)
        {

        }
        private void Select_All_Button(object sender, EventArgs e)
        {
            Track_Circuit.Checked = true;
            Traffic_Direction.Checked = true;
            MBL.Checked = true;
            Subroute.Checked = true;
            Route.Checked = true;
            Cycle.Checked = true;
            Overlap.Checked = true;
            MTK.Checked = true;
            Shunt_Signal.Checked = true;
            Three_Aspect_Signal.Checked = true;
            Point.Checked = true;
        }
        private void Deselect_All_Button(object sender, EventArgs e)
        {
            Track_Circuit.Checked = false;
            Traffic_Direction.Checked = false;
            MBL.Checked = false;
            Subroute.Checked = false;
            Route.Checked = false;
            Cycle.Checked = false;
            Overlap.Checked = false;
            MTK.Checked = false;
            Shunt_Signal.Checked = false;
            Three_Aspect_Signal.Checked = false;
            Point.Checked = false;
        }

        private void Search_MTK_Button(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            DialogResult result = file.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file_name = file.FileName;
                try
                {
                    MTK_Entry.Text = file_name.ToString();
                }
                catch (IOException ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }
    }
}
