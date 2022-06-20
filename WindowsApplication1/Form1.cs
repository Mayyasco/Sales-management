using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Net;

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {
        float xxx1, xxx2;
        Form2 f3; Form3 f33;
        public Form1()
        {
            InitializeComponent();
            for (int i = 0; i < 4; i++)
            {
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox2.Text.Trim().Length < 5) { MessageBox.Show(" ÇßÊÈ ÇáÇÓã ÇæáÇ ¡ æáÇ íÌÈ Çä íÞá Úä 5 ÍÑæÝ "); return; }
            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    MessageBox.Show("ÇÛáÞ ßá ãáÝÇÊ ÇßÓá ÇáãÝÊæÍÉ");
                    return;
                }
            } 
            
            //-----------------------------------------------
            string tem1 = Directory.GetCurrentDirectory() + "\\names\\" + "temp_name" + ".xlsx";
            string tem2 = Directory.GetCurrentDirectory() + "\\names\\" + textBox2.Text.Trim() + ".xlsx";
            if (File.Exists(tem2)) { MessageBox.Show("ÇáÇÓã ãæÌæÏ"); return; }
            if (!File.Exists(tem1)) { MessageBox.Show("ãáÝ ÞÇáÈ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
            ApplicationClass app;
            app = new ApplicationClass();
            Workbook workBook2 = app.Workbooks.Open(tem1, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet workSheet2 = (Worksheet)workBook2.ActiveSheet;
            //-----------------------------------------------
            string temp = Directory.GetCurrentDirectory() + "\\names\\names.xlsx";
            if (!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
            Workbook workBook1 = app.Workbooks.Open(temp,0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
            //----------------------------------------------
            int x =int.Parse( ((Range)workSheet1.Cells[2, 5]).Value2.ToString());
            x = x + 2;
            ((Range)workSheet1.Cells[x, 1]).Value2 = textBox2.Text.Trim();//ÇáÇÓã
            ((Range)workSheet1.Cells[x, 2]).Value2 =textBox3.Text;//ÇáÈáÏ
            ((Range)workSheet1.Cells[x, 3]).Value2 =textBox4.Text;//ÚäæÇä ÇáæÑÔÉ
            ((Range)workSheet1.Cells[x, 4]).Value2 = textBox5.Text;//ÑÞã ÇáÌæÇá
            ((Range)workSheet1.Cells[2, 5]).Value2 = x-1;
            ((Range)workSheet2.Cells[1, 1]).Value2 = "ÇáÈíÇä";
            //----------------------------------------------
            textBox2.Text = ""; textBox3.Text = ""; textBox4.Text = ""; textBox5.Text = "";
            textBox2.Focus();
            workBook1.Save(); 
            workBook1.Close(true, temp, false);
            workBook2.Close(true, tem2, false);
            app.Quit();
            GC.Collect();
            Marshal.FinalReleaseComObject(workSheet1);
            Marshal.FinalReleaseComObject(workSheet2);
            Marshal.FinalReleaseComObject(workBook1);
            Marshal.FinalReleaseComObject(workBook2);
            Marshal.FinalReleaseComObject(app);
            kill_excel();}
                catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
            }
            
        }
        private void kill_excel()
        {
            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    theprocess.Kill();
                    return;
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
             try
            {
                 if (textBox1.Text.Trim().Length < 1) { MessageBox.Show("ÇÏÎá Úáì ÇáÇÞá ÍÑÝ æÇÍÏ"); return; }
            //-----------------------------------------------
            
            string temp = Directory.GetCurrentDirectory() + "\\names\\names.xlsx";
            if (!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
            ApplicationClass app1;
            app1 = new ApplicationClass();
            Workbook workBook1 = app1.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
            //----------------------------------------------
            string str;
            int x = int.Parse(((Range)workSheet1.Cells[2, 5]).Value2.ToString());
            listBox1.Items.Clear();
            for (int i = 2; i < x+2; i++)
            {
                str = ((Range)workSheet1.Cells[i, 1]).Value2.ToString();
                if (radioButton1.Checked)
                {
                    if (str.StartsWith(textBox1.Text.Trim()))
                        listBox1.Items.Add(str);
                }
                else 
                {
                    if (str.Contains(textBox1.Text.Trim()))
                        listBox1.Items.Add(str);
                }
            }
            if(listBox1.Items.Count==0)
                MessageBox.Show("áÇ íæÌÏ ÊØÇÈÞ ãÚ ÇáÇÓã ÇáãßÊæÈ");
            workBook1.Close(false, temp, false);
            app1.Quit();
            GC.Collect();
            Marshal.FinalReleaseComObject(workSheet1);
            Marshal.FinalReleaseComObject(workBook1);
            Marshal.FinalReleaseComObject(app1);
            kill_excel();}
                 catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Process[] processlist = Process.GetProcesses();

                foreach (Process theprocess in processlist)
                {
                    if (theprocess.ProcessName == "EXCEL")
                    {
                        MessageBox.Show("ÇÛáÞ ßá ãáÝÇÊ ÇßÓá ÇáãÝÊæÍÉ");
                        return;
                    }
                }
                if (listBox1.SelectedItems.Count == 0) { MessageBox.Show("ÇÎÊÑ ÇáÇÓã ÇæáÇ"); return; }
                string temp = Directory.GetCurrentDirectory() + "\\names\\" + listBox1.SelectedItem.ToString() +".xlsx";
                if (!File.Exists(temp)) { MessageBox.Show("åÐÇ ÇáãáÝ ÛíÑ ãæÌæÏ"); return; }
                this.Hide();
                f3 = new Form2();
                f3.StartPosition = FormStartPosition.CenterScreen; 
                f3.Text += listBox1.SelectedItem.ToString();
                f3.f2 = this;
                f3.name = listBox1.SelectedItem.ToString();
                load_data(f3, temp, listBox1.SelectedItem.ToString());
                f3.button1.Enabled = false; f3.button2.Enabled = false;
                f3.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
            }
        }

        private void load_data(Form2 f,string temp,string nm)
        {
            try
            {
                ApplicationClass app;
                app = new ApplicationClass();
                //-----------------------------------------------
                Workbook workBook1 = app.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.Sheets[1];
                Worksheet workSheet2 = (Worksheet)workBook1.Sheets[2];
                //----------------------------------------------
                int x = int.Parse(((Range)workSheet1.Cells[2, 7]).Value2.ToString());
                float sum1 = 0;
                f.dataGridView1.Rows.Clear();
                if(x!=0)f.dataGridView1.Rows.Add(x);
                for(int i=0;i<x;i++)
                {
                f.dataGridView1.Rows[i].Cells[0].Value=((Range)workSheet1.Cells[i+2, 1]).Value2; 
                f.dataGridView1.Rows[i].Cells[1].Value=((Range)workSheet1.Cells[i+2, 2]).Value2;
                f.dataGridView1.Rows[i].Cells[2].Value=((Range)workSheet1.Cells[i+2, 3]).Value2;
                sum1 = sum1 + float.Parse(((Range)workSheet1.Cells[i + 2, 4]).Value2.ToString());
                f.dataGridView1.Rows[i].Cells[3].Value=((Range)workSheet1.Cells[i+2, 4]).Value2;
                f.dataGridView1.Rows[i].Cells[4].Value = ((Range)workSheet1.Cells[i+2, 5]).Value2;
                f.dataGridView1.Rows[i].Cells[5].Value = ((Range)workSheet1.Cells[i+2, 6]).Value2; 
                }
                f.textBox7.Text = sum1.ToString();
                //----------------------------------------------
                int y = int.Parse(((Range)workSheet2.Cells[2, 4]).Value2.ToString());
                float sum2 = 0;
                f.dataGridView2.Rows.Clear();
                if(y!=0)f.dataGridView2.Rows.Add(y);
                for(int j=0;j<y;j++)
                {
                f.dataGridView2.Rows[j].Cells[0].Value=((Range)workSheet2.Cells[j+2, 1]).Value2;
                sum2 = sum2 + float.Parse(((Range)workSheet2.Cells[j + 2, 1]).Value2.ToString());
                f.dataGridView2.Rows[j].Cells[1].Value = ((Range)workSheet2.Cells[j + 2, 2]).Value2;
                f.dataGridView2.Rows[j].Cells[2].Value = ((Range)workSheet2.Cells[j + 2, 3]).Value2;
                }
                f.textBox1.Text = sum2.ToString();
                float sum = sum1 - sum2;
                f.textBox8.Text = sum.ToString();
             
                //----------------------------------------------
                workBook1.Close(false, temp, false);
                app.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workSheet2);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app);
                kill_excel();
                //----------------------------------------------
                f.textBox5.Text = "ÈáÛ ÞíãÉ ÇáÏíä ÇáãÓÊÍÞ Úáíßã " + sum.ToString() + " Ôíßá ÇáÑÌÇÁ ÇáÏÝÚ  ááãÑÇÌÚÉ ÇÈæÓÇãÑ";
                f.label10.Text = f.textBox5.Text.Length.ToString();
                f.textBox9.Text = get_phone(nm);
            }
            catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
            }
        }

        private string get_phone(string nm)
        {
            try
            {
                //-----------------------------------------------
                string str="";
                string temp = Directory.GetCurrentDirectory() + "\\names\\names.xlsx";
                if (!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return str; }
                ApplicationClass app1;
                app1 = new ApplicationClass();
                Workbook workBook1 = app1.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
                //----------------------------------------------
               
                int x = int.Parse(((Range)workSheet1.Cells[2, 5]).Value2.ToString());
                for (int i = 2; i < x + 2; i++)
                {
                    if (nm == ((Range)workSheet1.Cells[i, 1]).Value2.ToString())
                    {
                        str = ((Range)workSheet1.Cells[i, 4]).Value2.ToString();
                        workBook1.Close(false, temp, false);
                        app1.Quit();
                        GC.Collect();
                        Marshal.FinalReleaseComObject(workSheet1);
                        Marshal.FinalReleaseComObject(workBook1);
                        Marshal.FinalReleaseComObject(app1);
                        kill_excel();
                        return str;
                    }
                }
                workBook1.Close(false, temp, false);
                app1.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app1);
                kill_excel();
                    return str;
            }
            catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
                return "";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {               
                //-----------------------------------------------
                
                string temp = Directory.GetCurrentDirectory() + "\\names\\names.xlsx";
                if (!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
                ApplicationClass app1;
                app1 = new ApplicationClass();
                Workbook workBook1 = app1.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
                //----------------------------------------------
                string str;
                int x = int.Parse(((Range)workSheet1.Cells[2, 5]).Value2.ToString());
                listBox1.Items.Clear();
                for (int i = 2; i < x + 2; i++)
                {
                    str = ((Range)workSheet1.Cells[i, 1]).Value2.ToString();
                    listBox1.Items.Add(str);
                }
                if (listBox1.Items.Count == 0)
                    MessageBox.Show("áÇ íæÌÏ ÃÓãÇÁ");
                workBook1.Close(false, temp, false);
                app1.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app1);
                kill_excel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Process[] processlist = Process.GetProcesses();

                foreach (Process theprocess in processlist)
                {
                    if (theprocess.ProcessName == "EXCEL")
                    {
                        MessageBox.Show("ÇÛáÞ ßá ãáÝÇÊ ÇßÓá ÇáãÝÊæÍÉ");
                        return;
                    }
                }
                if (listBox1.SelectedItems.Count == 0) { MessageBox.Show("ÇÎÊÑ ÇáÇÓã ÇæáÇ"); return; }
                DialogResult dialogResult = MessageBox.Show("åá ãÊÃßÏ ãä ÇáÍÐÝ¿", "ÊÃßíÏ ÍÐÝ", MessageBoxButtons.YesNo);
                if (dialogResult != DialogResult.Yes) return;
                //---------------------
                string temp = Directory.GetCurrentDirectory() + "\\names\\" + "names" + ".xlsx";
                if (!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
                //-----------------------------------------------
                ApplicationClass app1;
                app1 = new ApplicationClass();
                Workbook workBook1 = app1.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
                //----------------------------------------------
                string str;
                int x = int.Parse(((Range)workSheet1.Cells[2, 5]).Value2.ToString());
                for (int i = 2; i < x + 2; i++)
                {
                    str = ((Range)workSheet1.Cells[i, 1]).Value2.ToString();
                    if (str.Equals(listBox1.SelectedItems[0].ToString()))
                    {
                        ((Range)workSheet1.Cells[i, 1]).EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                        break;
                    }

                }
                ((Range)workSheet1.Cells[2, 5]).Value2 = x - 1;
                app1.DisplayAlerts = false;
                workBook1.Close(true, temp, false);
                app1.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app1);
                kill_excel();
                string temp1 = Directory.GetCurrentDirectory() + "\\names\\" + listBox1.SelectedItem.ToString() + ".xlsx";
                if (File.Exists(temp1)) File.Delete(temp1);
                listBox1.Items.Remove(listBox1.SelectedItems[0]);
                MessageBox.Show("Êã ÇáÍÐÝ ÈäÌÇÍ");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                    kill_excel();
                }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            textBox12.Text = ""; textBox13.Text = ""; textBox14.Text = "";
            try
            {
            Process[] processlist = Process.GetProcesses();
            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    MessageBox.Show("ÇÛáÞ ßá ãáÝÇÊ ÇßÓá ÇáãÝÊæÍÉ");
                    return;
                }
            }
            //--------------------------------------------------
            string temp = Directory.GetCurrentDirectory() + "\\names\\" + "names" + ".xlsx";
            if (!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
            ApplicationClass app1;
            app1 = new ApplicationClass();
            Workbook workBook1 = app1.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
            //----------------------------------------------
            string str;
            int x = int.Parse(((Range)workSheet1.Cells[2, 5]).Value2.ToString());
            for (int i = 2; i < x + 2; i++)
            {
                str = ((Range)workSheet1.Cells[i, 1]).Value2.ToString();
                    dataGridView1.Rows.Add(str);
                    find_c_p(str);
                dataGridView1.Rows[i - 2].Cells[1].Value = xxx1;
                dataGridView1.Rows[i-2].Cells[2].Value=xxx2;
                if(xxx1!=-1)
                dataGridView1.Rows[i - 2].Cells[3].Value = xxx1 - xxx2;
            }
            if (dataGridView1.Rows.Count == 0)
            MessageBox.Show("áÇ íæÌÏ ÃÓãÇÁ");
            workBook1.Close(false, temp, false);
            app1.Quit();
            GC.Collect();
            Marshal.FinalReleaseComObject(workSheet1);
            Marshal.FinalReleaseComObject(workBook1);
            Marshal.FinalReleaseComObject(app1);
           kill_excel();
           float sum1 = 0,sum2=0,sum3=0;
                for (int u=0;u<dataGridView1.Rows.Count-1 ;u++)
                {
                  sum1=sum1+  float.Parse(dataGridView1.Rows[u].Cells[1].Value.ToString());
                  sum2 = sum2 + float.Parse(dataGridView1.Rows[u].Cells[2].Value.ToString());
                  sum3 = sum3 + float.Parse(dataGridView1.Rows[u].Cells[3].Value.ToString());
                }
                textBox12.Text = sum1.ToString(); textBox13.Text = sum2.ToString(); textBox14.Text = sum3.ToString();
        }
        catch (Exception ex)
        {
            MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
            kill_excel();
        }
        }

        void  find_c_p(string str)
        {
            //-----------------------------------------------
            string temp = Directory.GetCurrentDirectory() + "\\names\\" + str + ".xlsx";
            if (!File.Exists(temp)) { MessageBox.Show(" ãáÝ" + str + " ÛíÑ ãæÌæÏ"); xxx1 = -1; xxx2 = -1; return; }
            ApplicationClass app;
            app = new ApplicationClass();
            //-----------------------------------------------
            Workbook workBook1 = app.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet workSheet1 = (Worksheet)workBook1.Sheets[1];
            Worksheet workSheet2 = (Worksheet)workBook1.Sheets[2];
            //-----------------------------------------------
             xxx1 = float.Parse(((Range)workSheet1.Cells[1, 8]).Value2.ToString());
             xxx2 = float.Parse(((Range)workSheet2.Cells[1, 5]).Value2.ToString());
            workBook1.Close(false, temp, false);
            app.Quit();
            GC.Collect();
            Marshal.FinalReleaseComObject(workSheet1);
            Marshal.FinalReleaseComObject(workSheet2);
            Marshal.FinalReleaseComObject(workBook1);
            Marshal.FinalReleaseComObject(app);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                Process[] processlist = Process.GetProcesses();

                foreach (Process theprocess in processlist)
                {
                    if (theprocess.ProcessName == "EXCEL")
                    {
                        MessageBox.Show("ÇÛáÞ ßá ãáÝÇÊ ÇßÓá ÇáãÝÊæÍÉ");
                        return;
                    }
                }
                if (listBox1.SelectedItems.Count == 0) { MessageBox.Show("ÇÎÊÑ ÇáÇÓã ÇæáÇ"); return; }
                //-----------------------------------------------
                ApplicationClass app1;
                app1 = new ApplicationClass();
                string temp = Directory.GetCurrentDirectory() + "\\names\\names.xlsx";
                if (!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
                Workbook workBook1 = app1.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
                //----------------------------------------------
                string str;
                int x = int.Parse(((Range)workSheet1.Cells[2, 5]).Value2.ToString());
                for (int i = 2; i < x + 2; i++)
                {
                    str = ((Range)workSheet1.Cells[i, 1]).Value2.ToString();
                    {
                        if (str.Equals(listBox1.SelectedItems[0].ToString()))
                        {
                            this.Hide();
                            f33 = new Form3();
                            f33.StartPosition = FormStartPosition.CenterScreen;
                            f33.Text = "äÇÝÐÉ ÇáÊÚÏíá";
                            f33.f2 = this;
                            if (((Range)workSheet1.Cells[i, 1]).Value2 != null)
                            {
                                f33.textBox2.Text = ((Range)workSheet1.Cells[i, 1]).Value2.ToString();
                                f33.s = ((Range)workSheet1.Cells[i, 1]).Value2.ToString();
                            }
                        if (((Range)workSheet1.Cells[i, 2]).Value2 != null)
                            f33.textBox3.Text = ((Range)workSheet1.Cells[i, 2]).Value2.ToString();
                        if (((Range)workSheet1.Cells[i, 3]).Value2 != null)
                            f33.textBox4.Text = ((Range)workSheet1.Cells[i, 3]).Value2.ToString();
                        if (((Range)workSheet1.Cells[i, 4]).Value2 != null)
                            f33.textBox5.Text = ((Range)workSheet1.Cells[i, 4]).Value2.ToString();
                        f33.ri = i;
                            f33.Show();
                            break;
                        }
                    }
                  
                }

                workBook1.Close(false, temp, false);
                app1.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app1);
                kill_excel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string url = textBox6.Text + textBox9.Text + textBox10.Text + textBox8.Text + textBox11.Text;
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse res = (HttpWebResponse)req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            MessageBox.Show(responseString);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string url = textBox7.Text;
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse res = (HttpWebResponse)req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            MessageBox.Show(responseString);
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            label10.Text = textBox8.Text.Length.ToString();
        }

      
      
    }
}