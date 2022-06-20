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
using System.Text.RegularExpressions;
using DGVPrinterHelper;
using System.Runtime.InteropServices;
using System.Net;

namespace WindowsApplication1
{
    public partial class Form2 : Form
    {
        public Form f2;
        public string name;
        int yu,uy;
        public Form2()
        {
            InitializeComponent();
            for (int i = 0; i < 6; i++)
            {
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].HeaderCell.Style.Font = new System.Drawing.Font("Arial", 16); ;
            }
            for (int i = 0; i < 3; i++)
            {
                dataGridView2.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[i].HeaderCell.Style.Font = new System.Drawing.Font("Arial", 16); ;
            }
           
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (button1.Enabled || button2.Enabled)
            {
                DialogResult result3 = MessageBox.Show("åá ÊÑíÏ ÍÝÙ ÇáÊÛíÑÇÊ ÇáÊí ÞãÊ ÈåÇ?",
             "ÊÍÐíÑ",
             MessageBoxButtons.YesNoCancel,
             MessageBoxIcon.Question,
             MessageBoxDefaultButton.Button3);
                if (result3 == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
                else if (result3 == DialogResult.No)
                {
                    f2.Show();
                }
                else if (result3 == DialogResult.Yes)
                {
                    button1_Click(null, null);
                    if(yu==0)
                    {
                        e.Cancel = true;
                        return;
                    }
                    button2_Click(null, null);
                    if (uy == 0)
                    {
                        e.Cancel = true;
                        return;
                    }
                    f2.Show();
                }
            }
            else
                f2.Show();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            yu = 1;
            try
            {
                Process[] processlist = Process.GetProcesses();

                foreach (Process theprocess in processlist)
                {
                    if (theprocess.ProcessName == "EXCEL")
                    {
                        MessageBox.Show("ÇÛáÞ ßá ãáÝÇÊ ÇßÓá ÇáãÝÊæÍÉ");
                        yu = 0; return; 
                    }
                }
                float res = 0, sum = 0, f; int i = 0, j = 0;
                //----------------------------------------------
                //validtion
                for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    j = i + 1;
                    if (string.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[0].Value))
                    { MessageBox.Show("ÇßÊÈ ÇáÈíÇä Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString()); yu = 0; return; }
                    if (dataGridView1.Rows[i].Cells[0].Value.ToString().Trim().Length < 2)
                    { MessageBox.Show("íÌÈ Ãä íßæä ÚÏÏ ÃÍÑÝ ÇáÈíÇä Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString() + " ÃßËÑ ãä ÍÑÝíä"); yu = 0; return; }
                    if (string.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[1].Value))
                    { MessageBox.Show("ÇßÊÈ ÇáÚÏÏ Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString());  yu = 0;return; }
                    if (!float.TryParse(dataGridView1.Rows[i].Cells[1].Value.ToString(), out f))
                    { MessageBox.Show("ÇßÊÈ ÇáÚÏÏ Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString() + " ÈÇáÔßá ÇáÕÍíÍ");yu = 0; return;  }
                    if (string.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[2].Value))
                    { MessageBox.Show("ÇßÊÈ ÓÚÑ ÇáæÍÏÉ Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString()); yu = 0;return;  }
                    if (!float.TryParse(dataGridView1.Rows[i].Cells[2].Value.ToString(), out f))
                    { MessageBox.Show("ÇßÊÈ ÓÚÑ ÇáæÍÏÉ Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString() + " ÈÇáÔßá ÇáÕÍíÍ"); yu = 0;return;  }
                }
                //-----------------------------------------------
                string temp = Directory.GetCurrentDirectory() + "\\names\\" + name +".xlsx";
                if (!File.Exists(temp)) { MessageBox.Show("åÐÇ ÇáãáÝ ÛíÑ ãæÌæÏ");yu = 0; return;  }
                ApplicationClass app;
                app = new ApplicationClass();
                //-----------------------------------------------
                Workbook workBook1 = app.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.Sheets[1];
                //-----------------------------------------------
                int x = int.Parse(((Range)workSheet1.Cells[2, 7]).Value2.ToString());
                for (i = x+2; i > 1; i--)
                {
                    ((Range)workSheet1.Cells[i, 1]).EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                }
                ((Range)workSheet1.Cells[2, 7]).Value2 = dataGridView1.Rows.Count - 1;
                for ( i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    ((Range)workSheet1.Cells[i + 2, 1]).Value2=dataGridView1.Rows[i].Cells[0].Value ;
                    ((Range)workSheet1.Cells[i + 2, 2]).Value2=dataGridView1.Rows[i].Cells[1].Value  ;
                    ((Range)workSheet1.Cells[i + 2, 3]).Value2=dataGridView1.Rows[i].Cells[2].Value  ;
                    res = float.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString()) * float.Parse(dataGridView1.Rows[i].Cells[1].Value.ToString());
                    dataGridView1.Rows[i].Cells[3].Value = res;
                    sum = sum + res;
                    ((Range)workSheet1.Cells[i + 2, 4]).Value2=dataGridView1.Rows[i].Cells[3].Value  ;
                    ((Range)workSheet1.Cells[i + 2, 5]).Value2=dataGridView1.Rows[i].Cells[4].Value  ;
                    ((Range)workSheet1.Cells[i + 2, 6]).Value2=dataGridView1.Rows[i].Cells[5].Value  ;
                }
                //----------------------------------------------
                ((Range)workSheet1.Cells[1, 8]).Value2 = sum;
                textBox7.Text = sum.ToString();
                res = sum - float.Parse(textBox1.Text);
                textBox8.Text = res.ToString();
                app.DisplayAlerts = false;
                workBook1.Close(true, temp, false);
                app.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app);
                kill_excel();
                MessageBox.Show("Êã ÇáÍÝÙ ÈäÌÇÍ Ýí ÌÏæá ÇáÏíæä");
                button1.Enabled = false;
                textBox5.Text = "ÈáÛ ÞíãÉ ÇáÏíä ÇáãÓÊÍÞ Úáíßã " + res.ToString() + " Ôíßá ÇáÑÌÇÁ ÇáÏÝÚ  ááãÑÇÌÚÉ ÇÈæÓÇãÑ";
                label10.Text = textBox5.Text.Length.ToString();
            }
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

        private void button2_Click(object sender, EventArgs e)
        {
            uy = 1;
            try
            {
                Process[] processlist = Process.GetProcesses();

                foreach (Process theprocess in processlist)
                {
                    if (theprocess.ProcessName == "EXCEL")
                    {
                        MessageBox.Show("ÇÛáÞ ßá ãáÝÇÊ ÇßÓá ÇáãÝÊæÍÉ");
                        uy = 0;return; 
                    }
                }
                //----------------------------------------------
                //validtion
                float res = 0, sum = 0,f; int i = 0,j=0;
                for (i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    j = i + 1;
                   
                    if (string.IsNullOrEmpty((string)dataGridView2.Rows[i].Cells[0].Value))
                    { MessageBox.Show("ÇßÊÈ ÇáÏÝÚÉ Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString());  uy = 0;return; }
                    if (!float.TryParse(dataGridView2.Rows[i].Cells[0].Value.ToString(), out f))
                    { MessageBox.Show("ÇßÊÈ ÇáÏÝÚÉ Ýí ÇáÓØÑ ÑÞã" + " : " + j.ToString() + " ÈÇáÔßá ÇáÕÍíÍ");uy = 0; return;  }
                }
                string temp = Directory.GetCurrentDirectory() + "\\names\\" + name + ".xlsx";
                if (!File.Exists(temp)) { MessageBox.Show("åÐÇ ÇáãáÝ ÛíÑ ãæÌæÏ");uy = 0; return;  }
                ApplicationClass app;
                app = new ApplicationClass();
                //-----------------------------------------------
                Workbook workBook1 = app.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.Sheets[2];
                //----------------------------------------------
                int x = int.Parse(((Range)workSheet1.Cells[2, 4]).Value2.ToString());
                for (i = x + 2; i > 1; i--)
                {
                    ((Range)workSheet1.Cells[i, 1]).EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                }
                ((Range)workSheet1.Cells[2, 4]).Value2 = dataGridView2.Rows.Count - 1;
                for (i = 0; i < dataGridView2.Rows.Count - 1; i++)
                {
                    ((Range)workSheet1.Cells[i + 2, 1]).Value2 = dataGridView2.Rows[i].Cells[0].Value;
                    ((Range)workSheet1.Cells[i + 2, 2]).Value2 = dataGridView2.Rows[i].Cells[1].Value;
                    ((Range)workSheet1.Cells[i + 2, 3]).Value2 = dataGridView2.Rows[i].Cells[2].Value;
                    sum = sum + float.Parse(dataGridView2.Rows[i].Cells[0].Value.ToString());
                }
                //----------------------------------------------
                ((Range)workSheet1.Cells[1, 5]).Value2 = sum;
                res = float.Parse(textBox7.Text) -sum;
                textBox8.Text = res.ToString();
                textBox1.Text = sum.ToString();
                app.DisplayAlerts = false;
                workBook1.Close(true, temp, false);
                app.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app);
                kill_excel();
                MessageBox.Show("Êã ÇáÍÝÙ ÈäÌÇÍ Ýí ÌÏæá ÇáÏÝÚÇÊ");
                button2.Enabled = false;
                textBox5.Text = "ÈáÛ ÞíãÉ ÇáÏíä ÇáãÓÊÍÞ Úáíßã " + res.ToString() + " Ôíßá ÇáÑÌÇÁ ÇáÏÝÚ  ááãÑÇÌÚÉ ÇÈæÓÇãÑ";
                label10.Text = textBox5.Text.Length.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
                kill_excel();
            }
            }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                dataGridView1.Rows[i].HeaderCell.Value = String.Format("{0}", i + 1);
            button1.Enabled = true;
            string s;
            s = DateTime.Now.Day.ToString()+"/"+DateTime.Now.Month.ToString()+"/"+DateTime.Now.Year.ToString();
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[5].Value =s;

        }

        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                dataGridView2.Rows[i].HeaderCell.Value = String.Format("{0}", i + 1);
            button2.Enabled = true;
            string s;
            s = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
            dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[2].Value = s;

        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                dataGridView1.Rows[i].HeaderCell.Value = String.Format("{0}", i + 1);
            button1.Enabled = true;
        }

        private void dataGridView2_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                dataGridView2.Rows[i].HeaderCell.Value = String.Format("{0}", i + 1);
            button2.Enabled = true;
        }
        private void print_all(int x)
        {
            try
            {
                int o = 0;
                float sum = 0, res = 0, f = 0, s = 0, s2 = 0;
                int i = 0, j2 = 0;
                if (textBox2.Text.Trim().Length < 1) { MessageBox.Show("ÇßÊÈ ÑÞã ÇáÓØÑ Ýí ÌÏæá ÊÓÌíá ÇáÏíæä"); return; }
                if (textBox3.Text.Trim().Length < 1) { MessageBox.Show("ÇßÊÈ ÑÞã ÇáÓØÑ Ýí ÌÏæá ÇáÏÝÚÇÊ"); return; }
                if (!int.TryParse(textBox2.Text.Trim(), out o)) { MessageBox.Show("ÇßÊÈ ÑÞã ÇáÓØÑ Ýí ÌÏæá ÊÓÌíá ÇáÏíæä ÈÇáÔßá ÇáÕÍíÍ"); return; }
                if (!int.TryParse(textBox3.Text.Trim(), out o)) { MessageBox.Show("ÇßÊÈ ÑÞã ÇáÓØÑ Ýí ÌÏæá ÇáÏÝÚÇÊ ÈÇáÔßá ÇáÕÍíÍ"); return; }
                if (int.Parse(textBox2.Text.Trim()) >= dataGridView1.Rows.Count) { MessageBox.Show("ÑÞã ÇáÓØÑ ÛíÑ ãæÌæÏ Ýí ÌÏæá ÊÓÌíá ÇáÏíæä"); return; }
                if (int.Parse(textBox3.Text.Trim()) >= dataGridView2.Rows.Count) { MessageBox.Show("ÑÞã ÇáÓØÑ ÛíÑ ãæÌæÏ Ýí ÌÏæá ÇáÏÝÚÇÊ"); return; }
                int c1 = int.Parse(textBox2.Text.Trim()); int c2 = int.Parse(textBox3.Text.Trim());
                dataGridView3.Rows.Clear();
                dataGridView3.Rows.Add();
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[0].Value = "ÇáÈíÇä";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[1].Value = "ÇáÚÏÏ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Value = "ÓÚÑ ÇáæÍÏÉ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Value = "ÇáÓÚÑ Çáßáí";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[4].Value = "ãáÇÍÙÇÊ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[5].Value = "ÇáÊÇÑíÎ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[0].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[1].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[4].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[5].Style.BackColor = Color.LightGray;
                if (c1 > 1)
                {
                    for (i = 0; i < c1 - 1; i++)
                    {

                        if (!string.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[2].Value) && float.TryParse(dataGridView1.Rows[i].Cells[2].Value.ToString(), out f))
                            if (!string.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[1].Value) && float.TryParse(dataGridView1.Rows[i].Cells[1].Value.ToString(), out f))
                            {
                                res = float.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString()) * float.Parse(dataGridView1.Rows[i].Cells[1].Value.ToString());
                                s = s + res;
                            }

                    }
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Value = "ÇáÓÇÈÞ";
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGray;
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Value = s;
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Style.BackColor = Color.LightGray;
                }


                sum = 0; res = 0;
                int u = 0;
                if (c1 > 1) u = 2; else u = 1;
                for (i = c1 - 1, j2 = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[j2 + u].Cells[0].Value = dataGridView1.Rows[i].Cells[0].Value;
                    dataGridView3.Rows[j2 + u].Cells[1].Value = dataGridView1.Rows[i].Cells[1].Value;
                    dataGridView3.Rows[j2 + u].Cells[2].Value = dataGridView1.Rows[i].Cells[2].Value;
                    dataGridView3.Rows[j2 + u].Cells[3].Value = dataGridView1.Rows[i].Cells[3].Value;
                    dataGridView3.Rows[j2 + u].Cells[4].Value = dataGridView1.Rows[i].Cells[4].Value;
                    dataGridView3.Rows[j2 + u].Cells[5].Value = dataGridView1.Rows[i].Cells[5].Value;
                    if (!string.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[2].Value) && float.TryParse(dataGridView1.Rows[i].Cells[2].Value.ToString(), out f))
                        if (!string.IsNullOrEmpty((string)dataGridView1.Rows[i].Cells[1].Value) && float.TryParse(dataGridView1.Rows[i].Cells[1].Value.ToString(), out f))
                        {
                            res = float.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString()) * float.Parse(dataGridView1.Rows[i].Cells[1].Value.ToString());
                            sum = sum + res;
                        }
                    j2++;
                }
                dataGridView3.Rows.Add();
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Value = "ÇáãÌãæÚ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Value = sum + s;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Style.BackColor = Color.LightGray;

                dataGridView3.Rows.Add();
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[0].Value = "ÇáÏÝÚÇÊ : ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Value = "ÇáÏÝÚÉ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[1].Value = "ãáÇÍÙÇÊ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Value = "ÇáÊÇÑíÎ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[1].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[0].Style.BackColor = Color.LightGray;
                if (c2 > 1)
                {
                    for (i = 0; i < c2 - 1; i++)
                    {

                        if (!string.IsNullOrEmpty((string)dataGridView2.Rows[i].Cells[0].Value) && float.TryParse(dataGridView2.Rows[i].Cells[0].Value.ToString(), out f))
                        {
                            res = float.Parse(dataGridView2.Rows[i].Cells[0].Value.ToString());
                            s2 = s2 + res;
                        }

                    }
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Value = "ÇáÓÇÈÞ";
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGray;
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Value = s2;
                    dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Style.BackColor = Color.LightGray;
                }

                float sum1 = 0, res1 = 0;
                int j = dataGridView3.Rows.Count - 1;
                for (i = c2 - 1; i < dataGridView2.Rows.Count - 1; i++)
                {
                    dataGridView3.Rows.Add();
                    dataGridView3.Rows[j].Cells[1].Value = dataGridView2.Rows[i].Cells[1].Value;
                    dataGridView3.Rows[j].Cells[2].Value = dataGridView2.Rows[i].Cells[2].Value;
                    dataGridView3.Rows[j].Cells[3].Value = dataGridView2.Rows[i].Cells[0].Value;
                    if (!string.IsNullOrEmpty((string)dataGridView2.Rows[i].Cells[0].Value) && float.TryParse(dataGridView2.Rows[i].Cells[0].Value.ToString(), out f))
                    {
                        res1 = float.Parse(dataGridView2.Rows[i].Cells[0].Value.ToString());
                        sum1 = sum1 + res1;
                    }
                    j++;
                }
                dataGridView3.Rows.Add();
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Value = "ÇáãÌãæÚ";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Value = sum1 + s2;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Style.BackColor = Color.LightGray;

                dataGridView3.Rows.Add();
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Value = "ÇáÕÇÝí";
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Value = (sum+s) - (sum1+s2);
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[2].Style.BackColor = Color.LightGray;
                dataGridView3.Rows[dataGridView3.Rows.Count - 2].Cells[3].Style.BackColor = Color.LightGray;
                //----------------------------------------------------------------------
                DGVPrinter printer = new DGVPrinter();
                printer.Title = "ÇáÇäæÇÑ ááÊÌåíÒÇÊ ÇáßåÑÈÇÆíÉ æÇáÕÍíÉ";
                printer.SubTitle = this.Text;
                printer.SubTitleFormatFlags = StringFormatFlags.LineLimit |
                                              StringFormatFlags.NoClip;
                printer.PageNumbers = true;
                printer.PageNumberInHeader = false;
                printer.PorportionalColumns = true;
                printer.HeaderCellAlignment = StringAlignment.Near;
                printer.Footer = "ÇáÇäæÇÑ ááÊÌåíÒÇÊ ÇáßåÑÈÇÆíÉ æÇáÕÍíÉ - ßÝÑÇáÏíß 0599887446";
                printer.FooterSpacing = 15;
                if (x == 1) printer.PrintDataGridView(dataGridView3);
                if (x == 0) printer.PrintPreviewDataGridView(dataGridView3);
            }
            catch (Exception ex)
            {
                MessageBox.Show("íæÌÏ áÏíß ãÔßáÉ");
            }
        }
    
        private void button3_Click(object sender, EventArgs e)
        {
            print_all(1);
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            print_all(0);
            
        }  

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            button2.Enabled = true;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            button1.Enabled = true;
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
           // MessageBox.Show("");
        }
        
        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPageIndex == 2)
            {
                button3.Visible = false; button4.Visible = false; label1.Visible = false; 
                textBox2.Visible = false; textBox3.Visible = false; textBox8.Visible = false;
                label2.Visible=false; label3.Visible=false; label4.Visible=false;
            }
            else
            {
                button3.Visible = true; button4.Visible = true; label1.Visible = true; 
                textBox2.Visible = true; textBox3.Visible = true; textBox8.Visible = true;
                label2.Visible=true; label3.Visible=true; label4.Visible=true;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string url = textBox4.Text;
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse res = (HttpWebResponse)req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            MessageBox.Show(responseString);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string url = textBox6.Text + textBox9.Text + textBox10.Text + textBox5.Text + textBox11.Text;
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse res = (HttpWebResponse)req.GetResponse();
            string responseString = new StreamReader(res.GetResponseStream()).ReadToEnd();
            MessageBox.Show(responseString);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
           label10.Text= textBox5.Text.Length.ToString();
        }

   
    }
}