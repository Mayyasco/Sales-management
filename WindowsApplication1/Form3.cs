using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace WindowsApplication1
{
    public partial class Form3 : Form
    {
        public Form f2;
        public int ri;
        public string s;
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {

                    f2.Show();
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
                if (textBox2.Text.Trim() != s)
                {
                    string tem1 = Directory.GetCurrentDirectory() + "\\names\\" + s + ".xlsx";
                    string tem2 = Directory.GetCurrentDirectory() + "\\names\\" + textBox2.Text.Trim() + ".xlsx";
                    if (File.Exists(tem2)) { MessageBox.Show("ÇáÇÓã ãæÌæÏ"); return; }
                    File.Move(tem1, tem2);

                }
                ApplicationClass app;
                app = new ApplicationClass();
                //-----------------------------------------------
                string temp = Directory.GetCurrentDirectory() + "\\names\\names.xlsx";
                 if(!File.Exists(temp)) { MessageBox.Show("ãáÝ ÇáÇÓãÇÁ ÛíÑ ãæÌæÏ"); return; }
                Workbook workBook1 = app.Workbooks.Open(temp, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet workSheet1 = (Worksheet)workBook1.ActiveSheet;
                //----------------------------------------------
                ((Range)workSheet1.Cells[ri, 1]).Value2 = textBox2.Text.Trim();//ÇáÇÓã
                ((Range)workSheet1.Cells[ri, 2]).Value2 = textBox3.Text;//ÇáÈáÏ
                ((Range)workSheet1.Cells[ri, 3]).Value2 = textBox4.Text;//ÚäæÇä ÇáæÑÔÉ
                ((Range)workSheet1.Cells[ri, 4]).Value2 = textBox5.Text;//ÑÞã ÇáÌæÇá
                //----------------------------------------------
                textBox2.Focus();
                workBook1.Save();
                workBook1.Close(true, temp, false);
                app.Quit();
                GC.Collect();
                Marshal.FinalReleaseComObject(workSheet1);
                Marshal.FinalReleaseComObject(workBook1);
                Marshal.FinalReleaseComObject(app);
                kill_excel();
                TabControl tc = (TabControl)f2.Controls[0];
                System.Windows.Forms.ListBox lb = (System.Windows.Forms.ListBox)tc.TabPages[0].Controls["listBox1"];
                lb.Items[lb.Items.IndexOf(s)] = textBox2.Text.Trim();
                MessageBox.Show("Êã ÇáÊÚÏíá ÈäÌÇÍ");
                this.Close();
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
    }
}