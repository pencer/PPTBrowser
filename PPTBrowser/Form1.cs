using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.IO;

namespace PPTBrowser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] filenames = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string filename in filenames)
            {
                label1.Text = filename;
                listBox1.Items.Add(filename);
                string[] item = {filename, "?", "0"};
                listView1.Items.Add(new ListViewItem(item));
            }
            //ExtractSlidesFromPPT(false);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.BaseFolderPath;

            listView1.FullRowSelect = true;
            listView1.View = System.Windows.Forms.View.Details;
            listView1.Columns.Add("File", 400);
            listView1.Columns.Add("Slides");
            listView1.Columns.Add("Status");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExtractSlidesFromPPT(true);
        }

        private void ExtractSlidesFromPPT(bool visible)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = null;
            Microsoft.Office.Interop.PowerPoint.Presentation ppt = null;
            try {
                // PPTのインスタンス作成
                app = new Microsoft.Office.Interop.PowerPoint.Application();
                
                // 表示する
                app.Visible = (visible) ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;

                foreach (string item in listBox1.Items)
                {
                    String pptfilename = item;// label1.Text;
                    label2.Text = "Opening " + item + "...";
                    // オープン
                    ppt = app.Presentations.Open(pptfilename,
                        Microsoft.Office.Core.MsoTriState.msoTrue,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoFalse);

                    // https://msdn.microsoft.com/JA-JP/library/office/ff746030.aspx
                    String basefilename = Properties.Settings.Default.BaseFolderPath + "\\" + pptfilename.Replace("\\", "_").Replace(":", "_");
                    label2.Text = "Slides.count=" + ppt.Slides.Count + ", " + basefilename;
                    for (int i = 1; i <= ppt.Slides.Count; i++)
                    {
                        // スライド番号は１から始まるのに注意
                        ppt.Slides.Range(i).Export(basefilename + i.ToString("_%03d") + ".png", "png", 640, 480);
                    }
                    ppt.Close();
                    //app.Presentations[1].Close();
                }

                // http://stackoverflow.com/questions/981547/powerpoint-launched-via-c-sharp-does-not-quit
                GC.Collect();
                GC.WaitForPendingFinalizers();
//                ppt.Close();
                Marshal.ReleaseComObject(ppt);
                app.Quit();
                Marshal.ReleaseComObject(app);
                
//                app.Presentations[0].Close();
            }
            catch { }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Folder Selection Dialog
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.SelectedPath = textBox1.Text;
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = dlg.SelectedPath;
                Properties.Settings.Default.BaseFolderPath = dlg.SelectedPath;
                Properties.Settings.Default.Save();
            }

        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] filename = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            textBox1.Text = filename[0];
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }

        }

        void DirSearch(string sDir)
        {
            string strext = "*.pptx";
            //listBox1.Items.Add(sDir);
            try
            {
                foreach (string f in Directory.GetFiles(sDir, strext))
                {
                    listBox1.Items.Add(f);
                }
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    DirSearch(d);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DirSearch(label3.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Folder Selection Dialog
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                label3.Text = dlg.SelectedPath;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExtractSlidesFromPPT(true);
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
