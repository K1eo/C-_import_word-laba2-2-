using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Access = Microsoft.Office.Interop;

namespace Import
{
    public partial class Form1 : Form
    {
        private string name ;
        public Form1()
        {
            InitializeComponent();
        }

        public void RunMacros(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oApp, oRunArgs);
             
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object oMissing = System.Reflection.Missing.Value;
                    Word.Application oWord = new Word.Application();
                    oWord.Visible = true;
                    Word.Documents oDocs = oWord.Documents;
                    object oFile = "f:\\Andre.doc";
  
           
            /*Word.Document oDoc = oDocs.Open("f:\\Andre.doc", ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            */
            var oDoc = oWord.Documents.Open("f:\\Andre.docx");
            oWord.Run("Sort");

                    oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDoc);
                    oDoc = null;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocs);
                    oDocs = null;
                    oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord);
                    oWord = null;
            GC.Collect();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            InitializeComponent();

            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Filter = "Файлы docx|*.docx";
            object oFile;
            if (OPF.ShowDialog() == DialogResult.OK)
            {
                oFile = (OPF.FileName);
            }
            else { oFile = "f:\\Andre.docx"; }
            object oMissing = System.Reflection.Missing.Value;
            Word.Application oWord = new Word.Application();
            oWord.Visible = true;
            Word.Documents oDocs = oWord.Documents;
            var oDoc = oWord.Documents.Open(oFile);
            oWord.Run("Sort");
           
            oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDoc);
            oDoc = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocs);
            oDocs = null;
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWord);
            oWord = null;
            GC.Collect();
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Form1 f = new Form1();
            f.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
