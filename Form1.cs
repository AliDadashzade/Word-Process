using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;  //Enter this library for word process
using System.Reflection;

namespace Word_Process
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object omissing = System.Reflection.Missing.Value;
            object dokumansonu = "\\endofdoc";
            word.Application olustur=new word.Application();
            olustur.Visible = true;
            word.Document icerik=new word.Document();
            icerik=olustur.Documents.Add(ref omissing);

            word.Paragraph paragraph1;
            paragraph1 = icerik.Content.Paragraphs.Add(ref omissing);
            paragraph1.Range.Text = richTextBox1.Text;
            paragraph1.Range.Font.Bold = 2;
            paragraph1.SpaceAfter = 15; //it determines how much space we have to leave
            paragraph1.Range.InsertParagraphAfter(); //

            word.Paragraph paragraph2;
            paragraph2 = icerik.Content.Paragraphs.Add(ref omissing);
            paragraph2.Range.Text = "Bye";
            paragraph2.Range.Font.Bold = 3;
            paragraph2.SpaceAfter = 20;


            word.Table tabloolustur;
            word.Range wrdrng = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            tabloolustur = icerik.Tables.Add(wrdrng, 3, 5, ref omissing, ref omissing);
            tabloolustur.Range.ParagraphFormat.SpaceAfter = 15;
            int r, c;
            string strtext;
            for (r = 1; r <= 3; r++)
            {
                for (c = 1; c <= 5; c++)
                {
                    strtext = "Satir" + r + "Sutun" + c;
                    tabloolustur.Cell(r,c).Range.Text = strtext;
                }
            }
        }
    }
}
