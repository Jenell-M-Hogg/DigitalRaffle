using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LinqToExcel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            
            openFileDialog1.Title = "Browse CSV Files";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.Multiselect = true;

            openFileDialog1.DefaultExt = "csv";
            openFileDialog1.Filter = "CSV (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                String[] tri = openFileDialog1.FileNames;
                String fileBox = "";
                for (int i = 0; i< tri.Length; i++)
                {
                    if (i != 0)
                    {
                        fileBox += ",";

                    }
                    fileBox += tri[i];

                }

                textBox1.Text = fileBox;
                label1.Text = (tri.Length).ToString() + " files selected";
            }
        }
      
        private void button2_Click(object sender, EventArgs e)
        {
            string path = this.textBox1.Text;
            var book = new LinqToExcel.ExcelQueryFactory(path);

            book.AddMapping<Person>(x => x.Name, "Name");
            book.AddMapping<Person>(x => x.Tickets, "# of tickets");

            var query = from x in book.Worksheet<Person>()
                        select x;

            int ttltickets= 0;
            foreach(var result in query)
            {
                String name = result.Name;
                ttltickets+=result.Tickets;
            }

            Random r = new Random();
            int winner = r.Next(1,ttltickets);

            ttltickets = 0;
            foreach (var result in query)
            {
                String name = result.Name;
                ttltickets += result.Tickets;
                if (ttltickets >= winner)
                {
                    Console.WriteLine(name);
                    break;
                }
            }

        }
    }
}
