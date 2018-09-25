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
            string[] paths = this.textBox1.Text.Split(',');
            int[] numbers = new int[paths.Length];

            //Count the number of tickets in total

            int ttltickets = 0;
            int ind = 0;
            foreach (string path in paths) {

                var book = new LinqToExcel.ExcelQueryFactory(path);

                book.AddMapping<Person>(x => x.Name, "Name");
                book.AddMapping<Person>(x => x.Tickets, "# of tickets");

                var query1 = from x in book.Worksheet<Person>()
                            select x;

                foreach (var result in query1)
                {
                    String name = result.Name;
                    ttltickets += result.Tickets;
                }

                numbers[ind] = ttltickets;
                ind++;

            }


            Random r = new Random();
            int winner = r.Next(1,ttltickets);

            ind = 0;
            while(numbers[ind]< winner && ind<= numbers.Length-1)
            {
                ind++;
            }

            string selectedPath = paths[ind];
            int countFrom = 0;
            if(ind> 0)
            {
                countFrom = numbers[ind - 1];
            }

            var book1 = new LinqToExcel.ExcelQueryFactory(selectedPath);

            book1.AddMapping<Person>(x => x.Name, "Name");
            book1.AddMapping<Person>(x => x.Tickets, "# of tickets");

            var query = from x in book1.Worksheet<Person>()
                        select x;
            ttltickets = countFrom;
            foreach (var result in query)
            {
                String name = result.Name;
                ttltickets += result.Tickets;
                if(ttltickets >= winner)
                {
                    Console.WriteLine(name);
                    break;
                }
            }

            numbers[ind] = ttltickets;
            ind++;

        }
    }
}
