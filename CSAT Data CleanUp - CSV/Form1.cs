using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.IO;

namespace CSAT_Data_CleanUp___CSV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string filePath1 = "";
        private string filePath2 = "";
        List<LineItem> rows = new List<LineItem>();
        LineItem row;
        LineItem header;
        private void button1_Click(object sender, EventArgs e)
        {

            openFileDialog1.Filter = "csv Files (*.csv)|*.csv";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                filePath1 = openFileDialog1.FileName;
            else
                filePath1 = string.Empty;
            if (!string.IsNullOrEmpty(filePath1))
            {
                TextFieldParser parser = new TextFieldParser(filePath1);

                parser.HasFieldsEnclosedInQuotes = true;
                parser.SetDelimiters(",");

                string[] fields;

                int counter = 0;
                while (!parser.EndOfData)
                {
                        
                    fields = parser.ReadFields();
                        if (!string.IsNullOrEmpty(fields[1]) && !string.IsNullOrEmpty(fields[3]) && !fields[1].Equals("No Agent", StringComparison.OrdinalIgnoreCase) &&
                            !fields[3].Equals("No Reply", StringComparison.OrdinalIgnoreCase) && !fields[5].Equals("noreply@lexisnexis.com", StringComparison.OrdinalIgnoreCase)
                            && !fields[6].Equals("LexisNexis", StringComparison.OrdinalIgnoreCase) && !fields[5].Equals("jps@lexisnexis.com", StringComparison.OrdinalIgnoreCase)
                            &&!fields[4].Equals("LNG-RDU-JurisProfessionalSolutions", StringComparison.OrdinalIgnoreCase) && !fields[3].Equals("JS Custom Reports", StringComparison.OrdinalIgnoreCase))
                        {
                            if (counter == 0)
                            {
                                header = new LineItem();
                                header.ticketNum = fields[0];
                                header.agent = fields[1];
                                header.closedTime = fields[2];
                                header.custID = fields[3];
                                header.name = fields[4];
                                header.email = fields[5];
                                header.company = fields[6];

                            }
                            else
                            {
                                row = new LineItem();
                                row.ticketNum = fields[0];
                                row.agent = fields[1];
                                row.closedTime = fields[2];
                                row.custID = fields[3];
                                row.name = fields[4];
                                row.email = fields[5];
                                row.company = fields[6];


                                rows.Add(row);
                            }
                        }
                    
                    counter++;
                }

                parser.Close();
            }
            button3.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<LineItem> output = rows.GroupBy(x => x.company).Select(x => x.First()).ToList();
            string dir = Path.GetDirectoryName(filePath1);
            DateTime dt = DateTime.Now;

            StreamWriter write = new StreamWriter(dir + @"\CSATOutput-" + dt.ToShortDateString().Replace("/", "-") + ".csv");
            write.WriteLine(header.ticketNum + "," + header.agent + "," + header.closedTime + "," + header.custID + "," + header.name + "," + header.email + "," + header.company);
            foreach (LineItem item in output)
                write.WriteLine(item.ticketNum + "," + item.agent + "," + item.closedTime + "," + item.custID + ",\"" + item.name + "\",\"" + item.email + "\",\"" + item.company + "\"");
            write.Flush();
            write.Close();
            MessageBox.Show("Done");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "csv Files (*.csv)|*.csv";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                filePath2 = openFileDialog1.FileName;
            else
                filePath2 = string.Empty;
            if (!string.IsNullOrEmpty(filePath2))
            {
                TextFieldParser parser = new TextFieldParser(filePath2);

                parser.HasFieldsEnclosedInQuotes = true;
                parser.SetDelimiters(",");

                string[] fields;

                int counter = 0;
                while (!parser.EndOfData)
                {

                    fields = parser.ReadFields();
                    if (!string.IsNullOrEmpty(fields[1]) && !string.IsNullOrEmpty(fields[3]) && !fields[1].Equals("No Agent", StringComparison.OrdinalIgnoreCase) &&
                        !fields[3].Equals("No Reply", StringComparison.OrdinalIgnoreCase) && !fields[5].Equals("noreply@lexisnexis.com", StringComparison.OrdinalIgnoreCase)
                        && !fields[6].Equals("LexisNexis", StringComparison.OrdinalIgnoreCase) && !fields[5].Equals("jps@lexisnexis.com", StringComparison.OrdinalIgnoreCase)
                        && !fields[4].Equals("LNG-RDU-JurisProfessionalSolutions", StringComparison.OrdinalIgnoreCase) && !fields[3].Equals("JS Custom Reports", StringComparison.OrdinalIgnoreCase))
                    {
                        if (counter != 0)
                        {
                            row = new LineItem();
                            row.ticketNum = fields[0];
                            row.agent = fields[1];
                            row.closedTime = fields[2];
                            row.custID = fields[3];
                            row.name = fields[4];
                            row.email = fields[5];
                            row.company = fields[6];


                            rows.Add(row);

                        }
                    }

                    counter++;
                }

                parser.Close();
            }
            button2.Enabled = true;
        }
    }
}
