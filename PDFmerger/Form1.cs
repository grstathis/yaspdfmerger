using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Text.RegularExpressions;


namespace PDFmerger
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string outfilename = "";
        // Open the output document
        PdfDocument outputDocument = new PdfDocument();
        //datatable to connect to datagridview 
        DataTable pdf_files = new DataTable();
        


        private void Form1_Load(object sender, EventArgs e)
        {
            //pdf_files.Columns.Add("FileName", typeof(string));
            //pdf_files.Columns.Add("Pages", typeof(string));
           
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
	        if (result == DialogResult.OK){
                foreach (String file in openFileDialog1.FileNames){
                    
                    // idx is the new index. The cells must also be accessed by an index
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = file;
                    
                }
                
	        }
	        
	    }

        private void merge_Click(object sender, EventArgs e)
        {

                 

            
            
            
            // If the file name is not an empty string display the result name in textbox for saving.
            if (textBox2.Text != "")
            {
                //textBox2.Text = saveFileDialog1.FileName;
            }
            else
            {
                // Displays a SaveFileDialog so the user can save the pdf

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "PDF file|*.pdf";
                saveFileDialog1.Title = "Save the Output File";
                saveFileDialog1.ShowDialog();
            }

            if (File.Exists(textBox2.Text))
            {
                File.Delete(textBox2.Text);
            }
            

            //Loop files selected...
            int nrows = dataGridView1.RowCount;
            int it = 0;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                it++;
                if (it == nrows) { break; }
                string file, pages;
                string[] spages;
                spages = new String[1];
                if (row.Cells[0].Value != null)
                {
                    file = row.Cells[0].Value.ToString();
                }
                else
                {
                    file = "";
                    MessageBox.Show("Nothing to Merge!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }
                if (row.Cells[1].Value != null)
                {
                    pages = row.Cells[1].Value.ToString();
                    pages = Regex.Replace(pages,@"\s+","");
                    pages = Regex.Replace(pages, @"[^(0-9)+,\-].", "");
                    pages = Regex.Replace(pages, @",+", ",");
                    Match match = Regex.Match(pages, @",");
                    if (match.Success)
                    {
                        spages = pages.Split(',');
                    }
                    else
                    {
                        spages[0] = pages;
                    }
                }
                else 
                {
                    pages = "all";
                }
                   
                if (File.Exists(file) == true){

                        PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import);
                        int count = inputDocument.PageCount;
                 
                        if (pages == "all")
                        {
                            for (int idx = 0; idx < count; idx++)
                            {
                                // Get the page from the external document...
                                PdfPage page = inputDocument.Pages[idx];
                                // ...and add it to the output document.
                                outputDocument.AddPage(page);
                            }
                        }
                        else
                        {
                            foreach (string spage in spages)
                            {
                                if (spage != "")
                                {
                                    Match match = Regex.Match(spage, @"-");
                                    if (match.Success)
                                    {
                                        string[] StartEndPages = spage.Split('-');
                                        int StartPage = Convert.ToInt32(StartEndPages[0]);
                                        int EndPage = Convert.ToInt32(StartEndPages[StartEndPages.Count() - 1]);
                                        if (StartPage > EndPage)
                                        {
                                            MessageBox.Show("The page numbers in file " + file + " are incorrect placed!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            DialogResult dialogResult = MessageBox.Show("Do you want to fix this automatically?", "Warning", MessageBoxButtons.YesNo);
                                            if (dialogResult == DialogResult.Yes)
                                            {
                                                int temp = StartPage;
                                                StartPage = EndPage;
                                                EndPage = temp;
                                            }
                                            else if (dialogResult == DialogResult.No)
                                            {
                                                continue;
                                            }


                                        }
                                        if (StartPage > count || EndPage > count)
                                        {
                                            MessageBox.Show("The page numbers in file " + file + " are greater than permitted!", "Warning",MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            DialogResult dialogResult = MessageBox.Show("Do you want to fix this automatically?", "Help", MessageBoxButtons.YesNo);
                                            if (dialogResult == DialogResult.Yes)
                                            {
                                                if (StartPage > count)
                                                {
                                                    StartPage = count;
                                                }
                                                if (EndPage > count)
                                                {
                                                    EndPage = count;
                                                }
                                            }
                                            else if (dialogResult == DialogResult.No)
                                            {
                                                MessageBox.Show("Continouning with Errors");
                                                continue;

                                            }
                                        }


                                        for (int idx = StartPage-1; idx <= EndPage-1; idx++)
                                        {
                                            // Get the page from the external document...
                                            PdfPage page = inputDocument.Pages[idx];
                                            // ...and add it to the output document.
                                            outputDocument.AddPage(page);
                                        }
                                    }
                                    else
                                    {
                                        int StartPage = Convert.ToInt32(spage);
                                        if (StartPage > count)
                                        {
                                            MessageBox.Show("The page numbers in file " + file + " are greater than permitted", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            DialogResult dialogResult = MessageBox.Show("Do you want to fix this automatically?", "Warning", MessageBoxButtons.YesNo);
                                            if (dialogResult == DialogResult.Yes)
                                            {
                                                if (StartPage > count)
                                                {
                                                    StartPage = count;
                                                }
                                            }
                                            else if (dialogResult == DialogResult.No)
                                            {
                                                continue;
                                            }
                                        }
                                        
                                            for (int idx = StartPage-1; idx <= StartPage-1; idx++)
                                            {
                                                // Get the page from the external document...
                                                PdfPage page = inputDocument.Pages[idx];
                                                // ...and add it to the output document.
                                                outputDocument.AddPage(page);
                                            }
                                        
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(file + " does not exist!");
                    }
                    

             }
           
            outfilename = textBox2.Text;
            if (outfilename != "")
            {
                if (outputDocument.PageCount != 0)
                {
                    outputDocument.Save(outfilename);
                    MessageBox.Show("PDF creationed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Output PDF has no pages, nothing to save", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please provide Output FileName!", "Warning",MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

       

        private void button2_Click(object sender, EventArgs e)
        {

            // Displays a SaveFileDialog so the user can save the pdf
            // assigned to Button2.
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "PDF file|*.pdf";
            saveFileDialog1.Title = "Save the Output File";
            saveFileDialog1.ShowDialog();

            // If the file name is not an empty string display the result name in textbox for saving.
            if (saveFileDialog1.FileName != "")
            {
                textBox2.Text = saveFileDialog1.FileName;
            }


        }


        
     }
    
}
