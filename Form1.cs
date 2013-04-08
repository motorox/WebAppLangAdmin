using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using System.Xml;
using System.Xml.Xsl;
using ExportToExcel;

namespace WebAppLangAdmin
{
    public partial class LangAdminForm : Form
    {
        private Dictionary<string, string> jsonElements;
        private DataTable myTable;

        public LangAdminForm()
        {
            InitializeComponent();
            //if (MessageBox.Show(
            //    "Have you saved your work?\nOpening a new file will clear out all list boxes.",
            //    "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            //{ Application.Exit(); }
            myTable = new DataTable();
            myTable.Columns.Add("label".ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int nrJsoane = 0;
            int size = -1;
            string notfound = "";
            string file = "";
            DialogResult result = this.openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                try
                {
                    string text = File.ReadAllText(file);
                    size = text.Length;
                    this.jsonElements = JsonConvert.DeserializeObject<Dictionary<string, string>>(text);
                    nrJsoane = this.jsonElements.Count;
                    this.comboBoxFileList.Items.Add(file);
                }
                catch (IOException)
                {
                    Console.WriteLine("Error reading file: " + file);
                    return;
                }

                // reading the list of lines and insert them into the datagrid.
                if (this.comboBoxFileList.Items.Count>1)
                {
                    // already one file added
                    myTable.Columns.Add(this.comboBoxFileList.Items.Count.ToString());
                    foreach (DataRow row in myTable.Rows)
                    {
                        if (this.jsonElements.ContainsKey(row["label"].ToString()))
                        {
                            row[this.comboBoxFileList.Items.Count.ToString()] = this.jsonElements[row["label"].ToString()];
                            this.jsonElements.Remove(row["label"].ToString());
                        }
                        else
                        {
                            Console.WriteLine("Key not found: " + row["label"].ToString());
                            notfound += row["label"].ToString() + " ";
                        }
                    }
                    //adding the remaining new items in the table.
                    foreach (var line in this.jsonElements)
                    {
                        DataRow dr = myTable.NewRow();
                        dr["label"] = line.Key;
                        dr[this.comboBoxFileList.Items.Count.ToString()] = line.Value;
                        myTable.Rows.Add(dr);
                    }
                }
                else
                {
                    myTable.Columns.Add(this.comboBoxFileList.Items.Count.ToString());
                    foreach(var line in this.jsonElements){
                        DataRow dr = myTable.NewRow();
                        dr["label"] = line.Key;
                        dr[this.comboBoxFileList.Items.Count.ToString()] = line.Value;

                        myTable.Rows.Add(dr);
                    }
                }
                dataGridView1.DataSource = myTable.DefaultView;
                this.comboBoxFileList.SelectedItem = file;
            }
            //Console.WriteLine("Dialog result: " + result.ToString());
            string strResult = String.Format("File: {0} - JSon elements: {1}\n\nKeys added: {2}\n\nKeys missing: {3}", file, nrJsoane, this.jsonElements.Keys.Count, notfound);
            MessageBox.Show(strResult, "Open report", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void comboBoxFileList_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Columns[this.comboBoxFileList.SelectedIndex].HeaderCell.Style.BackColor = Color.Red;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFilename = "C:\\Sample.xlsx";
                CreateExcelFile.CreateExcelDocument(this.myTable, excelFilename);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Couldn't create Excel file.\r\nException: " + ex.Message);
                return;
            }
            MessageBox.Show("Excel file successfuly created.");
            return;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string file;
            DialogResult result = this.saveFileDialog1.ShowDialog();
            if (result != DialogResult.Cancel)
            {
                file = saveFileDialog1.FileName;
                Console.WriteLine("Save to: " + file);
            }
        }


    }
}
