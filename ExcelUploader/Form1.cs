using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelUploader
{
    public partial class Form1 : Form
    {
     
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opn = new OpenFileDialog();

            if (opn.ShowDialog() == DialogResult.OK)
            {
                txt_path.Text = opn.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            if (string.IsNullOrEmpty(txt_path.Text.Trim())) {
                MessageBox.Show("Can you enter a path? otherwise how would we know what to upload you fool!!!!");
                return;
            }
            if (string.IsNullOrEmpty(txt_sheet.Text.Trim()) || string.IsNullOrEmpty(txt_sql.Text.Trim()))
            {
                MessageBox.Show("You need to enter  Sheet name and Field names!!!!");
                return;
            }


            List<string> flist = new List<string>() {
                "Level3", "Extended" , "ValueX"
            };

            DataSet dset =  LoadExcel(txt_path.Text.Trim(), txt_sheet.Text.Trim(), flist);
            dataGridView1.DataSource = dset.Tables[0];
        }



        private DataSet LoadExcel(string filePath, string SheetName, List<string> FieldsList)
        {
            try
            {
                //[Level3],[Extended],[ValueX]
                
                var query = "Select " + GetFieldListString(FieldsList) + " from [" + SheetName + "$]";
                String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", filePath);

                OleDbConnection conn = new OleDbConnection(excelConnString);
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                OleDbCommand cmd = new OleDbCommand(query, conn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataSet ds = new DataSet();
                da.Fill(ds);
                da.Dispose();
                conn.Close();
                conn.Dispose();

                return ds;
            }
            catch (Exception ex) {
                throw ex;
            }

        }

        private string GetFieldListString(List<string> FieldsList)
        {
            string str = "";
            foreach (var item in FieldsList)
            {
                str += item.Trim() + " ,";
            }

            return str.Substring(0,str.Length - 1) ;
        }
       

    }
}
