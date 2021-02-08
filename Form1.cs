using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace NOPIxls
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime starttime, endtime;

                this.dataGridView1.DataSource = null;
                this.dataGridView1.Rows.Clear();

                starttime = Convert.ToDateTime(this.dateTimePicker1.Text + " 0:00:00");//起始时间
                endtime = Convert.ToDateTime(this.dateTimePicker2.Text + " 23:59:59"); ;//结束时间

                string strsel = "";
                int n = 0;

                strsel = "Select * from SystemLog where LogTime between #" + starttime + "# and #" + endtime + "#";
               

                if (strsel != "")
                {
                    

                    OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;data source=MainData.mdb;Jet OLEDB:database password=jyxjdz");
                    OleDbCommand cmd = new OleDbCommand(strsel, con);
                    OleDbDataAdapter ada = new OleDbDataAdapter(cmd);
                    OleDbDataReader rd;
                    con.Open();
                    rd = cmd.ExecuteReader();
                    while (rd.Read())
                    {
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[n].Cells[0].Value = rd[0].ToString();
                        this.dataGridView1.Rows[n].Cells[1].Value = rd[1].ToString();
                        this.dataGridView1.Rows[n].Cells[2].Value = rd[2].ToString();
                        n += 1;
                    }
                    rd.Close();
                    con.Close();
                }
                else
                {
                    MessageBox.Show("未选择查询条件，无法获取信息", "提示");
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Rows.Clear();
                    //this.dataGridView2.Columns.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
