using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesktopApp4
{
    public partial class Form1 : Form
    {
        private OleDbConnection connection = new OleDbConnection();
        DataSet ds = new DataSet();
        OleDbDataAdapter da;
        string t1 = "0", t2 = "0", t3 = "0", t4 = "0", t5 = "0", t6 = "0", t7 = "0", t8 = "0", t9 = "0";
        string m1 = "0", m2 = "0", m3 = "0", m4 = "0", m5 = "0", m6 = "0", m7 = "0", m8 = "0";
        Thread th;
        private void ToolStripButton3_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0)
                {

                    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                    xcelApp.Application.Workbooks.Add(Type.Missing);

                    for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                    {
                        xcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                    }

                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            xcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToolStripButton1_Click(object sender, EventArgs e)
        {
            this.Close();
            th = new Thread(opennewform);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }
        private void opennewform(object obj)
        {
            Application.Run(new MainPage());

        }

        private void ToolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("Are you sure to save Changes", "Message", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
                if (dr == DialogResult.Yes)
                {
                    OleDbCommandBuilder cmd = new OleDbCommandBuilder(da);
                    da.Update(ds, "dataTable");
                    dataGridView1.Refresh();
                    MessageBox.Show("Record Updated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column2_KeyPress);
            if (dataGridView1.CurrentCell.ColumnIndex == 7 || dataGridView1.CurrentCell.ColumnIndex == 9 || dataGridView1.CurrentCell.ColumnIndex == 14
                || dataGridView1.CurrentCell.ColumnIndex == 16 || dataGridView1.CurrentCell.ColumnIndex == 21 || dataGridView1.CurrentCell.ColumnIndex == 24
                || dataGridView1.CurrentCell.ColumnIndex == 26 || dataGridView1.CurrentCell.ColumnIndex == 31 || dataGridView1.CurrentCell.ColumnIndex == 33
                || dataGridView1.CurrentCell.ColumnIndex == 38 || dataGridView1.CurrentCell.ColumnIndex == 40 || dataGridView1.CurrentCell.ColumnIndex == 45
                || dataGridView1.CurrentCell.ColumnIndex == 47 || dataGridView1.CurrentCell.ColumnIndex == 52 || dataGridView1.CurrentCell.ColumnIndex == 54
                || dataGridView1.CurrentCell.ColumnIndex == 59 || dataGridView1.CurrentCell.ColumnIndex == 61 || dataGridView1.CurrentCell.ColumnIndex == 2)
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column2_KeyPress);
                }

            }
            e.Control.KeyPress -= new KeyPressEventHandler(Column3_KeyPress);
            if (dataGridView1.CurrentCell.ColumnIndex == 5)
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column3_KeyPress);
                }

            }


        }

        private void Column2_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar) && e.KeyChar != 46 && e.KeyChar != 45)
            {

                e.Handled = true;
            }
        }
        private void Column3_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!char.IsNumber(e.KeyChar) && !char.IsControl(e.KeyChar) && e.KeyChar != 47)
            {

                e.Handled = true;
            }
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }

        public Form1()
        {
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\SDatabase.accdb";
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnHeadersHeight = dataGridView1.ColumnHeadersHeight * 2;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            dataGridView1.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);
            dataGridView1.Paint += new PaintEventHandler(dataGridView1_Paint);
            dataGridView1.Scroll += new ScrollEventHandler(dataGridView1_Scroll);
            dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);
            try
            {
                // TODO: This line of code loads data into the 'sDatabaseDataSet.dataTable' table. You can move, or remove it, as needed.
                this.dataTableTableAdapter.Fill(this.sDatabaseDataSet.dataTable);

                da = new OleDbDataAdapter("SELECT ID,StudentName,Symbol,MotherName,FatherName,Address," +
                                          "DOB,NepaliTh,NThGrade,NepaliPr,NPrGrade,NepaliTotal,NTotalGrade,NepalliGPA,EnglishTh,EThGrade,EnglishPr,EPrGrade,EnglishTotal," +
                                          "ETotalGrade,EnglishGPA,MathematicsTh,MGrade,MathGPA,SocialTh,SoThGrade,SocialPr,SoPrGrade,SocialTotal," +
                                          "SoTotalGrade,SocialGPA,ScienceTh,ScThGrade,SciencePr,ScPrGrade,ScienceTotal,ScTotalGrade," +
                                          "ScienceGPA,HealthTh,HThGrade,HealthPr,HPrGrade,HealthTotal,HTotalGrade,HealthGPA,MoralTh,MThGrade,MoralPr,MPrGrade," +
                                          "MoralTotal,MTotalGrade,MoralGPA,BusinessTh,BThGrade,BusinessPr,BPrGrade,BusinessTotal,BTotalGrade,BusinessGPA," +
                                          "LocalTh,LThGrade,LocalPr,LPrGrade,LocalTotal,LTotalGrade,LocalGPA,TotalTh,TotalThGrade,TotalPr,TotalPrGrade," +
                                        "Total,Grade,GPA,Remark from dataTable", connection.ConnectionString);



                //DataSet ds = new DataSet();
                ds = new System.Data.DataSet();
                da.Fill(ds, "dataTable");



                dataGridView1.DataSource = ds.Tables["dataTable"];
            }
            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = MessageBox.Show("Are you sure to save Changes", "Message", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
                if (dr == DialogResult.Yes)
                {
                    OleDbCommandBuilder cmd = new OleDbCommandBuilder(da);
                    da.Update(ds, "dataTable");
                    dataGridView1.Refresh();
                    MessageBox.Show("Record Updated");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                try
                {
                    if (e.ColumnIndex == 7 || e.ColumnIndex == 9)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[7].Value != null)
                        {
                            t1 = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
                            if (t1.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[8].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[9].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[10].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[11].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[12].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[13].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[8].Value = Grade((double.Parse(t1) / 75) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[9].Value != DBNull.Value)
                                {
                                    m1 = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();

                                    dataGridView1.Rows[e.RowIndex].Cells[10].Value = Grade((double.Parse(m1) / 25) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[11].Value = Convert.ToDouble(t1) + Convert.ToDouble(m1);
                                    dataGridView1.Rows[e.RowIndex].Cells[12].Value = Grade(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString()));
                                    dataGridView1.Rows[e.RowIndex].Cells[13].Value = GPA(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[11].Value.ToString()));
                                }
                            }

                        }

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                // english
                try
                {

                    if (e.ColumnIndex == 14 || e.ColumnIndex == 16)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[14].Value != null)
                        {
                            t2 = dataGridView1.Rows[e.RowIndex].Cells[14].Value.ToString();
                            if (t2.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[14].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[15].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[16].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[17].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[18].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[19].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[20].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[15].Value = Grade((double.Parse(t2) / 75) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[16].Value != DBNull.Value)
                                {
                                    m2 = dataGridView1.Rows[e.RowIndex].Cells[16].Value.ToString();
                                    dataGridView1.Rows[e.RowIndex].Cells[17].Value = Grade((double.Parse(m2) / 25) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[18].Value = Convert.ToDouble(t2) + Convert.ToDouble(m2);
                                    dataGridView1.Rows[e.RowIndex].Cells[19].Value = Grade(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[18].Value.ToString()));
                                    dataGridView1.Rows[e.RowIndex].Cells[20].Value = GPA(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[18].Value.ToString()));
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //math
                try
                {
                    if (e.ColumnIndex == 21)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[21].Value != null)
                        {
                            t3 = dataGridView1.Rows[e.RowIndex].Cells[21].Value.ToString();
                            if (t3.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[22].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[23].Value = "Abs";

                            }
                            else
                            {

                                dataGridView1.Rows[e.RowIndex].Cells[22].Value = Grade(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[21].Value.ToString()));
                                dataGridView1.Rows[e.RowIndex].Cells[23].Value = GPA(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[21].Value.ToString()));

                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //social
                try
                {
                    if (e.ColumnIndex == 24 || e.ColumnIndex == 26)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[24].Value != null)
                        {
                            t4 = dataGridView1.Rows[e.RowIndex].Cells[24].Value.ToString();
                            if (t4.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[25].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[26].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[27].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[28].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[29].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[30].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[25].Value = Grade((double.Parse(t4) / 75) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[26].Value != DBNull.Value)
                                {
                                    m3 = dataGridView1.Rows[e.RowIndex].Cells[26].Value.ToString();
                                    dataGridView1.Rows[e.RowIndex].Cells[27].Value = Grade((double.Parse(m3) / 25) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[28].Value = Convert.ToDouble(t4) + Convert.ToDouble(m3);
                                    dataGridView1.Rows[e.RowIndex].Cells[29].Value = Grade(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[28].Value.ToString()));
                                    dataGridView1.Rows[e.RowIndex].Cells[30].Value = GPA(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[28].Value.ToString()));
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //science
                try
                {
                    if (e.ColumnIndex == 31 || e.ColumnIndex == 33)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[31].Value != null)
                        {
                            t5 = dataGridView1.Rows[e.RowIndex].Cells[31].Value.ToString();
                            if (t5.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[32].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[33].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[34].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[35].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[36].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[37].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[32].Value = Grade((double.Parse(t5) / 75) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[33].Value != DBNull.Value)
                                {
                                    m4 = dataGridView1.Rows[e.RowIndex].Cells[33].Value.ToString();
                                    dataGridView1.Rows[e.RowIndex].Cells[34].Value = Grade((double.Parse(m4) / 25) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[35].Value = Convert.ToDouble(t5) + Convert.ToDouble(m4);
                                    dataGridView1.Rows[e.RowIndex].Cells[36].Value = Grade(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[35].Value.ToString()));
                                    dataGridView1.Rows[e.RowIndex].Cells[37].Value = GPA(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[35].Value.ToString()));
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //health
                try
                {
                    if (e.ColumnIndex == 38 || e.ColumnIndex == 40)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[38].Value != null)
                        {
                            t6 = dataGridView1.Rows[e.RowIndex].Cells[38].Value.ToString();
                            if (t6.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[39].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[40].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[41].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[42].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[43].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[44].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[39].Value = Grade((double.Parse(t6) / 30) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[40].Value != DBNull.Value)
                                {
                                    m5 = dataGridView1.Rows[e.RowIndex].Cells[40].Value.ToString();
                                    dataGridView1.Rows[e.RowIndex].Cells[41].Value = Grade((double.Parse(m5) / 20) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[42].Value = Convert.ToDouble(t6) + Convert.ToDouble(m5);
                                    dataGridView1.Rows[e.RowIndex].Cells[43].Value = Grade((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[42].Value.ToString()) / 50) * 100);

                                    dataGridView1.Rows[e.RowIndex].Cells[44].Value = GPA((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[42].Value.ToString()) / 50) * 100);
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //moral
                try
                {
                    if (e.ColumnIndex == 45 || e.ColumnIndex == 47)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[45].Value != null)
                        {
                            t7 = dataGridView1.Rows[e.RowIndex].Cells[45].Value.ToString();
                            if (t7.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[46].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[47].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[48].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[49].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[50].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[51].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[46].Value = Grade((double.Parse(t7) / 25) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[47].Value != DBNull.Value)
                                {
                                    m6 = dataGridView1.Rows[e.RowIndex].Cells[47].Value.ToString();
                                    dataGridView1.Rows[e.RowIndex].Cells[48].Value = Grade((double.Parse(m6) / 25) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[49].Value = Convert.ToDouble(t7) + Convert.ToDouble(m6);
                                    dataGridView1.Rows[e.RowIndex].Cells[50].Value = Grade((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[49].Value.ToString()) / 50) * 100);

                                    dataGridView1.Rows[e.RowIndex].Cells[51].Value = GPA((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[49].Value.ToString()) / 50) * 100);
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //occuption
                try
                {
                    if (e.ColumnIndex == 52 || e.ColumnIndex == 54)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[52].Value != null)
                        {
                            t8 = dataGridView1.Rows[e.RowIndex].Cells[52].Value.ToString();
                            if (t8.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[53].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[54].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[55].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[56].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[57].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[58].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[53].Value = Grade((double.Parse(t8) / 50) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[54].Value != DBNull.Value)
                                {
                                    m7 = dataGridView1.Rows[e.RowIndex].Cells[54].Value.ToString();
                                    dataGridView1.Rows[e.RowIndex].Cells[55].Value = Grade((double.Parse(m7) / 50) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[56].Value = Convert.ToDouble(t8) + Convert.ToDouble(m7);
                                    dataGridView1.Rows[e.RowIndex].Cells[57].Value = Grade(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[56].Value.ToString()));
                                    dataGridView1.Rows[e.RowIndex].Cells[58].Value = GPA(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[56].Value.ToString()));
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //local
                try
                {
                    if (e.ColumnIndex == 59 || e.ColumnIndex == 61)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[59].Value != null)
                        {
                            t9 = dataGridView1.Rows[e.RowIndex].Cells[59].Value.ToString();
                            if (t9.Contains("-"))
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[60].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[61].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[62].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[63].Value = "-";
                                dataGridView1.Rows[e.RowIndex].Cells[64].Value = "Abs";
                                dataGridView1.Rows[e.RowIndex].Cells[65].Value = "Abs";
                            }
                            else
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[60].Value = Grade((double.Parse(t9) / 50) * 100);
                                if (dataGridView1.Rows[e.RowIndex].Cells[61].Value != DBNull.Value)
                                {
                                    m8 = dataGridView1.Rows[e.RowIndex].Cells[61].Value.ToString();
                                    dataGridView1.Rows[e.RowIndex].Cells[62].Value = Grade((double.Parse(m8) / 50) * 100);
                                    dataGridView1.Rows[e.RowIndex].Cells[63].Value = Convert.ToDouble(t9) + Convert.ToDouble(m8);
                                    dataGridView1.Rows[e.RowIndex].Cells[64].Value = Grade(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[63].Value.ToString()));
                                    dataGridView1.Rows[e.RowIndex].Cells[65].Value = GPA(double.Parse(dataGridView1.Rows[e.RowIndex].Cells[63].Value.ToString()));
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //total

                if (t1.Contains("-") || t2.Contains("-") || t3.Contains("-") || t4.Contains("-")
                    || t5.Contains("-") || t6.Contains("-") || t7.Contains("-") || t8.Contains("-") || t9.Contains("-"))
                {
                    dataGridView1.Rows[e.RowIndex].Cells[66].Value = "Abs";
                    dataGridView1.Rows[e.RowIndex].Cells[67].Value = "-";
                    dataGridView1.Rows[e.RowIndex].Cells[68].Value = "Abs";
                    dataGridView1.Rows[e.RowIndex].Cells[69].Value = "-";
                    dataGridView1.Rows[e.RowIndex].Cells[70].Value = "Abs";
                    dataGridView1.Rows[e.RowIndex].Cells[71].Value = "Abs";
                    dataGridView1.Rows[e.RowIndex].Cells[72].Value = "Abs";
                    dataGridView1.Rows[e.RowIndex].Cells[73].Value = "Abs";

                }
                else
                {
                    if (dataGridView1.Rows[e.RowIndex].Cells[7].Value != null || dataGridView1.Rows[e.RowIndex].Cells[14].Value != null)
                    {

                        dataGridView1.Rows[e.RowIndex].Cells[66].Value = Convert.ToDouble(t1) + Convert.ToDouble(t2) + Convert.ToDouble(t3) + Convert.ToDouble(t4)
                          + Convert.ToDouble(t5) + Convert.ToDouble(t6) + Convert.ToDouble(t7) + Convert.ToDouble(t8) + Convert.ToDouble(t9);

                        dataGridView1.Rows[e.RowIndex].Cells[67].Value = Grade((Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[66].Value.ToString()) / 555) * 100);
                        if (dataGridView1.Rows[e.RowIndex].Cells[9].Value != DBNull.Value)
                        {
                            dataGridView1.Rows[e.RowIndex].Cells[68].Value = Convert.ToDouble(m1) + Convert.ToDouble(m2) + Convert.ToDouble(m3) + Convert.ToDouble(m4)
                           + Convert.ToDouble(m5) + Convert.ToDouble(m6) + Convert.ToDouble(m7) + Convert.ToDouble(m8);

                            dataGridView1.Rows[e.RowIndex].Cells[69].Value = Grade((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[68].Value.ToString()) / 245) * 100);
                            dataGridView1.Rows[e.RowIndex].Cells[70].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[66].Value) + Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[68].Value);
                            dataGridView1.Rows[e.RowIndex].Cells[71].Value = Grade((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[70].Value.ToString()) / 800) * 100);

                            dataGridView1.Rows[e.RowIndex].Cells[72].Value = GPA((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[70].Value.ToString()) / 800) * 100);
                            if (dataGridView1.Rows[e.RowIndex].Cells[70].Value != null)
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[73].Value = Rem((double.Parse(dataGridView1.Rows[e.RowIndex].Cells[70].Value.ToString()) / 800) * 100);
                            }
                        }
                    }
                }

            }


            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private string Rem(double num)
        {
            string myGrade = "";

            if (num >= 90 && num <= 100)
            {
                myGrade = "Outstanding";
            }
            else if (num >= 80 && num < 90)
            {
                myGrade = "Excellent";
            }
            else if (num >= 70 && num < 80)
            {
                myGrade = "Very Goood";
            }
            else if (num >= 60 && num < 70)
            {
                myGrade = "Good";
            }
            else if (num >= 50 && num < 60)
            {
                myGrade = " Above Average";
            }
            else if (num >= 40 && num < 50)
            {
                myGrade = "Average";
            }
            else if (num >= 30 && num < 40)
            {
                myGrade = "Below Average";
            }
            else if (num >= 20 && num < 30)
            {
                myGrade = "Insufficient";
            }
            else if (num < 20 && num > 0)
            {
                myGrade = "Insufficient";
            }
            else
            {
                myGrade = "Not Graded";
            }

            return myGrade;


        }

        private double GPA(double num)
        {
            double myGrade = 0.0;

            if (num >= 90 && num <= 100)
            {
                myGrade = 4.0;
            }
            else if (num >= 80 && num < 90)
            {
                myGrade = 3.6;
            }
            else if (num >= 70 && num < 80)
            {
                myGrade = 3.2;
            }
            else if (num >= 60 && num < 70)
            {
                myGrade = 2.8;
            }
            else if (num >= 50 && num < 60)
            {
                myGrade = 2.4;
            }
            else if (num >= 40 && num < 50)
            {
                myGrade = 2.0;
            }
            else if (num >= 30 && num < 40)
            {
                myGrade = 1.6;
            }
            else if (num >= 20 && num < 30)
            {
                myGrade = 1.2;
            }
            else if (num < 20 && num > 0)
            {
                myGrade = 0.8;
            }
            else
            {
                myGrade = 0;
            }

            return myGrade;


        }
        private string Grade(double num)
        {

            string myGrade = "";

            if (num >= 90 && num <= 100)
            {
                myGrade = "A+";
            }
            else if (num >= 80 && num < 90)
            {
                myGrade = "A";
            }
            else if (num >= 70 && num < 80)
            {
                myGrade = "B+";
            }
            else if (num >= 60 && num < 70)
            {
                myGrade = "B";
            }
            else if (num >= 50 && num < 60)
            {
                myGrade = "C+";
            }
            else if (num >= 40 && num < 50)
            {
                myGrade = "C";
            }
            else if (num >= 30 && num < 40)
            {
                myGrade = "D+";
            }
            else if (num >= 20 && num < 30)
            {
                myGrade = "D";
            }
            else if (num < 20 && num > 0)
            {
                myGrade = "E";
            }
            else
            {
                myGrade = "Undefinde";
            }

            return myGrade;


        }
        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            Rectangle rheader = dataGridView1.DisplayRectangle;
            rheader.Height = dataGridView1.ColumnHeadersHeight / 2;
            dataGridView1.Invalidate(rheader);

        }
        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {

            Rectangle r1 = dataGridView1.GetCellDisplayRectangle(7, -1, true);
            int w1 = dataGridView1.GetCellDisplayRectangle(8, -1, true).Width;
            int w2 = dataGridView1.GetCellDisplayRectangle(9, -1, true).Width;
            int w3 = dataGridView1.GetCellDisplayRectangle(10, -1, true).Width;
            int w4 = dataGridView1.GetCellDisplayRectangle(11, -1, true).Width;
            int w5 = dataGridView1.GetCellDisplayRectangle(12, -1, true).Width;
            int w6 = dataGridView1.GetCellDisplayRectangle(13, -1, true).Width;



            r1.X += 1;
            r1.Y += 1;
            r1.Width = r1.Width + w1 + w2 + w3 + w4 + w5 + w6 - 2;
            r1.Height = r1.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r1);
            StringFormat forma = new StringFormat();
            forma.Alignment = StringAlignment.Center;
            forma.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Nepali", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r1, forma);

            Rectangle r2 = dataGridView1.GetCellDisplayRectangle(14, -1, true);
            int w7 = dataGridView1.GetCellDisplayRectangle(15, -1, true).Width;
            int w8 = dataGridView1.GetCellDisplayRectangle(16, -1, true).Width;
            int w9 = dataGridView1.GetCellDisplayRectangle(17, -1, true).Width;
            int w10 = dataGridView1.GetCellDisplayRectangle(18, -1, true).Width;
            int w11 = dataGridView1.GetCellDisplayRectangle(19, -1, true).Width;
            int w12 = dataGridView1.GetCellDisplayRectangle(20, -1, true).Width;



            r2.X += 1;
            r2.Y += 1;
            r2.Width = r2.Width + w7 + w8 + w9 + w10 + w11 + w12 - 2;
            r2.Height = r2.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r2);
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            format.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("English", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r2, format);


            Rectangle r3 = dataGridView1.GetCellDisplayRectangle(21, -1, true);
            int w13 = dataGridView1.GetCellDisplayRectangle(22, -1, true).Width;
            int w14 = dataGridView1.GetCellDisplayRectangle(23, -1, true).Width;




            r3.X += 1;
            r3.Y += 1;
            r3.Width = r3.Width + w13 + w14 - 2;
            r3.Height = r3.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r3);
            StringFormat format1 = new StringFormat();
            format1.Alignment = StringAlignment.Center;
            format1.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Mathematics", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r3, format1);


            Rectangle r4 = dataGridView1.GetCellDisplayRectangle(24, -1, true);
            int w15 = dataGridView1.GetCellDisplayRectangle(25, -1, true).Width;
            int w16 = dataGridView1.GetCellDisplayRectangle(26, -1, true).Width;
            int w17 = dataGridView1.GetCellDisplayRectangle(27, -1, true).Width;
            int w18 = dataGridView1.GetCellDisplayRectangle(28, -1, true).Width;
            int w19 = dataGridView1.GetCellDisplayRectangle(29, -1, true).Width;
            int w20 = dataGridView1.GetCellDisplayRectangle(30, -1, true).Width;



            r4.X += 1;
            r4.Y += 1;
            r4.Width = r4.Width + w15 + w16 + w17 + w18 + w19 + w20 - 2;
            r4.Height = r4.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r4);
            StringFormat format2 = new StringFormat();
            format2.Alignment = StringAlignment.Center;
            format2.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Social", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r4, format2);











            Rectangle r5 = dataGridView1.GetCellDisplayRectangle(31, -1, true);
            int w21 = dataGridView1.GetCellDisplayRectangle(32, -1, true).Width;
            int w22 = dataGridView1.GetCellDisplayRectangle(33, -1, true).Width;
            int w23 = dataGridView1.GetCellDisplayRectangle(34, -1, true).Width;
            int w24 = dataGridView1.GetCellDisplayRectangle(35, -1, true).Width;
            int w25 = dataGridView1.GetCellDisplayRectangle(36, -1, true).Width;
            int w26 = dataGridView1.GetCellDisplayRectangle(37, -1, true).Width;



            r5.X += 1;
            r5.Y += 1;
            r5.Width = r5.Width + w21 + w22 + w23 + w24 + w25 + w26 - 2;
            r5.Height = r5.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r5);
            StringFormat format3 = new StringFormat();
            format3.Alignment = StringAlignment.Center;
            format3.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Science", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r5, format2);

            Rectangle r6 = dataGridView1.GetCellDisplayRectangle(38, -1, true);
            int w27 = dataGridView1.GetCellDisplayRectangle(39, -1, true).Width;
            int w28 = dataGridView1.GetCellDisplayRectangle(40, -1, true).Width;
            int w29 = dataGridView1.GetCellDisplayRectangle(41, -1, true).Width;
            int w30 = dataGridView1.GetCellDisplayRectangle(42, -1, true).Width;
            int w31 = dataGridView1.GetCellDisplayRectangle(43, -1, true).Width;
            int w32 = dataGridView1.GetCellDisplayRectangle(44, -1, true).Width;



            r6.X += 1;
            r6.Y += 1;
            r6.Width = r6.Width + w27 + w28 + w29 + w30 + w31 + w32 - 2;
            r6.Height = r6.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r6);
            StringFormat format4 = new StringFormat();
            format4.Alignment = StringAlignment.Center;
            format4.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Health & Physical Education", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r6, format4);

            Rectangle r7 = dataGridView1.GetCellDisplayRectangle(45, -1, true);
            int w33 = dataGridView1.GetCellDisplayRectangle(46, -1, true).Width;
            int w34 = dataGridView1.GetCellDisplayRectangle(47, -1, true).Width;
            int w35 = dataGridView1.GetCellDisplayRectangle(48, -1, true).Width;
            int w36 = dataGridView1.GetCellDisplayRectangle(49, -1, true).Width;
            int w37 = dataGridView1.GetCellDisplayRectangle(50, -1, true).Width;
            int w38 = dataGridView1.GetCellDisplayRectangle(51, -1, true).Width;



            r7.X += 1;
            r7.Y += 1;
            r7.Width = r7.Width + w33 + w34 + w35 + w36 + w37 + w38 - 2;
            r7.Height = r7.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r7);
            StringFormat format5 = new StringFormat();
            format5.Alignment = StringAlignment.Center;
            format5.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Moral Education", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r7, format5);

            Rectangle r8 = dataGridView1.GetCellDisplayRectangle(52, -1, true);
            int w39 = dataGridView1.GetCellDisplayRectangle(53, -1, true).Width;
            int w40 = dataGridView1.GetCellDisplayRectangle(54, -1, true).Width;
            int w41 = dataGridView1.GetCellDisplayRectangle(55, -1, true).Width;
            int w42 = dataGridView1.GetCellDisplayRectangle(56, -1, true).Width;
            int w43 = dataGridView1.GetCellDisplayRectangle(57, -1, true).Width;
            int w44 = dataGridView1.GetCellDisplayRectangle(58, -1, true).Width;



            r8.X += 1;
            r8.Y += 1;
            r8.Width = r8.Width + w39 + w40 + w41 + w42 + w43 + w44 - 2;
            r8.Height = r8.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r8);
            StringFormat format6 = new StringFormat();
            format6.Alignment = StringAlignment.Center;
            format6.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Occuption,business and Technology", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r8, format6);

            Rectangle r9 = dataGridView1.GetCellDisplayRectangle(59, -1, true);
            int w45 = dataGridView1.GetCellDisplayRectangle(60, -1, true).Width;
            int w46 = dataGridView1.GetCellDisplayRectangle(61, -1, true).Width;
            int w47 = dataGridView1.GetCellDisplayRectangle(62, -1, true).Width;
            int w48 = dataGridView1.GetCellDisplayRectangle(63, -1, true).Width;
            int w49 = dataGridView1.GetCellDisplayRectangle(64, -1, true).Width;
            int w50 = dataGridView1.GetCellDisplayRectangle(65, -1, true).Width;



            r9.X += 1;
            r9.Y += 1;
            r9.Width = r9.Width + w45 + w46 + w47 + w48 + w49 + w50 - 2;
            r9.Height = r9.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r9);
            StringFormat format7 = new StringFormat();
            format7.Alignment = StringAlignment.Center;
            format7.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Computer", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r9, format7);




            Rectangle r10 = dataGridView1.GetCellDisplayRectangle(66, -1, true);
            int w51 = dataGridView1.GetCellDisplayRectangle(67, -1, true).Width;
            int w52 = dataGridView1.GetCellDisplayRectangle(68, -1, true).Width;
            int w53 = dataGridView1.GetCellDisplayRectangle(69, -1, true).Width;
            int w54 = dataGridView1.GetCellDisplayRectangle(70, -1, true).Width;
            int w55 = dataGridView1.GetCellDisplayRectangle(71, -1, true).Width;
            int w56 = dataGridView1.GetCellDisplayRectangle(72, -1, true).Width;


            r10.X += 1;
            r10.Y += 1;
            r10.Width = r10.Width + w51 + w52 + w53 + w54 + w55 + w56 - 2;
            r10.Height = r10.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r10);
            StringFormat format8 = new StringFormat();
            format8.Alignment = StringAlignment.Center;
            format8.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Total", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r10, format8);



        }
        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            Rectangle rheader = dataGridView1.DisplayRectangle;
            rheader.Height = dataGridView1.ColumnHeadersHeight / 2;
            dataGridView1.Invalidate(rheader);

        }
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex > -1)
            {
                Rectangle r2 = e.CellBounds;
                r2.Y += e.CellBounds.Height / 2;
                r2.Height = e.CellBounds.Height / 2;
                e.PaintBackground(r2, true);
                e.PaintContent(r2);
                e.Handled = true;
            }

        }


    }
}