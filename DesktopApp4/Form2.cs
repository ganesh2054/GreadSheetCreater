using DesktopApp4.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
namespace DesktopApp4
{
    public partial class Form2 : Form
    {
        DataSet ds = new DataSet();
        OleDbDataAdapter da;

        Thread th;
     
        int SelectedRowIndex;
        private OleDbConnection connection = new OleDbConnection();
        public Form2()
        {
            var currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
            AppDomain.CurrentDomain.SetData("DataDirectory", currentDirectory);


            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|SDatabase.accdb";
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 46;
            dataGridView1.Columns[0].Name = "S.N.";
            dataGridView1.Columns[1].Name = "Student's Name";
            dataGridView1.Columns[2].Name = "Symbol No";
            dataGridView1.Columns[3].Name = "Mother's Name";
            dataGridView1.Columns[4].Name = "Father's Name";
            dataGridView1.Columns[5].Name = "DOB";
            dataGridView1.Columns[6].Name = "Address";
            dataGridView1.Columns[7].Name = "Nepali Th";

            dataGridView1.Columns[8].Name = "Nepali Pr";

            dataGridView1.Columns[9].Name = "Total";
            ;
            dataGridView1.Columns[10].Name = "GPA";
            dataGridView1.Columns[11].Name = "English Th";

            dataGridView1.Columns[12].Name = "English Pr";

            dataGridView1.Columns[13].Name = "Total";

            dataGridView1.Columns[14].Name = "GPA";
            dataGridView1.Columns[15].Name = "Mathematich Th";


            dataGridView1.Columns[16].Name = "GPA";
            dataGridView1.Columns[17].Name = "Social Th";

            dataGridView1.Columns[18].Name = "Social Pr";

            dataGridView1.Columns[19].Name = "Total";

            dataGridView1.Columns[20].Name = "GPA";
            dataGridView1.Columns[21].Name = "Science Th";

            dataGridView1.Columns[22].Name = "Science Pr";

            dataGridView1.Columns[23].Name = "Total";

            dataGridView1.Columns[24].Name = "GPA";
            dataGridView1.Columns[25].Name = "Health Th";

            dataGridView1.Columns[26].Name = "Health Pr";

            dataGridView1.Columns[27].Name = "Total";

            dataGridView1.Columns[28].Name = "GPA";
            dataGridView1.Columns[29].Name = "Moral Th";

            dataGridView1.Columns[30].Name = "Moral Pr";

            dataGridView1.Columns[31].Name = "Total";

            dataGridView1.Columns[32].Name = "GPA";
            dataGridView1.Columns[33].Name = "Occuption Th";

            dataGridView1.Columns[34].Name = "Occuption Pr";


            dataGridView1.Columns[35].Name = "Total";

            dataGridView1.Columns[36].Name = "GPA";
            dataGridView1.Columns[37].Name = "Computer Th";

            dataGridView1.Columns[38].Name = "Computer Pr";

            dataGridView1.Columns[39].Name = "Total";

            dataGridView1.Columns[40].Name = "GPA";
            dataGridView1.Columns[41].Name = "Total Th";

            dataGridView1.Columns[42].Name = "Total Pr";

            dataGridView1.Columns[43].Name = "Total";

            dataGridView1.Columns[44].Name = "GPA";
            dataGridView1.Columns[45].Name = "Remark";
            dataGridView1.ColumnHeadersHeight = dataGridView1.ColumnHeadersHeight * 2;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            dataGridView1.CellPainting += new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);
            dataGridView1.Paint += new PaintEventHandler(dataGridView1_Paint);
            dataGridView1.Scroll += new ScrollEventHandler(dataGridView1_Scroll);
            dataGridView1.ColumnWidthChanged += new DataGridViewColumnEventHandler(dataGridView1_ColumnWidthChanged);

            //Select from database
            try
            {
                using (OleDbConnection con = new OleDbConnection(connection.ConnectionString))
                {
                    con.Open();
                    da = new OleDbDataAdapter("SELECT StudentName,Symbol,MotherName,FatherName,DOB,Address,NThGrade,NPrGrade,NTotalGrade,NepalliGPA," +
                        "EThGrade,EPrGrade,ETotalGrade,EnglishGPA,MGrade,MathGPA,SoThGrade,SoPrGrade,SoTotalGrade,SocialGPA" +
                        ",ScThGrade,ScPrGrade,ScTotalGrade,ScienceGPA,HThGrade,HPrGrade,HTotalGrade,HealthGPA," +
                        "MThGrade,MPrGrade,MTotalGrade,MoralGPA,BThGrade,BPrGrade,BTotalGrade,BusinessGPA," +
                        "LThGrade,LPrGrade,LTotalGrade,LocalGPA,TotalThGrade,TotalPrGrade,Grade,GPA,Remark from dataTable", connection.ConnectionString);


                    {
                        ds = new System.Data.DataSet();
                        da.Fill(ds, "dataTable");



                        //cmd.CommandType = CommandType.Text;
                        //sda = new OleDbDataAdapter(cmd);
                        //sda.Fill(dt);
                        //Add columns
                        //dataGridView1.Columns[0].DataPropertyName = "ID";
                        dataGridView1.Columns[1].DataPropertyName = "StudentName";
                        dataGridView1.Columns[2].DataPropertyName = "Symbol";
                        dataGridView1.Columns[3].DataPropertyName = "MotherName";
                        dataGridView1.Columns[4].DataPropertyName = "FatherName";
                        dataGridView1.Columns[5].DataPropertyName = "DOB";
                        dataGridView1.Columns[6].DataPropertyName = "Address";
                        dataGridView1.Columns[7].DataPropertyName = "NThGrade";
                        dataGridView1.Columns[8].DataPropertyName = "NPrGrade";
                        dataGridView1.Columns[9].DataPropertyName = "NTotalGrade";
                        dataGridView1.Columns[10].DataPropertyName = "NepalliGPA";
                        dataGridView1.Columns[11].DataPropertyName = "EThGrade";

                        dataGridView1.Columns[12].DataPropertyName = "EPrGrade";

                        dataGridView1.Columns[13].DataPropertyName = "ETotalGrade";

                        dataGridView1.Columns[14].DataPropertyName = "EnglishGPA";
                        dataGridView1.Columns[15].DataPropertyName = "MGrade";


                        dataGridView1.Columns[16].DataPropertyName = "MathGPA";
                        dataGridView1.Columns[17].DataPropertyName = "SoThGrade";

                        dataGridView1.Columns[18].DataPropertyName = "SoPrGrade";

                        dataGridView1.Columns[19].DataPropertyName = "SoTotalGrade";

                        dataGridView1.Columns[20].DataPropertyName = "SocialGPA";
                        dataGridView1.Columns[21].DataPropertyName = "ScThGrade";

                        dataGridView1.Columns[22].DataPropertyName = "ScPrGrade";

                        dataGridView1.Columns[23].DataPropertyName = "ScTotalGrade";

                        dataGridView1.Columns[24].DataPropertyName = "ScienceGPA";
                        dataGridView1.Columns[25].DataPropertyName = "HThGrade";

                        dataGridView1.Columns[26].DataPropertyName = "HPrGrade";

                        dataGridView1.Columns[27].DataPropertyName = "HTotalGrade";

                        dataGridView1.Columns[28].DataPropertyName = "HealthGPA";
                        dataGridView1.Columns[29].DataPropertyName = "MThGrade";

                        dataGridView1.Columns[30].DataPropertyName = "MPrGrade";

                        dataGridView1.Columns[31].DataPropertyName = "MTotalGrade";

                        dataGridView1.Columns[32].DataPropertyName = "MoralGPA";
                        dataGridView1.Columns[33].DataPropertyName = "BThGrade";

                        dataGridView1.Columns[34].DataPropertyName = "BPrGrade";


                        dataGridView1.Columns[35].DataPropertyName = "BTotalGrade";

                        dataGridView1.Columns[36].DataPropertyName = "BusinessGPA";
                        dataGridView1.Columns[37].DataPropertyName = "LThGrade";

                        dataGridView1.Columns[38].DataPropertyName = "LPrGrade";

                        dataGridView1.Columns[39].DataPropertyName = "LTotalGrade";

                        dataGridView1.Columns[40].DataPropertyName = "LocalGPA";
                        dataGridView1.Columns[41].DataPropertyName = "TotalThGrade";

                        dataGridView1.Columns[42].DataPropertyName = "TotalPrGrade";

                        dataGridView1.Columns[43].DataPropertyName = "Grade";

                        dataGridView1.Columns[44].DataPropertyName = "GPA";
                        dataGridView1.Columns[45].DataPropertyName = "Remark";

                        dataGridView1.DataSource = ds.Tables["dataTable"];
                        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        dataGridView1.MultiSelect = false;
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

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




            r1.X += 1;
            r1.Y += 1;
            r1.Width = r1.Width + w1 + w2 + w3 - 2;
            r1.Height = r1.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r1);
            StringFormat forma = new StringFormat();
            forma.Alignment = StringAlignment.Center;
            forma.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Nepali", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r1, forma);

            Rectangle r2 = dataGridView1.GetCellDisplayRectangle(11, -1, true);
            int w7 = dataGridView1.GetCellDisplayRectangle(12, -1, true).Width;
            int w8 = dataGridView1.GetCellDisplayRectangle(13, -1, true).Width;
            int w9 = dataGridView1.GetCellDisplayRectangle(14, -1, true).Width;




            r2.X += 1;
            r2.Y += 1;
            r2.Width = r2.Width + w7 + w8 + w9 - 2;
            r2.Height = r2.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r2);
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Center;
            format.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("English", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r2, format);


            Rectangle r3 = dataGridView1.GetCellDisplayRectangle(15, -1, true);
            int w13 = dataGridView1.GetCellDisplayRectangle(16, -1, true).Width;





            r3.X += 1;
            r3.Y += 1;
            r3.Width = r3.Width + w13 - 2;
            r3.Height = r3.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r3);
            StringFormat format1 = new StringFormat();
            format1.Alignment = StringAlignment.Center;
            format1.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Mathematics", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r3, format1);


            Rectangle r4 = dataGridView1.GetCellDisplayRectangle(17, -1, true);
            int w15 = dataGridView1.GetCellDisplayRectangle(18, -1, true).Width;
            int w16 = dataGridView1.GetCellDisplayRectangle(19, -1, true).Width;
            int w17 = dataGridView1.GetCellDisplayRectangle(20, -1, true).Width;




            r4.X += 1;
            r4.Y += 1;
            r4.Width = r4.Width + w15 + w16 + w17 - 2;
            r4.Height = r4.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r4);
            StringFormat format2 = new StringFormat();
            format2.Alignment = StringAlignment.Center;
            format2.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Social", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r4, format2);











            Rectangle r5 = dataGridView1.GetCellDisplayRectangle(21, -1, true);
            int w21 = dataGridView1.GetCellDisplayRectangle(22, -1, true).Width;
            int w22 = dataGridView1.GetCellDisplayRectangle(23, -1, true).Width;
            int w23 = dataGridView1.GetCellDisplayRectangle(24, -1, true).Width;




            r5.X += 1;
            r5.Y += 1;
            r5.Width = r5.Width + w21 + w22 + w23 - 2;
            r5.Height = r5.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r5);
            StringFormat format3 = new StringFormat();
            format3.Alignment = StringAlignment.Center;
            format3.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Science", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r5, format2);

            Rectangle r6 = dataGridView1.GetCellDisplayRectangle(25, -1, true);
            int w27 = dataGridView1.GetCellDisplayRectangle(26, -1, true).Width;
            int w28 = dataGridView1.GetCellDisplayRectangle(27, -1, true).Width;
            int w29 = dataGridView1.GetCellDisplayRectangle(28, -1, true).Width;




            r6.X += 1;
            r6.Y += 1;
            r6.Width = r6.Width + w27 + w28 + w29 - 2;
            r6.Height = r6.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r6);
            StringFormat format4 = new StringFormat();
            format4.Alignment = StringAlignment.Center;
            format4.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Health & Physical Education", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r6, format4);

            Rectangle r7 = dataGridView1.GetCellDisplayRectangle(29, -1, true);
            int w33 = dataGridView1.GetCellDisplayRectangle(30, -1, true).Width;
            int w34 = dataGridView1.GetCellDisplayRectangle(31, -1, true).Width;
            int w35 = dataGridView1.GetCellDisplayRectangle(32, -1, true).Width;




            r7.X += 1;
            r7.Y += 1;
            r7.Width = r7.Width + w33 + w34 + w35 - 2;
            r7.Height = r7.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r7);
            StringFormat format5 = new StringFormat();
            format5.Alignment = StringAlignment.Center;
            format5.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Moral Education", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r7, format5);

            Rectangle r8 = dataGridView1.GetCellDisplayRectangle(33, -1, true);
            int w39 = dataGridView1.GetCellDisplayRectangle(34, -1, true).Width;
            int w40 = dataGridView1.GetCellDisplayRectangle(35, -1, true).Width;
            int w41 = dataGridView1.GetCellDisplayRectangle(36, -1, true).Width;




            r8.X += 1;
            r8.Y += 1;
            r8.Width = r8.Width + w39 + w40 + w41 - 2;
            r8.Height = r8.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor), r8);
            StringFormat format6 = new StringFormat();
            format6.Alignment = StringAlignment.Center;
            format6.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Occuption,business and Technology", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r8, format6);

            Rectangle r9 = dataGridView1.GetCellDisplayRectangle(37, -1, true);
            int w45 = dataGridView1.GetCellDisplayRectangle(38, -1, true).Width;
            int w46 = dataGridView1.GetCellDisplayRectangle(39, -1, true).Width;
            int w47 = dataGridView1.GetCellDisplayRectangle(40, -1, true).Width;




            r9.X += 1;
            r9.Y += 1;
            r9.Width = r9.Width + w45 + w46 + w47 - 2;
            r9.Height = r9.Height / 2 - 2;
            e.Graphics.FillRectangle(new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Blue), r9);
            StringFormat format7 = new StringFormat();
            format7.Alignment = StringAlignment.Center;
            format7.LineAlignment = StringAlignment.Center;
            e.Graphics.DrawString("Computer", dataGridView1.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor), r9, format7);




            Rectangle r10 = dataGridView1.GetCellDisplayRectangle(41, -1, true);
            int w51 = dataGridView1.GetCellDisplayRectangle(42, -1, true).Width;
            int w52 = dataGridView1.GetCellDisplayRectangle(43, -1, true).Width;
            int w53 = dataGridView1.GetCellDisplayRectangle(44, -1, true).Width;



            r10.X += 1;
            r10.Y += 1;
            r10.Width = r10.Width + w51 + w52 + w53 - 2;
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





        private void PrintDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                Image img = Resources.nepal;
                Pen blackPen = new Pen(Color.Black, 3);

                e.Graphics.DrawImage(img, 10, 10, img.Height, img.Width);
                e.Graphics.DrawString("Kalika Rural Municipality", new Font("Arial", 16, FontStyle.Bold), Brushes.Blue, new Point(210, 45));
                e.Graphics.DrawString("Office of Rural Municipal Executive", new Font("Arial", 16, FontStyle.Bold), Brushes.BlueViolet, new Point(185, 75));
                e.Graphics.DrawString("Bacic Education Examination Committee", new Font("Arial", 16, FontStyle.Bold), Brushes.Blue, new Point(177, 110));
                e.Graphics.DrawString("Dhaibung Rasuwa", new Font("Arial", 16, FontStyle.Bold), Brushes.Blue, new Point(265, 140));
                e.Graphics.DrawString("BASIC EDUCATION CERTIFICATE(Grad-8) EXAMINATION " + toolStripTextBox1, new Font("Arial", 16, FontStyle.Bold), Brushes.Blue, new Point(70, 170));
                e.Graphics.DrawString("Grade Sheet", new Font("Arial", 24, FontStyle.Bold), Brushes.Blue, new Point(270, 200));
                e.Graphics.DrawString(" THE GRADE(S) Secured by ", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(10, 275));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[1].Value.ToString(), new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(290, 275));

                e.Graphics.DrawString(" Son/Daughter Of Mr. ", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(620, 275));
                e.Graphics.DrawString(" " + dataGridView1.Rows[SelectedRowIndex].Cells[4].Value.ToString(), new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(10, 315));
                e.Graphics.DrawString(" / Mrs.", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(320, 315));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[3].Value.ToString(), new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(390, 315));
                e.Graphics.DrawString(" Date of Birth  ", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(690, 315));
                e.Graphics.DrawString("" + dataGridView1.Rows[SelectedRowIndex].Cells[5].Value.ToString(), new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(10, 355));
                e.Graphics.DrawString("(In BS)", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(120, 355));
                e.Graphics.DrawString(textBox1.Text + "/" + textBox2.Text + "/" + textBox5.Text, new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(195, 355));
                e.Graphics.DrawString("(In AD)", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(310, 355));
                e.Graphics.DrawString("Address of Student", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(380, 355));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[6].Value.ToString(), new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(560, 355));
                e.Graphics.DrawString("Symbol No", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(10, 400));

                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[2].Value.ToString(), new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(120, 400));
                e.Graphics.DrawString("Shree Kalika Himalaya Secondary School,Kalika-2,", new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(310, 400));
                e.Graphics.DrawString("of", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(285, 400));
                e.Graphics.DrawString("Rasuwa", new Font("Arial", 15, FontStyle.Bold), Brushes.Blue, new Point(10, 435));
                e.Graphics.DrawString("In Basic Education(Grade8) Examination of " + toolStripTextBox1.Text + " BS /" + textBox6.Text + " AD Given Below:", new Font("Arial", 15, FontStyle.Regular), Brushes.Blue, new Point(95, 435));
                e.Graphics.DrawLine(blackPen, 690, 30, 800, 30);
                e.Graphics.DrawLine(blackPen, 690, 30, 690, 150);
                e.Graphics.DrawLine(blackPen, 800, 30, 800, 150);
                e.Graphics.DrawLine(blackPen, 690, 150, 800, 150);
                e.Graphics.DrawLine(blackPen, 277, 310, 620, 310);
                e.Graphics.DrawLine(blackPen, 10, 350, 310, 350);
                e.Graphics.DrawLine(blackPen, 390, 350, 690, 350);
                e.Graphics.DrawLine(blackPen, 10, 390, 120, 390);
                e.Graphics.DrawLine(blackPen, 195, 390, 300, 390);
                //symbol
                e.Graphics.DrawLine(blackPen, 120, 430, 279, 430);
                e.Graphics.DrawLine(blackPen, 310, 430, 811, 430);

                e.Graphics.DrawLine(blackPen, 560, 390, 810, 390);
                e.Graphics.DrawLine(blackPen, 30, 480, 30, 940);
                e.Graphics.DrawLine(blackPen, 30, 480, 800, 480);
                e.Graphics.DrawLine(blackPen, 800, 480, 800, 940);
                e.Graphics.DrawLine(blackPen, 30, 540, 800, 540);
                e.Graphics.DrawLine(blackPen, 30, 940, 800, 940);
                e.Graphics.DrawLine(blackPen, 300, 480, 300, 900);
                e.Graphics.DrawLine(blackPen, 375, 480, 375, 900);
                e.Graphics.DrawLine(blackPen, 450, 520, 450, 900);
                e.Graphics.DrawLine(blackPen, 525, 520, 525, 900);
                e.Graphics.DrawLine(blackPen, 600, 480, 600, 900);
                e.Graphics.DrawLine(blackPen, 675, 480, 675, 900);
                e.Graphics.DrawLine(blackPen, 10, 460, 93, 460);
                e.Graphics.DrawString("1.Th:Theory ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(10, 950));
                e.Graphics.DrawString("2.Pr:Practical ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(10, 965));
                e.Graphics.DrawString("3.Abs:Absent ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(10, 980));
                e.Graphics.DrawLine(blackPen, 60, 1040, 210, 1040);
                e.Graphics.DrawLine(blackPen, 220, 1040, 350, 1040);
                e.Graphics.DrawLine(blackPen, 370, 1040, 530, 1040);
                e.Graphics.DrawLine(blackPen, 540, 1040, 660, 1040);
                e.Graphics.DrawLine(blackPen, 670, 1040, 810, 1040);
                e.Graphics.DrawString("Date of Issue ", new Font("Arial", 14, FontStyle.Bold), Brushes.Blue, new Point(60, 1045));
                e.Graphics.DrawString("Written by ", new Font("Arial", 14, FontStyle.Bold), Brushes.Blue, new Point(230, 1045));
                e.Graphics.DrawString("Head Teacher ", new Font("Arial", 14, FontStyle.Bold), Brushes.Blue, new Point(380, 1045));
                e.Graphics.DrawString("Checked by ", new Font("Arial", 14, FontStyle.Bold), Brushes.Blue, new Point(550, 1045));
                e.Graphics.DrawString("Verified by ", new Font("Arial", 14, FontStyle.Bold), Brushes.Blue, new Point(680, 1045));
                e.Graphics.DrawString("CREDIT ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 490));
                e.Graphics.DrawString(" HOUR", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 515));
                e.Graphics.DrawLine(blackPen, 375, 520, 600, 520);
                e.Graphics.DrawString(" OBTAINED GRADE", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(390, 490));
                e.Graphics.DrawString("GRADE ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(610, 490));
                e.Graphics.DrawString(" POINT", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(610, 515));
                e.Graphics.DrawString(" REMARK", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(685, 515));
                e.Graphics.DrawString(" TH", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(385, 520));
                e.Graphics.DrawString(" PR", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(460, 520));
                e.Graphics.DrawString(" FINAL", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(535, 520));
                e.Graphics.DrawLine(blackPen, 70, 480, 70, 940);
                e.Graphics.DrawString(" S.N.", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 520));
                e.Graphics.DrawString(" Subject", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(100, 515));
                e.Graphics.DrawString("Com.Nepali", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 560));
                e.Graphics.DrawString("Com.English", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 590));
                e.Graphics.DrawString("Mathematics", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 620));
                e.Graphics.DrawString("Science & Environment", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 650));
                e.Graphics.DrawString("Education", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 665));
                e.Graphics.DrawString("Social Studies &  ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 695));
                e.Graphics.DrawString("Population Education  ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 710));
                e.Graphics.DrawString(" Health & Physical ", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 740));
                e.Graphics.DrawString(" Education", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 755));
                e.Graphics.DrawString(" Moral Education", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 787));
                e.Graphics.DrawString(" Occupation,Business &", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 819));

                e.Graphics.DrawString(" Technology Education", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 834));
                e.Graphics.DrawString(" Opt.Computer", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(80, 866));
                e.Graphics.DrawLine(blackPen, 30, 900, 800, 900);
                e.Graphics.DrawString(" GRADE POINT AVERAGE(GPA):" + dataGridView1.Rows[SelectedRowIndex].Cells[44].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(280, 910));
                e.Graphics.DrawString("1", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 560));
                e.Graphics.DrawString("2", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 590));
                e.Graphics.DrawString("3", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 620));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 650));
                e.Graphics.DrawString("5", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 695));
                e.Graphics.DrawString("6", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 740));
                e.Graphics.DrawString("7", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 787));
                e.Graphics.DrawString("8", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 819));
                e.Graphics.DrawString("9", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(32, 866));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 560));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 590));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 620));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 650));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 695));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 740));
                e.Graphics.DrawString("2", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 787));
                e.Graphics.DrawString("2", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 819));
                e.Graphics.DrawString("4", new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(310, 866));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[7].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 560));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[8].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 560));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[9].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 560));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[10].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 560));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[11].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 590));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[12].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 590));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[13].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 590));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[14].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 590));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[15].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 620));

                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[15].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 620));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[16].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 620));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[21].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 650));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[22].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 650));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[23].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 650));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[24].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 650));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[17].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 695));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[18].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 695));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[19].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 695));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[20].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 695));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[25].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 740));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[26].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 740));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[27].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 740));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[28].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 740));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[39].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 787));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[30].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 787));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[31].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 787));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[32].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 787));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[33].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 819));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[34].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 819));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[35].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 819));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[36].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 819));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[37].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(405, 866));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[38].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(465, 866));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[39].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(540, 866));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[40].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(615, 866));
                e.Graphics.DrawString(dataGridView1.Rows[SelectedRowIndex].Cells[45].Value.ToString(), new Font("Arial", 12, FontStyle.Bold), Brushes.Blue, new Point(680, 910));
        }
            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
}

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();

        }
        public class mystruct
        {

            public int year;
            public int month;
            public int day;

        };



        private void ToolStripButton1_Click_1(object sender, EventArgs e)
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
                if (toolStripTextBox1.Text == "")
                {
                    MessageBox.Show("Please fill Year field");
                }
                else
                {
                    int yearValue = 0;
                    yearValue = int.Parse(toolStripTextBox1.Text.ToString());
                    mystruct mm = new mystruct();
                    date ss = new date();


                    mm = ss.cnvToEnglish(12, 26, yearValue);
                    textBox6.Text = mm.year.ToString();
                    textBox4.Text = dataGridView1.Rows[SelectedRowIndex].Cells[5].Value.ToString();



                    date s = new date();
                    var num1 = int.Parse(textBox4.Text.IndexOf("/").ToString());
                    var num = int.Parse(textBox4.Text.LastIndexOf("/").ToString());
                    int num3 = int.Parse(textBox4.Text.Length.ToString());

                    int num4 = num3 - num;
                    int y = 0, mo = 0, d = 0;

                    var t = num - num1;
                    string to = "";
                    if (t == 2)
                    {
                        to = int.Parse(textBox4.Text.Substring(5, 1)).ToString();
                        mo = int.Parse(to);
                    }
                    else if (t == 3)
                    {
                        to = int.Parse(textBox4.Text.Substring(5, 2)).ToString();
                        mo = int.Parse(to);
                    }
                    string too = textBox4.Text.Substring(0, 4).ToString();

                    y = int.Parse(too);


                    string tooo = textBox4.Text.Substring(num + 1, num4 - 1).ToString();

                    d = int.Parse(tooo.ToString());

                    mystruct m = new mystruct();



                    m = s.cnvToEnglish(mo, d, y);


                    textBox1.Text = m.year.ToString();
                    textBox2.Text = m.month.ToString();
                    textBox5.Text = m.day.ToString();
                    printPreviewDialog1.Document = printDocument1;
                    printPreviewDialog1.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    SelectedRowIndex = e.RowIndex;
                }
                else
                {
                    MessageBox.Show("Please celect valid row");
                }
        }
            catch (Exception ex)
            {
                MessageBox.Show("error\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
   
        }
    }
}
