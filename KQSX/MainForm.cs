using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Graph = System.Windows.Forms.DataVisualization.Charting;

namespace KQSX
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void CheckBtn_Click(object sender, EventArgs e)
        {
            string startDate = dateTimePicker1.Value.ToString("yyyyMMdd");
            string endDate = dateTimePicker2.Value.ToString("yyyyMMdd");
            DataTable dt = ReadData(startDate, endDate);

            DataTable dtInfos = new DataTable();
            dtInfos.Columns.Add("HaiSoCuoi");
            dtInfos.Columns.Add("ChuKyLonNhat", typeof(int));
            dtInfos.Columns.Add("ChuKyNhoNhat", typeof(int));
            dtInfos.Columns.Add("ChuKyTrungBinh", typeof(decimal));
            dtInfos.Columns.Add("XuatHienLanCuoi");
            dtInfos.Columns.Add("SoNgayChuaXuatHien", typeof(int));
            dtInfos.Columns.Add("SoNgayConLai", typeof(int));
            dtInfos.Columns.Add("ChuKys");
         //   dtInfos.Columns.Add("SoNgayConLaiTrungBinh");
            dtInfos.Columns.Add("DuDoanChuKy");
            dtInfos.Columns.Add("OK");

            int MaxNum = int.Parse(dt.Rows[0]["row"].ToString());
            Dictionary<string, int> infos = new Dictionary<string, int>();
            for (int i = 0; i < 100; i++)
            {
                DataRow dr = dtInfos.NewRow();
                string str = i.ToString();
                if (i < 10)
                {
                    str = "0" + str;
                }

                dr["HaiSoCuoi"] = str;

                string filter = string.Format("KQ like '%{0}%'", str);
                dt.DefaultView.RowFilter = filter;
                DataTable dtNew = dt.DefaultView.ToTable();
                int temp = 0;
                int period = 0;
                int tempPeriod = 0;
                int xuatHienLanCuoi = int.Parse(dtNew.Rows[0]["row"].ToString());
                int soNgayChuaXuatHien = MaxNum - xuatHienLanCuoi;
                string strXuatHienLanCuoi = dtNew.Rows[0]["ngaynhap"].ToString();
                int total = 0;
                List<int> cacChuKy = new List<int>();
                List<int> cacChuKyTiep = new List<int>();
                Dictionary<int, int> chukylaplai = new Dictionary<int, int>();
                int lastPeriod = 0;
                for (int j = dtNew.Rows.Count - 1; j >= 0; j--)
                {
                    int next = int.Parse(dtNew.Rows[j]["row"].ToString());

                    if (j < dtNew.Rows.Count - 1)
                    {
                        tempPeriod = next - temp;
                        total += tempPeriod;
                        cacChuKy.Add(tempPeriod);
                        if (tempPeriod > soNgayChuaXuatHien)
                        {
                            if (chukylaplai.ContainsKey(tempPeriod))
                            {
                                chukylaplai[tempPeriod] = chukylaplai[tempPeriod] + 1;
                            }
                            else chukylaplai[tempPeriod] = 1;
                        }
                    }

                    if (period < tempPeriod) period = tempPeriod;
                    if (j == 0) lastPeriod = tempPeriod;
                    temp = next;
                }

                for (int k = 0; k < cacChuKy.Count; k++) {
                    if (cacChuKy[k] == lastPeriod && k < cacChuKy.Count - 1 && cacChuKy[k + 1] > soNgayChuaXuatHien) //&& !cacChuKyTiep.Contains(cacChuKy[k + 1]))
                    {
                        cacChuKyTiep.Add(cacChuKy[k + 1]);
                    }
                }

                decimal chukytb = 0;
                int tongchukytiep = 0;
                foreach (int ck in cacChuKyTiep) {
                    tongchukytiep += ck;
                }
                //chukytb = decimal.Round(decimal.Divide(tongchukytiep, cacChuKyTiep.Count>0?cacChuKyTiep.Count:1),2);
                //decimal average = decimal.Round(decimal.Divide(total, dtNew.Rows.Count), 2);
                chukytb = decimal.Round(decimal.Divide(total, cacChuKy.Count), 2);
                dr["ChuKyLonNhat"] = cacChuKy.Max();
                dr["ChuKyNhoNhat"] = cacChuKy.Min();
                dr["ChuKyTrungBinh"] = chukytb;
                dr["XuatHienLanCuoi"] = strXuatHienLanCuoi;
                dr["SoNgayChuaXuatHien"] = MaxNum - xuatHienLanCuoi;
                dr["SoNgayConLai"] = period - (MaxNum - xuatHienLanCuoi);
               // dr["SoNgayConLaiTrungBinh"] = average - (MaxNum - xuatHienLanCuoi);
                dr["Chukys"] = string.Join(",", cacChuKy.ToArray());
              //  int max = chukylaplai.Max(x => x.Value);
                int chuKyDuDoan = (int)Math.Ceiling(chukytb);//chukylaplai.Where(k => k.Value == max).Select(x => x.Key).First();
                dr["DuDoanChuKy"] = chuKyDuDoan + "|" + string.Join(",", cacChuKyTiep.ToArray());
               // dr["SoNgayConLai"] = chuKyDuDoan - (MaxNum - xuatHienLanCuoi);
                dr["OK"] = (cacChuKyTiep.Contains(chuKyDuDoan) ) ? "1" : "0";
                dtInfos.Rows.Add(dr);
            }

            dtInfos.DefaultView.Sort = "OK desc, SoNgayConLai asc, SoNgayChuaXuatHien desc";
            dtInfos.DefaultView.RowFilter = "SoNgayChuaXuatHien>0";
            dataGridView1.DataSource = dtInfos.DefaultView.ToTable();
           // panel1.Visible = false;
            //System.Drawing.Point point = new System.Drawing.Point(10, 0);
            //foreach (DataRowView drv in dtInfos.DefaultView)
            //{
            //    if (point.Y == 0) point = new System.Drawing.Point(10, 10);
            //    else point.Y = point.Y + 1500;
            //    ShowChart(drv, point);
            //}
        }

        private DataTable ReadData(string startDate, string endDate)
        {
            if (string.IsNullOrEmpty(startDate)) startDate = "0";
            if (string.IsNullOrEmpty(endDate)) endDate = "99999";
            DataTable dt = new DataTable();

            string connectionString = @"Provider=Microsoft.JET.OLEDB.4.0;data source=C:\Users\qtruongthanh\AppData\Local\VirtualStore\Program Files\HD Software\HDSCMBv1.02\data.mdb;Jet OLEDB:Database Password=03203471677";

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all rows from the Sheet
                //cmd.CommandText = "SELECT 1 as row, ketqua_ID, Format(ngaynhap, 'yyyymmdd')  as ngaynhap, Format(ngaynhap, 'MM')  as thangnhap, Right( G0,2) + ',' + Right( G1,2) + ',' + Right( G21,2) + ',' + Right( G22,2) + ',' + Right( G31,2) + ',' + Right( G32,2) + ',' + Right( G33,2) + ',' + Right( G34,2) + ',' + Right( G35,2) + ',' + Right( G36,2) + ',' + Right( G41,2) + ',' + Right( G42,2) + ',' + Right( G43,2) + ',' + Right( G44,2) + ',' + Right( G51,2) + ',' + Right( G52,2) + ',' + Right( G53,2) + ',' + Right( G54,2) + ',' + Right( G55,2) + ',' + Right( G56,2) + ',' + Right( G61,2) + ',' + Right( G62,2) + ',' + Right( G63,2) + ',' + Right( G71,2) + ',' + Right( G72,2) + ',' + Right( G73,2) + ',' + Right( G74,2) as KQ FROM [KetQua] where   (Format(ngaynhap, 'yyyymmdd') <= '" + endDate + "') and  (Format(ngaynhap, 'yyyymmdd') >= '" + startDate + "') order by ngaynhap desc";
                cmd.CommandText = "SELECT 1 as row, ketqua_ID, Format(ngaynhap, 'yyyymmdd')  as ngaynhap, Format(ngaynhap, 'MM')  as thangnhap, Right( G0,2) as KQ FROM [KetQua] where   (Format(ngaynhap, 'yyyymmdd') <= '" + endDate + "') and  (Format(ngaynhap, 'yyyymmdd') >= '" + startDate + "') order by ngaynhap desc";
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                da.Fill(dt);

                cmd = null;
                conn.Close();
            }

            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                dt.Rows[i]["row"] = dt.Rows.Count - i;
            }

            return dt;
        }

        private void ShowChart(DataRowView drv, System.Drawing.Point point)
        {
            int MaxX = drv["Chukys"].ToString().Split(',').Length;
            int MaxY = int.Parse(drv["ChuKyLonNhat"].ToString());
            // Create new Graph
            Graph.Chart chart = new Graph.Chart();
            chart.Location = point;
            chart.Size = new System.Drawing.Size(1100, 1500);
            // Add a chartarea called "draw", add axes to it and color the area black
            chart.ChartAreas.Add("draw");
            chart.ChartAreas["draw"].AxisX.Minimum = 0;
            chart.ChartAreas["draw"].AxisX.Maximum = MaxX;
            chart.ChartAreas["draw"].AxisX.Interval = 5;
            chart.ChartAreas["draw"].AxisX.MajorGrid.LineColor = Color.White;
            chart.ChartAreas["draw"].AxisX.MajorGrid.LineDashStyle = Graph.ChartDashStyle.Dash;
            chart.ChartAreas["draw"].AxisY.Minimum = 0;
            chart.ChartAreas["draw"].AxisY.Maximum = MaxY;
            chart.ChartAreas["draw"].AxisY.Interval = 1;
            chart.ChartAreas["draw"].AxisY.MajorGrid.LineColor = Color.White;
            chart.ChartAreas["draw"].AxisY.MajorGrid.LineDashStyle = Graph.ChartDashStyle.Dash;

            chart.ChartAreas["draw"].BackColor = Color.Black;

            // Create a new function series
            chart.Series.Add("MyFunc");
            // Set the type to line      
            chart.Series["MyFunc"].ChartType = Graph.SeriesChartType.Line;
            // Color the line of the graph light green and give it a thickness of 3
            chart.Series["MyFunc"].Color = Color.LightGreen;
            chart.Series["MyFunc"].BorderWidth = 3;
            //This function cannot include zero, and we walk through it in steps of 0.1 to add coordinates to our series
            chart.Series["MyFunc"].Points.DataBindXY(drv["Chukys"].ToString().Split(','), drv["Chukys"].ToString().Split(',').Select(x=>int.Parse(x)).ToArray());

            chart.Series["MyFunc"].LegendText = "Hai so cuoi: " + drv["HaiSoCuoi"].ToString();
            // Create a new legend called "MyLegend".
            chart.Legends.Add("MyLegend");
            chart.Legends["MyLegend"].Position = new Graph.ElementPosition(10,10,10,10) ;// .BorderColor = Color.Tomato; // I like tomato juice!
            panel1.Controls.Add(chart);
            chart.Show();

            	DataTable dt = new DataTable();
                string str = "";
                foreach (DataColumn dc in dt.Columns) {
                    str += dc.ColumnName + ",";
                }

        }
    }
}
