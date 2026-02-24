using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml;
using ClosedXML.Excel;

namespace TaxRefund
{
    public partial class frmTracuu : Form
    {
        SqlDataAdapter da;
        System.Data.DataTable dt;
        SqlCommand cmd;

        public object DataThread { get; private set; }

        private DataGridViewPopupHelper _popupHelper;

        // Add these to your frmTracuu class
        private decimal maxVATValue = 0;
        private const int ChartColumnIndex = 5; // Adjust based on your column order

        private bool isDataGridViewInCustomLayout = false;
        private Panel gridContainer; // Store reference to the container panel       

        private System.Windows.Forms.ScrollBars originalScrollBars;
        private Control originalParent;

        private Size originalDataGridViewSize;
        private DockStyle originalDockStyle;
        private AnchorStyles originalAnchorStyles;

        private int originalHeight;
        private int originalWidth;
        private System.Drawing.Point originalLocation;
        private DockStyle originalDock;
        private AnchorStyles originalAnchor;

        private bool isCustomLayout = false;

        public frmTracuu()
        {
            InitializeComponent();

            //progressDataGridView.SearchCompleted += ProgressDataGridView_SearchCompleted;

            dataGridView1.DataBindingComplete += DataGridView1_DataBindingComplete;

            _popupHelper = new DataGridViewPopupHelper();
            dataGridView1.CellClick += DataGridView1_CellClick;

            // Subscribe to events
            dataGridView1.CellMouseMove += DataGridView1_CellMouseMove;
            dataGridView1.CellFormatting += DataGridView1_CellFormatting;

            // Set default cursor
            dataGridView1.Cursor = Cursors.Default;

            chart1.MouseMove += chart1_MouseMove;
        }

        private void DataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (dataGridView1.Columns.Contains("TrigiaHHchuaVAT") || dataGridView1.Columns.Contains("SotienVATDH") || dataGridView1.Columns.Contains("SotienDVNHH"))
            {
                ApplyDataGridViewStyles(dataGridView1);
                SetupDataGridView();
                SetupDataBindings();
                DisplayTotals(dataGridView1);

                // Defer expensive column sizing so the message pump can run
                DeferOptimizeColumnSizing();
            }
        }
        private void ApplyDataGridViewStyles(DataGridView dataGridView1)
        {
            if (dataGridView1.Rows.Count == 0 || dataGridView1.Columns.Count == 0) return;

            DataGridViewCellStyle numberStyle = new DataGridViewCellStyle
            {
                Format = "N0", // Number format with zero decimal places
                Alignment = DataGridViewContentAlignment.MiddleRight,
                BackColor = Color.LightYellow // Removed optional highlight
            };

            // Define numeric columns with their display names
            var numericColumns = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "TrigiaHHchuaVAT", "Goods Value (ex VAT)" },
                    { "SotienVATDH", "VAT Refund Amount" },
                    { "SotienDVNHH", "Service Fee" }
                };

            foreach (var numericCol in numericColumns)
            {
                if (dataGridView1.Columns.Contains(numericCol.Key))
                {
                    var col = dataGridView1.Columns[numericCol.Key];
                    col.DefaultCellStyle = numberStyle;
                    col.HeaderText = numericCol.Value;
                    col.Width = 150; // Set consistent width for numeric columns
                }
            }
        }
        private decimal SafeParseDecimal(object value)
        {
            decimal result;
            return decimal.TryParse(value?.ToString(), out result) ? result : 0;
        }
        private void DisplayTotals(System.Windows.Forms.DataGridView dgv)
        {
            try
            {
                // Calculate and display totals with proper null handling
                decimal tongsotienvatdh = 0;
                decimal tongtrigiahhchuavat = 0;
                decimal tongsotiendvnhh = 0;

                tongsotienvatdh = dgv.Rows
                    .Cast<DataGridViewRow>()
                    .Where(row => !row.IsNewRow && row.Cells["SotienVATDH"]?.Value != null)
                    .Sum(row => SafeParseDecimal(row.Cells["SotienVATDH"].Value));
                txtTotalRefundAmt.Text = tongsotienvatdh.ToString("#,##0");

                // Calculate TrigiaHHchuaVAT sum using LINQ
                tongtrigiahhchuavat = dgv.Rows
                    .Cast<DataGridViewRow>()
                    .Where(row => !row.IsNewRow && row.Cells["TrigiaHHchuaVAT"]?.Value != null)
                    .Sum(row =>
                    {
                        decimal value;
                        return decimal.TryParse(row.Cells["TrigiaHHchuaVAT"].Value?.ToString(), out value) ? value : 0;
                    });
                txtTotalValue.Text = tongtrigiahhchuavat.ToString("#,##0");

                // Calculate SotienDVNHH sum using LINQ
                tongsotiendvnhh = dgv.Rows
                    .Cast<DataGridViewRow>()
                    .Where(row => !row.IsNewRow && row.Cells["SotienDVNHH"]?.Value != null)
                    .Sum(row =>
                    {
                        decimal value;
                        return decimal.TryParse(row.Cells["SotienDVNHH"].Value?.ToString(), out value) ? value : 0;
                    });
                txtTotalBankServiceFee.Text = tongsotiendvnhh.ToString("#,##0");

                // Count distinct passengers - FIXED: Get DataTable from DataGridView
                int totalPassenger = 0;
                System.Data.DataTable dt = dgv.DataSource as System.Data.DataTable;
                if (dt != null && dt.Rows.Count > 0 && dt.Columns.Contains("SoHC"))
                {
                    totalPassenger = dt.AsEnumerable()
                                    .Select(r => r.Field<string>("SoHC"))
                                    .Where(x => !string.IsNullOrWhiteSpace(x))
                                    .Distinct()
                                    .Count();
                }
                txtTotalPassenger.Text = totalPassenger.ToString("#,##0");

                // Count total datarows (excluding the new row if present)
                int rowCount = dgv.Rows.Count;
                if (dgv.AllowUserToAddRows)
                {
                    rowCount = Math.Max(0, rowCount - 1); // Subtract the new row if it exists
                }
                txtTotalrows.Text = rowCount.ToString("#,##0");
            }
            catch (Exception ex)
            {
                // Handle any errors gracefully
                MessageBox.Show($"Error calculating totals: {ex.Message}", "Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Set default values on error
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtTotalrows.Text = "0";
            }
        }

        private void LoadChartDataAsAreaChart()
        {
            chart1.Series.Clear();
            chart1.Annotations.Clear();

            var refundSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Refund Amount (VND)")
            {
                ChartType = SeriesChartType.Area,
                XValueType = ChartValueType.DateTime,
                YValueType = ChartValueType.Double,
                Color = Color.FromArgb(120, 65, 105, 225),
                BorderColor = Color.FromArgb(255, 30, 144, 255),
                BorderWidth = 2,
                LabelFormat = "#,##0",
                LabelForeColor = Color.DarkBlue,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8f, FontStyle.Regular)
            };

            #region
            //var trendSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Trend (7-day avg)")
            //{ ...
            #endregion

            System.Data.DataTable data = GetBC09Data();
            if (data == null) return;

            // Sort data to ensure the trend line calculates correctly
            DataView dv = data.DefaultView;
            dv.Sort = "NgayHT ASC";
            System.Data.DataTable sortedData = dv.ToTable();

            foreach (DataRow row in sortedData.Rows)
            {
                if (row["NgayHT"] != DBNull.Value && row["TOTAL_REFUND"] != DBNull.Value)
                {
                    DateTime dateValue = Convert.ToDateTime(row["NgayHT"]);
                    double refundValue = Convert.ToDouble(row["TOTAL_REFUND"]);

                    // Create the point
                    int pointIndex = refundSeries.Points.AddXY(dateValue, refundValue);
                }
            }

            chart1.Series.Add(refundSeries);

            // Add title with custom font
            chart1.Titles.Clear();
            Title chartTitle = new Title($"Refund Amount by Payment Date ({txtRfdate.Text} to {txtRtdate.Text}) - Total Records: {data.Rows.Count}");
            chartTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10f, FontStyle.Regular);
            chartTitle.ForeColor = Color.DarkBlue;
            chart1.Titles.Add(chartTitle);

            // Chart Area Styling
            var chartArea = chart1.ChartAreas[0];

            chartArea.BackColor = Color.FromArgb(245, 247, 250);

            // Axis Formatting
            chartArea.AxisY.LabelStyle.Format = "#,##0"; // Thousands separator for Y Axis
            chartArea.AxisX.LabelStyle.Format = "dd/MM"; // Clean date format
            chartArea.AxisX.IntervalType = DateTimeIntervalType.Days;

            // Gradient effect
            refundSeries.BackGradientStyle = GradientStyle.TopBottom;
            refundSeries.BackSecondaryColor = Color.FromArgb(50, 135, 206, 235);

            // Max Annotation
            if (refundSeries.Points.Count > 0)
            {
                var maxPoint = refundSeries.Points.FindMaxByValue();
                chart1.Annotations.Add(new CalloutAnnotation
                {
                    Text = $"Peak: {maxPoint.YValues[0]:#,##0}",
                    AnchorDataPoint = maxPoint,
                    AnchorOffsetY = -10,
                    BackColor = Color.LemonChiffon
                });
            }

            // Configure Vertical Line (X-Axis)
            chartArea.CursorX.IsUserEnabled = true;
            chartArea.CursorX.IsUserSelectionEnabled = false;
            chartArea.CursorX.LineColor = Color.FromArgb(100, Color.RoyalBlue); // Semi-transparent
            chartArea.CursorX.LineDashStyle = ChartDashStyle.Dash;
            chartArea.CursorX.LineWidth = 1;

            // Configure Horizontal Line (Y-Axis) - Optional
            chartArea.CursorY.IsUserEnabled = true;
            chartArea.CursorY.LineColor = Color.FromArgb(100, Color.RoyalBlue);
            chartArea.CursorY.LineDashStyle = ChartDashStyle.Dash;

            // Create a hidden "Blue Tooltip"
            var customTooltip = new CalloutAnnotation
            {
                Name = "BlueTip",
                Visible = false,
                BackColor = Color.RoyalBlue,
                ForeColor = Color.White,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8f, FontStyle.Regular),
                CalloutStyle = CalloutStyle.RoundedRectangle,
                AnchorAlignment = ContentAlignment.TopCenter
            };
            chart1.Annotations.Add(customTooltip);

            ConfigureCommonChartSettings();
        }
        private void LoadChartData()
        {
            chart1.Series.Clear();

            // Create series for refund amounts
            System.Windows.Forms.DataVisualization.Charting.Series refundSeries =
                new System.Windows.Forms.DataVisualization.Charting.Series("Total Refund Amount (VND)");

            refundSeries.ChartType = SeriesChartType.Line;
            refundSeries.XValueType = ChartValueType.DateTime;
            refundSeries.YValueType = ChartValueType.Double;
            refundSeries.IsValueShownAsLabel = true;
            refundSeries.LabelFormat = "#,##0";
            refundSeries.BorderWidth = 2;
            refundSeries.LabelForeColor = Color.DarkOrange;
            refundSeries.Color = Color.DarkOrange;

            // Marker settings
            refundSeries.MarkerStyle = MarkerStyle.Circle;
            refundSeries.MarkerSize = 8;
            refundSeries.MarkerColor = Color.DarkOrange;
            refundSeries.MarkerBorderColor = Color.White;

            // Get chart area reference
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea = chart1.ChartAreas[0];

            // BLUR/SOFTEN BACKGROUND SETTINGS
            // ================================

            // 1. Soft gradient background for entire chart area
            chartArea.BackColor = Color.WhiteSmoke;
            chartArea.BackGradientStyle = GradientStyle.TopBottom;
            chartArea.BackSecondaryColor = Color.Lavender;

            // 2. Soften plot area specifically
            chartArea.AxisX.LineColor = Color.FromArgb(100, Color.Gray); // Semi-transparent axis lines
            chartArea.AxisY.LineColor = Color.FromArgb(100, Color.Gray);

            // 3. Soft gridlines
            chartArea.AxisX.MajorGrid.LineColor = Color.FromArgb(30, Color.Gray);
            chartArea.AxisY.MajorGrid.LineColor = Color.FromArgb(30, Color.Gray);
            chartArea.AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            chartArea.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            // 4. Soft shadow effect around plot area
            chartArea.ShadowColor = Color.FromArgb(40, 100, 100, 100);
            chartArea.ShadowOffset = 3;

            // 5. Optional: Add a subtle border with shadow
            chartArea.BorderColor = Color.FromArgb(150, Color.Silver);
            chartArea.BorderWidth = 1;
            chartArea.BorderDashStyle = ChartDashStyle.Solid;

            // Continue with your existing data loading...
            refundSeries.Font = new System.Drawing.Font("Microsoft Sans Serif", 8f, FontStyle.Regular);
            refundSeries["BarLabelStyle"] = "Center";
            refundSeries["PixelPointWidth"] = "true";
            refundSeries.SmartLabelStyle.Enabled = true;
            refundSeries.SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Yes;

            // Rotate labels if needed
            chartArea.AxisX.LabelStyle.Angle = -30;

            // Get data from database
            System.Data.DataTable data = GetBC09Data();

            // Add data points to series
            foreach (DataRow row in data.Rows)
            {
                if (row["NgayHT"] != DBNull.Value && row["TOTAL_REFUND"] != DBNull.Value)
                {
                    DateTime ngayHT = (DateTime)row["NgayHT"];
                    decimal totalRefund = (decimal)row["TOTAL_REFUND"];
                    refundSeries.Points.AddXY(ngayHT, totalRefund);
                }
            }

            chart1.Series.Add(refundSeries);

            // Configure axis fonts with soft colors
            chartArea.AxisX.LabelStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10f);
            chartArea.AxisY.LabelStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10f);
            chartArea.AxisY.LabelStyle.Format = "#,##0";
            chartArea.AxisX.LabelStyle.ForeColor = Color.FromArgb(150, 0, 0, 0); // Semi-black
            chartArea.AxisY.LabelStyle.ForeColor = Color.FromArgb(150, 0, 0, 0);

            chartArea.AxisX.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 10f, FontStyle.Bold);
            chartArea.AxisY.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

            // Adjust view if we have data
            if (refundSeries.Points.Count > 0)
            {
                DateTime startDate = DateTime.Parse(txtRfdate.Text);
                DateTime endDate = DateTime.Parse(txtRtdate.Text);

                chartArea.AxisX.Minimum = startDate.ToOADate();
                chartArea.AxisX.Maximum = endDate.ToOADate();
                chartArea.AxisX.Interval = CalculateOptimalInterval(refundSeries.Points);
            }

            // Add title with custom font
            chart1.Titles.Clear();
            Title chartTitle = new Title($"Refund Amount by Payment Date ({txtRfdate.Text} to {txtRtdate.Text}) - Total Records: {data.Rows.Count}");
            chartTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10f, FontStyle.Regular);
            chartTitle.ForeColor = Color.DarkBlue;
            chart1.Titles.Add(chartTitle);

            // Configure legend with soft background
            if (chart1.Legends.Count == 0)
            {
                chart1.Legends.Add(new System.Windows.Forms.DataVisualization.Charting.Legend());
            }
            chart1.Legends[0].Font = new System.Drawing.Font("Microsoft Sans Serif", 10f, FontStyle.Regular);
            chart1.Legends[0].BackColor = Color.FromArgb(220, 240, 255); // Soft blue background

            // Optional: Add overall chart background image or color
            chart1.BackColor = Color.WhiteSmoke;
            chart1.BackGradientStyle = GradientStyle.TopBottom;
            chart1.BackSecondaryColor = Color.Lavender;
        }
        private void LoadChartDataAsStackedColumn()
        {
            chart1.Series.Clear();

            // Create multiple series for different categories if you have them
            var vatSeries = new System.Windows.Forms.DataVisualization.Charting.Series("VAT Refund")
            {
                ChartType = SeriesChartType.StackedColumn,
                Color = Color.FromArgb(255, 70, 130, 180), // Steel Blue
                IsValueShownAsLabel = true,
                LabelFormat = "#,##0"
            };

            var serviceFeeSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Service Fee")
            {
                ChartType = SeriesChartType.StackedColumn,
                Color = Color.FromArgb(255, 100, 149, 237), // Cornflower Blue
                IsValueShownAsLabel = true,
                LabelFormat = "#,##0"
            };

            var goodsValueSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Goods Value")
            {
                ChartType = SeriesChartType.StackedColumn,
                Color = Color.FromArgb(255, 30, 144, 255), // Dodger Blue
                IsValueShownAsLabel = true,
                LabelFormat = "#,##0"
            };

            // Get data - assuming your data has multiple metrics
            var data = GetBC09Data();

            // Group by date (monthly, weekly, or daily depending on data range)
            var groupedData = data.AsEnumerable()
                .GroupBy(row => ((DateTime)row["NgayHT"]).Date)
                .OrderBy(g => g.Key);

            foreach (var group in groupedData)
            {
                var date = group.Key;
                decimal totalVAT = group.Sum(r => r.Field<decimal?>("TOTAL_REFUND") ?? 0);
                //decimal totalServiceFee = group.Sum(r => r.Field<decimal?>("TongSotienDVNHH") ?? 0);
                //decimal totalGoodsValue = group.Sum(r => r.Field<decimal?>("TongTrigiaHHchuaVAT") ?? 0);

                vatSeries.Points.AddXY(date, totalVAT);
                //serviceFeeSeries.Points.AddXY(date, totalServiceFee);
                //goodsValueSeries.Points.AddXY(date, totalGoodsValue);
            }

            chart1.Series.Add(vatSeries);
            chart1.Series.Add(serviceFeeSeries);
            chart1.Series.Add(goodsValueSeries);

            // Configure chart area for better readability
            var chartArea = chart1.ChartAreas[0];
            chartArea.AxisX.IntervalType = DateTimeIntervalType.Days;
            chartArea.AxisX.LabelStyle.Format = "MMM dd";
            chartArea.AxisY.LabelStyle.Format = "#,##0";

            // Add data labels with background
            foreach (System.Windows.Forms.DataVisualization.Charting.Series series in chart1.Series)
            {
                foreach (DataPoint point in series.Points)
                {
                    point.LabelBackColor = Color.FromArgb(200, 255, 255, 255);
                    point.LabelBorderColor = Color.LightGray;
                    point.LabelBorderWidth = 1;
                }
            }

            ConfigureCommonChartSettings();
        }
        private void LoadChartDataAsSplineChart()
        {
            chart1.Series.Clear();

            var refundSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Refund Amount")
            {
                ChartType = SeriesChartType.Spline,
                XValueType = ChartValueType.DateTime,
                YValueType = ChartValueType.Double,
                Color = Color.FromArgb(255, 46, 139, 87), // Sea Green
                BorderWidth = 4,
                ShadowOffset = 2,
                ShadowColor = Color.FromArgb(100, 0, 0, 0),
                MarkerStyle = MarkerStyle.Circle,
                MarkerSize = 10,
                MarkerColor = Color.White,
                MarkerBorderColor = Color.FromArgb(255, 46, 139, 87),
                MarkerBorderWidth = 2
            };

            var data = GetBC09Data();
            var points = new List<Tuple<DateTime, decimal>>();

            foreach (DataRow row in data.Rows)
            {
                if (row["NgayHT"] != DBNull.Value && row["TOTAL_REFUND"] != DBNull.Value)
                {
                    DateTime ngayHT = (DateTime)row["NgayHT"];
                    decimal totalRefund = (decimal)row["TOTAL_REFUND"];
                    refundSeries.Points.AddXY(ngayHT, totalRefund);
                    points.Add(new Tuple<DateTime, decimal>(ngayHT, totalRefund));
                }
            }

            chart1.Series.Add(refundSeries);

            // Add annotations for important points
            if (points.Count > 0)
            {
                var maxPoint = points.OrderByDescending(p => p.Item2).First();
                var minPoint = points.OrderBy(p => p.Item2).First();
                var avgValue = points.Average(p => p.Item2);

                // Max value annotation
                var maxAnnotation = new CalloutAnnotation
                {
                    Text = $"Peak: {maxPoint.Item2:#,##0}",
                    AnchorDataPoint = refundSeries.Points.First(p => p.XValue == maxPoint.Item1.ToOADate()),
                    AnchorOffsetY = 20,
                    Font = new System.Drawing.Font("Arial", 10, FontStyle.Bold),
                    ForeColor = Color.DarkGreen,
                    BackColor = Color.FromArgb(220, 255, 240)
                };

                // Average line
                var avgLine = new HorizontalLineAnnotation
                {
                    AxisX = chart1.ChartAreas[0].AxisX,
                    AxisY = chart1.ChartAreas[0].AxisY,
                    Y = (double)avgValue,
                    LineColor = Color.Red,
                    LineWidth = 2,
                    LineDashStyle = ChartDashStyle.Dash,
                    ClipToChartArea = chart1.ChartAreas[0].Name
                };

                var avgAnnotation = new TextAnnotation
                {
                    Text = $"Avg: {avgValue:#,##0}",
                    X = 95, // Position as percentage
                    Y = (double)avgValue,
                    Font = new System.Drawing.Font("Arial", 9, FontStyle.Italic),
                    ForeColor = Color.Red
                };

                chart1.Annotations.Add(maxAnnotation);
                chart1.Annotations.Add(avgLine);
                chart1.Annotations.Add(avgAnnotation);
            }

            ConfigureCommonChartSettings();

            // Customize for spline chart
            var chartArea = chart1.ChartAreas[0];
            chartArea.BackColor = Color.White;
            chartArea.BackGradientStyle = GradientStyle.DiagonalRight;
            chartArea.BackSecondaryColor = Color.FromArgb(240, 248, 255);
        }
        private void LoadChartDataAsComboChart()
        {
            chart1.Series.Clear();

            // Column series for refund amounts
            var columnSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Daily Refund")
            {
                ChartType = SeriesChartType.Column,
                XValueType = ChartValueType.DateTime,
                YValueType = ChartValueType.Double,
                Color = Color.FromArgb(220, 100, 149, 237), // Cornflower Blue
                BorderColor = Color.FromArgb(255, 65, 105, 225),
                BorderWidth = 1,
                IsValueShownAsLabel = true,
                LabelFormat = "#,##0",
                LabelAngle = -90,
                LabelForeColor = Color.DarkBlue
            };

            // Line series for cumulative total
            var lineSeries = new System.Windows.Forms.DataVisualization.Charting.Series("Cumulative Total")
            {
                ChartType = SeriesChartType.Line,
                YValueType = ChartValueType.Double,
                Color = Color.FromArgb(255, 255, 69, 0), // Orange Red
                BorderWidth = 3,
                MarkerStyle = MarkerStyle.Diamond,
                MarkerSize = 8,
                MarkerColor = Color.White,
                MarkerBorderColor = Color.FromArgb(255, 255, 69, 0),
                MarkerBorderWidth = 2
            };

            // Get and process data
            var data = GetBC09Data();
            var dailyData = data.AsEnumerable()
                .GroupBy(row => ((DateTime)row["NgayHT"]).Date)
                .Select(g => new
                {
                    Date = g.Key,
                    DailyTotal = g.Sum(r => r.Field<decimal>("TOTAL_REFUND"))
                })
                .OrderBy(x => x.Date)
                .ToList();

            decimal runningTotal = 0;
            foreach (var day in dailyData)
            {
                columnSeries.Points.AddXY(day.Date, day.DailyTotal);
                runningTotal += day.DailyTotal;
                lineSeries.Points.AddXY(day.Date, runningTotal);
            }

            // Add to chart
            chart1.Series.Add(columnSeries);
            chart1.Series.Add(lineSeries);

            // Configure secondary axis for line series
            lineSeries.YAxisType = AxisType.Secondary;
            var chartArea = chart1.ChartAreas[0];
            chartArea.AxisY2.Enabled = AxisEnabled.True;
            chartArea.AxisY2.LabelStyle.Format = "#,##0";
            chartArea.AxisY2.Title = "Cumulative Total (VND)";
            chartArea.AxisY.Title = "Daily Refund (VND)";

            // Different gridlines for each axis
            chartArea.AxisY.MajorGrid.LineColor = Color.FromArgb(50, Color.Blue);
            chartArea.AxisY2.MajorGrid.LineColor = Color.FromArgb(50, Color.Red);

            ConfigureCommonChartSettings();
        }
        private void ConfigureCommonChartSettings()
        {

            var chartArea = chart1.ChartAreas[0];

            // Professional styling
            chartArea.BackColor = Color.FromArgb(248, 249, 250);
            chartArea.ShadowColor = Color.FromArgb(40, 0, 0, 0);
            chartArea.ShadowOffset = 3;

            // Axis styling
            chartArea.AxisX.LineColor = Color.FromArgb(150, Color.Gray);
            chartArea.AxisY.LineColor = Color.FromArgb(150, Color.Gray);
            chartArea.AxisX.MajorGrid.LineColor = Color.FromArgb(30, Color.Gray);
            chartArea.AxisY.MajorGrid.LineColor = Color.FromArgb(30, Color.Gray);

            // Chart background
            chart1.BackColor = Color.White;
            chart1.BorderlineColor = Color.LightGray;
            chart1.BorderlineWidth = 1;
            chart1.BorderlineDashStyle = ChartDashStyle.Solid;

            // Legend
            if (chart1.Legends.Count > 0)
            {
                chart1.Legends[0].BackColor = Color.FromArgb(240, 248, 255);
                chart1.Legends[0].BorderColor = Color.LightGray;
                chart1.Legends[0].BorderWidth = 1;
                chart1.Legends[0].Font = new System.Drawing.Font("Segoe UI", 9);
            }

            // Title
            if (chart1.Titles.Count == 0)
            {
                chart1.Titles.Add(new Title());
            }
            chart1.Titles[0].Font = new System.Drawing.Font("Segoe UI", 12, FontStyle.Bold);
            chart1.Titles[0].ForeColor = Color.FromArgb(50, 50, 100);
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            // 1. Hit Test to find the data point
            var hit = chart1.HitTest(e.X, e.Y);
            var tip = chart1.Annotations["BlueTip"] as CalloutAnnotation;
            var area = chart1.ChartAreas[0];

            if (hit.ChartElementType == ChartElementType.DataPoint)
            {
                var point = hit.Series.Points[hit.PointIndex];

                // 2. Position the Vertical Crosshair line exactly on the data point
                area.CursorX.SetCursorPosition(point.XValue);
                area.CursorY.SetCursorPosition(point.YValues[0]);

                // 3. Update and show the Blue Pop-up
                tip.Text = $"Date: {DateTime.FromOADate(point.XValue):dd/MM/yyyy}\nRefund: {point.YValues[0]:#,##0} VND";
                tip.AnchorDataPoint = point;
                tip.Visible = true;
                chart1.Cursor = Cursors.Cross;
            }
            else
            {
                // Hide everything when not hovering over data
                tip.Visible = false;
                area.CursorX.SetCursorPosition(double.NaN); // Hide line
                area.CursorY.SetCursorPosition(double.NaN); // Hide line
                chart1.Cursor = Cursors.Default;
            }
        }
        //private void chart1_MouseMove(object sender, MouseEventArgs e)
        //{ ... }
        private System.Data.DataTable GetBC09Data()
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();
            string rfDate = txtRfdate.Text.Trim();
            string rtDate = txtRtdate.Text.Trim();

            DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");
            DateTime prtdate = DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var rtdate = prtdate.ToString("yyyy-MM-dd hh:mm:ss tt");

            System.Data.DataTable data = new System.Data.DataTable();

            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                    // Using parameterized query to prevent SQL injection
                    string query = @"SELECT NgayHT, SUM(SotienVATDH) AS TOTAL_REFUND 
                               FROM BC09 
                               WHERE NgayHT BETWEEN @StartDate AND @EndDate 
                               GROUP BY NgayHT
                               ORDER BY NgayHT";

                    SqlCommand command = new SqlCommand(query, conn);
                    command.Parameters.AddWithValue("@StartDate", rfdate);
                    command.Parameters.AddWithValue("@EndDate", rtdate);

                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(data);
                }
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return data;
        }

        private double CalculateOptimalInterval(DataPointCollection points)
        {
            if (points.Count < 2) return 1;

            DateTime firstDate = DateTime.FromOADate(points[0].XValue);
            DateTime lastDate = DateTime.FromOADate(points[points.Count - 1].XValue);

            double totalDays = (lastDate - firstDate).TotalDays;

            // Adjust interval based on total days in range
            if (totalDays <= 7) return 1; // Daily for 1 week
            if (totalDays <= 30) return 8; // Weekly for 1 month
            if (totalDays <= 90) return 15; // Bi-weekly for 3 months
            return 30; // Monthly for longer periods
        }

        // Implemented safe painting for the VAT chart column (removed NotImplementedException)
        private void DataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            try
            {
                // Only paint our chart column and only for data rows (not header)
                if (e.ColumnIndex == ChartColumnIndex && e.RowIndex >= 0)
                {
                    e.PaintBackground(e.CellBounds, true);

                    // Get the VAT value from another column (adjust index/name as needed)
                    var vatCell = dataGridView1.Rows[e.RowIndex].Cells["SotienVATDH"];
                    if (vatCell != null && vatCell.Value != null &&
                        decimal.TryParse(vatCell.Value.ToString(), out decimal value))
                    {
                        float percentage = maxVATValue > 0 ? (float)(value / maxVATValue) : 0f;

                        // Draw chart bar
                        System.Drawing.Rectangle fillRect = new System.Drawing.Rectangle(
                            e.CellBounds.X + 2,
                            e.CellBounds.Y + 5,
                            (int)((e.CellBounds.Width - 4) * percentage),
                            e.CellBounds.Height - 10
                        );

                        using (Brush brush = new SolidBrush(Color.SteelBlue))
                        {
                            e.Graphics.FillRectangle(brush, fillRect);
                        }

                        // Draw value text
                        string valueText = value.ToString("N0");
                        TextRenderer.DrawText(
                            e.Graphics,
                            valueText,
                            dataGridView1.Font,
                            new System.Drawing.Point(e.CellBounds.X + 5, e.CellBounds.Y + 20),
                            Color.Black
                        );
                    }

                    e.Handled = true;
                }
                else
                {
                    // Default painting
                    e.Paint(e.CellBounds, DataGridViewPaintParts.All);
                    e.Handled = true;
                }
            }
            catch
            {
                // If anything goes wrong during custom paint, fallback to default painting to avoid UI thread stalls.
                try
                {
                    e.Paint(e.CellBounds, DataGridViewPaintParts.All);
                }
                catch { /* swallow to avoid throwing during paint */ }
                e.Handled = true;
            }
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //// Only paint our chart column and only for data rows (not header)
            //if (e.ColumnIndex == ChartColumnIndex && e.RowIndex >= 0)
            //{
            //    e.PaintBackground(e.CellBounds, true);
            //    ...
            //    e.Handled = true;
            //}
        }

        private void LoadDataWithChart()
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();
            // Your existing data loading code, but modified to:
            // 1. Find max VAT value
            // 2. Ensure chart column exists

            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
                // Example loading code (adapt to your existing method):
                string query = "SELECT KyhieuHD, SoHD, NgayHD, TenDNBH, SotienVATDH FROM BC09 WHERE NgayHT IS NOT NULL";

                da = new SqlDataAdapter(query, conn);
                dt = new System.Data.DataTable();
                da.Fill(dt);

                // Find max VAT value for scaling
                maxVATValue = dt.AsEnumerable()
                    .Where(row => row["SotienVATDH"] != DBNull.Value)
                    .Max(row => Convert.ToDecimal(row["SotienVATDH"]));

                dataGridView1.DataSource = dt;

                // Ensure chart column exists and is visible
                if (!dataGridView1.Columns.Contains("colVatChart"))
                {
                    DataGridViewColumn chartColumn = new DataGridViewTextBoxColumn();
                    chartColumn.HeaderText = "VAT Chart";
                    chartColumn.Name = "colVatChart";
                    dataGridView1.Columns.Add(chartColumn);
                }

                // Make rows taller to accommodate chart
                dataGridView1.RowTemplate.Height = 60;
            }
            // Close Connection 
            conn.Close();
            conn.Dispose();

            // Defer sizing to avoid blocking UI thread
            DeferOptimizeColumnSizing();
        }

        private void DataGridView1_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var columnName = dataGridView1.Columns[e.ColumnIndex].Name;

                if (columnName.Equals("SoHC", StringComparison.OrdinalIgnoreCase) ||
                    columnName.Equals("SoHD", StringComparison.OrdinalIgnoreCase))
                {
                    dataGridView1.Cursor = Cursors.Hand;
                    return;
                }
            }
            dataGridView1.Cursor = Cursors.Default;
        }

        private void DataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var columnName = dataGridView1.Columns[e.ColumnIndex].Name;

                if (columnName.Equals("SoHC", StringComparison.OrdinalIgnoreCase) ||
                    columnName.Equals("SoHD", StringComparison.OrdinalIgnoreCase))
                {
                    // Style as clickable text
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Blue;
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.Font =
                        new System.Drawing.Font("Microsoft Sans Serif", 12f, FontStyle.Underline);
                }
            }
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ensure we're clicking on a valid cell (not header or outside grid)
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var columnName = dataGridView1.Columns[e.ColumnIndex].Name;
            var row = dataGridView1.Rows[e.RowIndex];

            // Check if clicked column is "SoHC" or "SoHD"
            if (columnName.Equals("SoHC", StringComparison.OrdinalIgnoreCase) ||
                columnName.Equals("SoHD", StringComparison.OrdinalIgnoreCase))
            {
                string value = row.Cells[e.ColumnIndex].Value?.ToString();

                if (!string.IsNullOrEmpty(value))
                {
                    _popupHelper.ShowDetailsPopup(columnName, value, this);
                }
            }
        }

        private void CustomDataThreadCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var dt = dataGridView1.DataSource as System.Data.DataTable;
            if (dt != null && dt.Rows.Count > 0)
            {
                DataGridViewCellStyle style = new DataGridViewCellStyle();
                style.Format = "N0";
                this.dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle = style;
                this.dataGridView1.Columns["SotienVATDH"].DefaultCellStyle = style;
                this.dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle = style;

                this.dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                this.dataGridView1.Columns["SotienVATDH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                this.dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                SetupDataGridView();

                // Calculate and display totals
                decimal tongsotienvatdh = Convert.ToDecimal(dt.Compute("Sum(SotienVATDH)", string.Empty));
                txtTotalRefundAmt.Text = tongsotienvatdh.ToString("#,##0");

                decimal tongtrigiahhchuavat = Convert.ToDecimal(dt.Compute("Sum(TrigiaHHchuaVAT)", string.Empty));
                txtTotalValue.Text = tongtrigiahhchuavat.ToString("#,##0");

                decimal tongsotiendvnhh = Convert.ToDecimal(dt.Compute("Sum(SotienDVNHH)", string.Empty));
                txtTotalBankServiceFee.Text = tongsotiendvnhh.ToString("#,##0");

                int totalPassenger = dt.AsEnumerable()
                                    .Select(r => r.Field<string>("SoHC"))
                                    .Where(x => !string.IsNullOrWhiteSpace(x))
                                    .Distinct()
                                    .Count();
                txtTotalPassenger.Text = totalPassenger.ToString("#,##0");


                // Count total datarows                        
                txtTotalrows.Text = dataGridView1.Rows.Count.ToString("#,##0");

                // Data binding
                SetupDataBindings();

                // Defer sizing after binding
                DeferOptimizeColumnSizing();
            }
            else
            {
                // Clear all fields if no results
                ClearResultFields();
            }
        }

        private void SetupDataBindings()
        {
            txtSoHC.DataBindings.Clear();
            txtSoHC.DataBindings.Add("Text", dataGridView1.DataSource, "SoHC");
            txtSoHD.DataBindings.Clear();
            txtSoHD.DataBindings.Add("Text", dataGridView1.DataSource, "SoHD");
            txtNgayHD.DataBindings.Clear();
            txtNgayHD.DataBindings.Add("Text", dataGridView1.DataSource, "NgayHD");
            txtGoodsValue.DataBindings.Clear();
            txtGoodsValue.DataBindings.Add("Text", dataGridView1.DataSource, "TrigiaHHchuaVAT", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
            txtRefundAmt.DataBindings.Clear();
            txtRefundAmt.DataBindings.Add("Text", dataGridView1.DataSource, "SotienVATDH", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
            txtMasoDN.DataBindings.Clear();
            txtMasoDN.DataBindings.Add("Text", dataGridView1.DataSource, "MasoDN");
            txtTenDNBH.DataBindings.Clear();
            txtTenDNBH.DataBindings.Add("Text", dataGridView1.DataSource, "TenDNBH");
        }

        private void ClearResultFields()
        {
            txtTotalRefundAmt.Text = "0";
            txtTotalValue.Text = "0";
            txtTotalBankServiceFee.Text = "0";
            txtTotalPassenger.Text = "0";
            txtTotalrows.Text = "0";
        }

        private bool ValidInput()
        {
            if (txtFilePath.Text.Trim() == "")
            {
                MessageBox.Show("Please select Excel file to import.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog myOpenFileDialog = new OpenFileDialog();
            myOpenFileDialog.CheckFileExists = true;
            myOpenFileDialog.DefaultExt = ".xls";
            myOpenFileDialog.InitialDirectory = @"D:\VIEC CO QUAN\HOAN THUE\So lieu cua NH\2024";
            myOpenFileDialog.Multiselect = false;

            // Dùng OpenFileDialog và đưa path và tên tập tin vào txtExcelFilePath
            if (myOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.txtFilePath.Text = myOpenFileDialog.FileName;
            }
        }

        private void RestoreDataGridView()
        {
            if (!isDataGridViewInCustomLayout) return;

            // Remove from custom panel
            if (dataGridView1.Parent != null)
            {
                dataGridView1.Parent.Controls.Remove(dataGridView1);
            }

            // Dispose the container
            if (gridContainer != null)
            {
                this.Controls.Remove(gridContainer);
                gridContainer.Dispose();
                gridContainer = null;
            }

            // --- RESTORE ORIGINAL DIMENSIONS ---
            dataGridView1.Dock = originalDock;

            // Only set Location and Size manually if Dock is not Fill
            if (dataGridView1.Dock == DockStyle.None)
            {
                dataGridView1.Location = originalLocation;
                dataGridView1.Size = new System.Drawing.Size(originalWidth, originalHeight);
            }

            dataGridView1.Anchor = originalAnchor;
            dataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Both;

            // Re-add to the form
            this.Controls.Add(dataGridView1);
            dataGridView1.BringToFront();

            isDataGridViewInCustomLayout = false;
        }

        private void SetupDataGridView()
        {
            // Set font for headers           
            var headerFont = new System.Drawing.Font("Microsoft Sans Serif", 11f, FontStyle.Bold);
            dataGridView1.EnableHeadersVisualStyles = false; // Disable system styles
            dataGridView1.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                Font = headerFont,
                ForeColor = Color.DeepSkyBlue, // FromArgb(51, 122, 183),
                BackColor = Color.AntiqueWhite, // FromArgb(51, 122, 183), // Bootstrap primary blue
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Padding = new Padding(3)
            };

            // Optional: Add gradient background
            #region
            //dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            //dataGridView1.ColumnHeadersHeight = 30;
            //dataGridView1.Paint += (sender, e) =>
            //{ ... }
            #endregion

            // Set font for the entire grid
            dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11f, FontStyle.Regular);

            // Set up for the entire grid
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.DefaultCellStyle.BackColor = Color.FloralWhite;
            dataGridView1.DefaultCellStyle.ForeColor = Color.DarkBlue;

            // IMPORTANT: Do not autosize every cell synchronously (costly). Use deferred/limited sizing.
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            dataGridView1.GridColor = Color.Silver;
            dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.MultiSelect = true;

            // Attempt to enable double buffering to reduce painting overhead
            try
            {
                var prop = typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
                prop?.SetValue(dataGridView1, true, null);
            }
            catch
            {
                // ignore if cannot set
            }

            // Define all custom mappings here. Use the actual DataTable column name as the Key.
            var columnHeaderMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "STT", "No." },
                    { "SoHD", "Invoice No." },
                    { "NgayHD", "Invoice Date" },
                    { "KyhieuHD", "Series" },
                    { "HoTenHK", "Passenger Name" },
                    { "SoHC", "Passport No." },
                    { "NgayHC", "Passport Date" },
                    { "Quoctich", "Nationality" },
                    { "NgayHT", "Refund Date" },
                    { "MasoDN", "Business ID" },
                    { "TenDNBH", "Business Name" },
                    { "Ghichu", "Notes" },
                    { "NgaynhapHT", "Data Entry Date" },
                    { "LoginName", "Login ID" }
                };

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                if (columnHeaderMap.TryGetValue(column.Name, out var displayHeader))
                {
                    column.HeaderText = displayHeader;
                }
            }
        }

        // Defer optimization so UI thread can pump messages and service COM
        private void DeferOptimizeColumnSizing()
        {
            if (this.IsHandleCreated && !this.Disposing && !this.IsDisposed)
            {
                try
                {
                    this.BeginInvoke(new System.Action(OptimizeColumnSizing));
                }
                catch
                {
                    // ignore if BeginInvoke fails
                }
            }
        }

        // Lightweight column sizing that uses DisplayedCells mode (much cheaper than AllCells)
        private void OptimizeColumnSizing()
        {
            if (dataGridView1 == null || dataGridView1.IsDisposed) return;

            try
            {
                dataGridView1.SuspendLayout();

                // Resize only displayed cells to keep cost bounded.
                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

                // Optionally cap large widths:
                foreach (DataGridViewColumn col in dataGridView1.Columns)
                {
                    if (col.Width > 600) col.Width = 600;
                }
            }
            catch
            {
                // swallow any error during optimization to avoid blocking UI thread
            }
            finally
            {
                try { dataGridView1.ResumeLayout(); } catch { }
            }
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }       

        private void ExportExcel()
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DateTime now = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);            
            var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            try
            {
                int rowCount = dataGridView1.Rows.Count;
                int colCount = dataGridView1.Columns.Count;

                // 1. Prepare data in a 2D Array (Faster than cell-by-cell Interop calls)
                object[,] data = new object[rowCount + 1, colCount];

                // Create Header
                for (int j = 0; j < colCount; j++)
                {
                    data[0, j] = dataGridView1.Columns[j].HeaderText;
                }

                // Fill Content
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        data[i + 1, j] = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }

                // 2. Write the array to Excel in one go
                Microsoft.Office.Interop.Excel.Range startCell = worksheet.Cells[1, 1];
                Microsoft.Office.Interop.Excel.Range endCell = worksheet.Cells[rowCount + 1, colCount];

                // Define the range using the indexer [start, end]
                Microsoft.Office.Interop.Excel.Range writeRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount + 1, colCount]];

                // 3. Set the values
                // 3.1 Target Column B (Invoice No.)
                // We target from row 2 (to skip header) down to the last data row
                Microsoft.Office.Interop.Excel.Range invoiceColumn = worksheet.Range["B2", $"B{rowCount + 1}"];
                //Set the NumberFormat to "@" (Text)
                invoiceColumn.NumberFormat = "@";

                // 3.2. Target Column G (Passport No.) from Row 2 to the end
                Microsoft.Office.Interop.Excel.Range passportColumn = worksheet.Range["G2", $"G{rowCount + 1}"];
                // Set to Text format
                passportColumn.NumberFormat = "@";               

                // 4. Now it is safe to insert the data
                writeRange.Value2 = data;               

                //Header Style (First Row)
                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, colCount]];
                headerRange.Font.Bold = true;
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                // 5. Formatting
                // Apply Font to the entire used range
                Microsoft.Office.Interop.Excel.Range fullRange = worksheet.UsedRange;
                fullRange.Font.Name = "Times New Roman";
                fullRange.Font.Size = 14;

                // Dynamic formatting for specific columns (J-L and I)
                // Adjust these letters/indices if your grid structure changes
                worksheet.Range["J2", $"L{rowCount + 1}"].NumberFormat = "#,##0";
                worksheet.Range["D2", $"D{rowCount + 1}"].NumberFormat = "dd/MM/yyyy";
                worksheet.Range["E2", $"E{rowCount + 1}"].NumberFormat = "dd/MM/yyyy";
                worksheet.Range["H2", $"H{rowCount + 1}"].NumberFormat = "dd/MM/yyyy";
                worksheet.Range["P2", $"P{rowCount + 1}"].NumberFormat = "dd/MM/yyyy";

                // 6. AUTO-FIT Columns (Makes column width flexible based on data)
                fullRange.Columns.AutoFit();

                // 7. Save Dialog
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.FileName = $"VAT refund report {now:dd-MM-yyyy HHmmss}.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("Export successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                // Proper Cleanup
                workbook.Close(false);
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        private void ExportSpreadsheet()
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("No data to export.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DateTime now = DateTime.Now;

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("VAT Report");

                    // 1. Set Global Font (Times New Roman, 14pt)
                    worksheet.Style.Font.FontName = "Times New Roman";
                    worksheet.Style.Font.FontSize = 14;

                    // 2. Export Headers from DataGridView
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        var cell = worksheet.Cell(1, j + 1);
                        cell.Value = dataGridView1.Columns[j].HeaderText;
                        cell.Style.Font.Bold = true;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    }

                    // 3. Export Data Rows
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            var cell = worksheet.Cell(i + 2, j + 1);
                            var val = dataGridView1.Rows[i].Cells[j].Value;

                            // Handle specific column types
                            string colLetter = worksheet.Column(j + 1).ColumnLetter();

                            if (val != null)
                            {
                                // Target Column B (Invoice) and G (Passport) as TEXT
                                if (colLetter == "B" || colLetter == "G")
                                {
                                    cell.SetValue(val.ToString());
                                    cell.Style.NumberFormat.Format = "@";
                                }
                                else
                                {
                                    cell.Value = val.ToString();
                                }
                            }
                        }
                    }

                    // 4. Apply Column-Specific Formatting (Triggering the Thousand Separator)
                    int lastRow = dataGridView1.Rows.Count + 1;

                    // Thousand Separator for J, K, L (Columns 10, 11, 12)
                    var amountRange = worksheet.Range(2, 10, lastRow, 12);
                    foreach (var cell in amountRange.Cells())
                    {
                        if (double.TryParse(cell.Value.ToString(), out double d))
                        {
                            cell.Value = d; // Convert string to number so format triggers
                            cell.Style.NumberFormat.Format = "#,##0";
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                        }
                    }

                    // Date Formatting for D, E, H, P
                    string[] dateCols = { "D", "E", "H", "P" };
                    foreach (var col in dateCols)
                    {
                        var dateRange = worksheet.Range($"{col}2:{col}{lastRow}");
                        dateRange.Style.NumberFormat.Format = "dd/MM/yyyy";
                    }

                    // 5. Auto-Fit Columns
                    worksheet.Columns().AdjustToContents();

                    // 6. Save Dialog
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Excel Files|*.xlsx";
                        saveFileDialog.FileName = $"VAT refund report {now:dd-MM-yyyy HHmmss}.xlsx";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            workbook.SaveAs(saveFileDialog.FileName);
                            MessageBox.Show("Export successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            // No "finally" block needed for cleanup! 'using' handles everything.
        }
        private void ExportToOpenOffice()
        {
            DateTime now = DateTime.Now;

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    ExportHeaders(worksheet);
                    ExportData(worksheet);
                    ApplyFormatting(worksheet);

                    SaveWorkbook(workbook, now);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ExportHeaders(IXLWorksheet worksheet)
        {
            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                worksheet.Cell(1, j + 1).Value = dataGridView1.Columns[j].HeaderText;
            }
        }

        private void ExportData(IXLWorksheet worksheet)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    var cellValue = dataGridView1.Rows[i].Cells[j].Value;
                    var cell = worksheet.Cell(i + 2, j + 1);

                    SetCellValueWithFormatting(cell, cellValue, j);
                }
            }
        }

        private void SetCellValueWithFormatting(IXLCell cell, object cellValue, int columnIndex)
        {
            if (cellValue == null)
            {
                cell.Value = "";
                return;
            }

            // Try to parse as number for specific columns that should have thousand separators
            if (IsNumericColumn(columnIndex))
            {
                if (decimal.TryParse(cellValue.ToString(), out decimal numericValue))
                {
                    cell.Value = numericValue;
                    // The formatting will be applied later in ApplyFormatting method
                }
                else
                {
                    cell.Value = cellValue.ToString();
                }
            }
            else
            {
                cell.Value = cellValue.ToString();
            }
        }

        private bool IsNumericColumn(int columnIndex)
        {
            // Define which columns should be treated as numbers with thousand separators
            // Based on your formatting ranges (J2:L5000 = columns 10-12)
            return columnIndex >= 9 && columnIndex <= 11; // Columns J, K, L (0-based index)
        }

        private void ApplyFormatting(IXLWorksheet worksheet)
        {
            ApplyColumnWidths(worksheet);
            ApplyGlobalFontStyles(worksheet);
            ApplyNumberFormats(worksheet);
        }

        private void ApplyColumnWidths(IXLWorksheet worksheet)
        {
            var columnWidths = new Dictionary<string, double>
            {
                ["A"] = 16,
                ["B"] = 16,
                ["C"] = 16,
                ["D"] = 16,
                ["E"] = 30,
                ["F"] = 16,
                ["G"] = 16,
                ["J"] = 16,
                ["K"] = 16,
                ["L"] = 16,
                ["M"] = 16,
                ["N"] = 50,
                ["O"] = 20,
                ["P"] = 20,
                ["Q"] = 20
            };

            foreach (var width in columnWidths)
            {
                worksheet.Column(width.Key).Width = width.Value;
            }
        }

        private void ApplyGlobalFontStyles(IXLWorksheet worksheet)
        {
            int lastRow = dataGridView1.RowCount + 1;
            string lastColumn = GetLastColumnLetter();

            // Data range font styles
            var dataRange = worksheet.Range($"A1:{lastColumn}{lastRow}");
            dataRange.Style.Font.FontName = "Times New Roman";
            dataRange.Style.Font.FontSize = 14;

            // Header specific styling
            var headerRange = worksheet.Range($"A1:{lastColumn}1");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        private string GetLastColumnLetter()
        {
            // Assuming your last column is Q based on the original code
            // You can make this dynamic based on dataGridView1.ColumnCount if needed
            return "Q";
        }

        private void ApplyNumberFormats(IXLWorksheet worksheet)
        {
            int lastRow = dataGridView1.RowCount + 1;

            // Apply thousand separator format to numeric columns (J, K, L)
            // Using #,##0 format which adds thousand separators
            worksheet.Range($"J2:L{lastRow}").Style.NumberFormat.Format = "#,##0";

            // Apply date format
            worksheet.Range($"I2:I{lastRow}").Style.NumberFormat.Format = "dd/MM/yyyy";
        }

        private void SaveWorkbook(XLWorkbook workbook, DateTime now)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "Save File";
                saveFileDialog.FileName = $"VAT refund report {now:dd-MM-yyyy hhmmss tt}.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("Export Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private void ImportBC03()
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();
            string loginName = ((Form)this.MdiParent).Controls["lblLoginName"].Text;
            //string loginName = "HQ10-0152";

            // --------Progess bar process----
            this.pro_panel1.Visible = true;
            this.pro_label.Text = "Preparing Excel...";
            this.pro_panel1.Refresh();
            //-------------------------------
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    using (var trans = conn.BeginTransaction())
                    {
                        using (SqlCommand cmd = new SqlCommand("", conn, trans))
                        {
                            cmd.CommandText = "SET ANSI_WARNINGS OFF";
                            cmd.ExecuteNonQuery();

                            //----Del tempt table----
                            cmd.CommandText = "Delete from tmpBC03";
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();

                            //Import from Excel file
                            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=Excel 8.0;";

                            //Create connecting object with Excel file
                            OleDbConnection Econn = new OleDbConnection(connectionString);
                            System.Data.DataTable Exceldt = new System.Data.DataTable();

                            OleDbDataAdapter dap = new OleDbDataAdapter("Select * from [BC03-BC BAN NGOAI TE$]", Econn);
                            dap.Fill(Exceldt);
                            Econn.Close();

                            //dgvNhaplieuExcel.DataSource = Exceldt;
                            //int dong = dgvNhaplieuExcel.RowCount;

                            for (int i = 9; i < Exceldt.Rows.Count - 4; i++)
                            {
                                DataRow r = Exceldt.Rows[i];

                                string thoigianGD = r[1].ToString().Trim();
                                string masoGD = Convert.ToString(r[2].ToString());
                                string hotenHK = Convert.ToString(r[3].ToString());
                                string sotienVNDHT = Convert.ToString(r[4].ToString()).Replace(",", "");

                                //var sotienvndht = sotienVNDHT.Replace(",","");
                                //string Ghichu = Convert.ToString(r["Ghichu"].ToString());

                                DateTime strngaynhapht = DateTime.Now;
                                string ngaynhapht = strngaynhapht.ToString("yyyy-MM-dd hh:mm:ss tt");

                                cmd.CommandText = "Insert Into tmpBC03 (ThoigianGD, MasoGD, HotenHK, SotienVNDHT, NgaynhapHT, LoginName) Values " +
                                    "(@ThoigianGD, @MasoGD, @HotenHK, @SotienVNDHT, @NgaynhapHT, @LoginName)";

                                //parameters declare

                                if (string.IsNullOrWhiteSpace(thoigianGD))
                                {
                                    cmd.Parameters.Add("@ThoigianGD", SqlDbType.DateTime).Value = DBNull.Value;
                                }
                                else //date is not null
                                {
                                    IFormatProvider provider = new CultureInfo("fr-FR");

                                    string[] listfullthoigiangd = thoigianGD.Split(' ');

                                    List<string> thoigiangd = new List<string>(listfullthoigiangd);
                                    thoigiangd.RemoveAt(1);

                                    string strthoigiangd = string.Join("", thoigiangd).TrimEnd();

                                    if (strthoigiangd.Length == 10)
                                    {
                                        DateTime pthoigianGD = DateTime.Parse(thoigianGD, provider);
                                        var strthoigianGD = pthoigianGD.ToString("yyyy-MM-dd hh:mm:ss tt");
                                        cmd.Parameters.Add("@ThoigianGD", SqlDbType.DateTime).Value = strthoigianGD;
                                    }
                                    else //datetime lenght is not 10 
                                    {
                                        cmd.Parameters.Add("@ThoigianGD", SqlDbType.DateTime).Value = DBNull.Value;
                                    }
                                }
                                cmd.Parameters.Add("@MasoGD", SqlDbType.VarChar).Value = masoGD;
                                cmd.Parameters.Add("@HotenHK", SqlDbType.VarChar).Value = hotenHK;
                                cmd.Parameters.Add("@SotienVNDHT", SqlDbType.Int).Value = sotienVNDHT;
                                cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;
                                cmd.Parameters.Add("@LoginName", SqlDbType.VarChar).Value = loginName;

                                //Thuc hien cau lenh
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();

                                //-------Progress processing-----------                            
                                this.pro_label.Text = "Reading Excel: " + (i - 8).ToString() + " of " + (Exceldt.Rows.Count - 13).ToString();
                                this.pro_panel1.Refresh();
                                //-------------------------------------
                            }
                            cmd.CommandText = "SET ANSI_WARNINGS ON";
                            cmd.ExecuteNonQuery();
                            trans.Commit();

                            SqlCommand cmd_ht = new SqlCommand("Select Distinct ROW_NUMBER() OVER (ORDER BY tmpb.MasoGD ASC) AS [STT], tmpb.ThoigianGD, tmpb.MasoGD, tmpb.HotenHK, tmpb.SotienVNDHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC03 as tmpb " +
                                                "Where tmpb.MasoGD Not In (Select Distinct b.MasoGD From BC03 as b Inner Join " +
                                                "tmpBC03 as tmpb On b.MasoGD = tmpb.MasoGD and b.ThoigianGD = tmpb.ThoigianGD)", conn);

                            SqlDataAdapter da_ht = new SqlDataAdapter();
                            da_ht.SelectCommand = cmd_ht;
                            dt = new System.Data.DataTable();
                            da_ht.Fill(dt);

                            //Caculate datarows before import new datarows
                            SqlCommand cmd_countbeimpt = new SqlCommand("SELECT COUNT(MasoGD) FROM BC03", conn);
                            int countbeimpt = Convert.ToInt32(cmd_countbeimpt.ExecuteScalar());

                            //Check if the data existed?
                            SqlCommand cmd_kt = new SqlCommand("Select Distinct ROW_NUMBER() OVER (ORDER BY tmpb.MasoGD ASC) AS [STT], b.MasoGD, b.ThoigianGD, b.HotenHK, b.SotienVNDHT, b.Ghichu, b.NgaynhapHT, b.LoginName " +
                                                "From BC03 as b Inner Join " +
                                                "tmpBC03 as tmpb On b.MasoGD = tmpb.MasoGD and b.ThoigianGD = tmpb.ThoigianGD", conn);

                            da = new SqlDataAdapter();
                            da.SelectCommand = cmd_kt;
                            System.Data.DataTable dtb_kt = new System.Data.DataTable();
                            da.Fill(dtb_kt);

                            if (dtb_kt.Rows.Count >= 1)
                            {
                                SetupDataGridView();
                                //-----Them cot STT----------------
                                //dtb_kt.Columns.Add("STT");
                                //for (int i = 0; i < dtb_kt.Rows.Count; i++)
                                //{
                                //    dtb_kt.Rows[i]["STT"] = i + 1;
                                //}
                                dataGridView1.DataSource = dtb_kt;
                                dataGridView1.Columns["STT"].DisplayIndex = 0;

                                DialogResult diag = MessageBox.Show("" + dtb_kt.Rows.Count.ToString() + " datarow(s) existed. Do you really want to export existed datarow(s)?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                if (diag == DialogResult.Yes) // Co ket xuat du lieu trung
                                {
                                    ExportExcel();
                                }

                                //Import data into the database
                                cmd.CommandText = "Insert Into BC03 Select Distinct tmpb.ThoigianGD, tmpb.MasoGD, tmpb.HotenHK, tmpb.SotienVNDHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC03 as tmpb " +
                                                "Where tmpb.MasoGD Not In (Select Distinct b.MasoGD From BC03 as b Inner Join " +
                                                "tmpBC03 as tmpb On b.MasoGD = tmpb.MasoGD and b.ThoigianGD = tmpb.ThoigianGD)";
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                //Import data
                                cmd.CommandText = "Insert Into BC03 Select Distinct tmpb.ThoigianGD, tmpb.MasoGD, tmpb.HotenHK, tmpb.SotienVNDHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC03 as tmpb ";
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                            }

                            SqlCommand cmd_countatmpt = new SqlCommand("SELECT COUNT(MasoGD) FROM BC03", conn);
                            int countatimpt = Convert.ToInt32(cmd_countatmpt.ExecuteScalar());

                            int impdatarows = countatimpt - countbeimpt;

                            if (impdatarows > 0)
                            {
                                MessageBox.Show("Imported" + " " + impdatarows + " " + " datarow(s) successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                //Count datarows in the DB                     
                                txtTotalrows.Text = Convert.ToString(countatimpt.ToString("#,##0"));

                                this.Refresh();
                            }
                        }
                    }
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception ex)
            {
                this.pro_panel1.Visible = false;
                //-----------------
                MessageBox.Show("Datarows imported unsuccessfully: " + ex.Message + "", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //-------------------------------
            this.pro_panel1.Visible = false;
        }
        private void ImportBC04()
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();
            string loginName = ((Form)this.MdiParent).Controls["lblLoginName"].Text;
            //string loginName = "HQ10-0152";

            //--------Progess bar process----
            this.pro_panel1.Visible = true;
            this.pro_label.Text = "Preparing Excel...";
            this.pro_panel1.Refresh();
            //-------------------------------
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    using (var trans = conn.BeginTransaction())
                    {
                        using (SqlCommand cmd = new SqlCommand("", conn, trans))
                        {
                            cmd.CommandText = "SET ANSI_WARNINGS OFF";
                            cmd.ExecuteNonQuery();

                            //Xoa bang tam
                            cmd.CommandText = "Delete from tmpBC04";
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();

                            //Nhap du lieu tu Excel
                            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=Excel 8.0;";

                            //}
                            // Tạo đối tượng kết nối voi file Excel
                            OleDbConnection Econn = new OleDbConnection(connectionString);
                            System.Data.DataTable Exceldt = new System.Data.DataTable();

                            OleDbDataAdapter dap = new OleDbDataAdapter("Select * From [BC04-BANG KE THANH TOAN$]", Econn);

                            dap.Fill(Exceldt);
                            Econn.Close();
                            //dgvNhaplieuExcel.DataSource = Exceldt;
                            //int dong = dgvNhaplieuExcel.RowCount;

                            #region
                            for (int i = 13; i < Exceldt.Rows.Count - 8; i++)
                            {
                                DataRow r = Exceldt.Rows[i];

                                string kyhieuSoNgay = r[1].ToString().Trim();
                                string masoGD = Convert.ToString(r[2].ToString());
                                string tenDNBH = Convert.ToString(r[3].ToString());
                                string sotienVATHD = Convert.ToString(r[4].ToString()).Replace(",", "");
                                string ngayHT = r[5].ToString().Trim();

                                DateTime strngaynhapht = DateTime.Now;
                                string ngaynhapht = strngaynhapht.ToString("yyyy-MM-dd hh:mm:ss tt");

                                cmd.CommandText = "Insert Into tmpBC04 (KyhieuSoNgay, MasoGD, TenDNBH, SotienVATHD, NgayHT, NgaynhapHT, LoginName) Values " +
                                                "(@KyhieuSoNgay, @MasoGD, @TenDNBH, @SotienVATHD, @NgayHT, @NgaynhapHT, @LoginName)";

                                //parameters declare
                                cmd.Parameters.Add("@KyhieuSoNgay", SqlDbType.VarChar).Value = kyhieuSoNgay;
                                cmd.Parameters.Add("@MasoGD", SqlDbType.VarChar).Value = masoGD;
                                cmd.Parameters.Add("@TenDNBH", SqlDbType.NVarChar).Value = tenDNBH;
                                cmd.Parameters.Add("@SotienVATHD", SqlDbType.Int).Value = sotienVATHD;

                                if (string.IsNullOrWhiteSpace(ngayHT))
                                {
                                    cmd.Parameters.Add("@NgayHT", SqlDbType.DateTime).Value = DBNull.Value;
                                }
                                else //date is not null
                                {
                                    IFormatProvider provider = new CultureInfo("fr-FR");

                                    DateTime pngayHT = DateTime.Parse(ngayHT, provider);
                                    var strngayHT = pngayHT.ToString("yyyy-MM-dd hh:mm:ss tt");
                                    cmd.Parameters.Add("@NgayHT", SqlDbType.DateTime).Value = strngayHT;
                                }

                                cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;
                                cmd.Parameters.Add("@LoginName", SqlDbType.VarChar).Value = loginName;

                                //Thuc hien cau lenh
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();

                                //-------Progress processing-----------                            
                                this.pro_label.Text = "Reading Excel: " + (i - 12).ToString() + " of " + (Exceldt.Rows.Count - 21).ToString();
                                this.pro_panel1.Refresh();
                                //------------------------------
                            }
                            cmd.CommandText = "SET ANSI_WARNINGS ON";
                            cmd.ExecuteNonQuery();
                            trans.Commit();

                            SqlCommand cmd_ht = new SqlCommand("Select Distinct ROW_NUMBER() OVER (ORDER BY tmpb.KyhieuSoNgay ASC) AS [STT], tmpb.KyhieuSoNgay, tmpb.MasoGD, tmpb.TenDNBH, tmpb.SotienVATHD, tmpb.NgayHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC04 as tmpb " +
                                                "Where tmpb.KyhieuSoNgay Not In (Select Distinct b.KyhieuSoNgay From BC04 as b Inner Join " +
                                                "tmpBC04 as tmpb On b.KyhieuSoNgay = tmpb.KyhieuSoNgay and b.MasoGD = tmpb.MasoGD) Order by tmpb.MasoGD", conn);

                            SqlDataAdapter da_ht = new SqlDataAdapter();
                            da_ht.SelectCommand = cmd_ht;
                            dt = new System.Data.DataTable();
                            da_ht.Fill(dt);

                            #region
                            //if (dt.Rows.Count >= 1)
                            //{
                            //    SetupDataGridView();

                            //    ////Them cot STT
                            //    //dt.Columns.Add("STT");
                            //    //for (int i = 0; i < dt.Rows.Count; i++)
                            //    //{
                            //    //    dt.Rows[i]["STT"] = i + 1;
                            //    //}
                            //    dgvNhaplieuExcel.DataSource = dt;
                            //    dgvNhaplieuExcel.Columns["STT"].DisplayIndex = 0;
                            //}
                            #endregion

                            //Caculate datarows before import new datarows
                            SqlCommand cmd_countbeimpt = new SqlCommand("SELECT COUNT(KyhieuSoNgay) FROM BC04", conn);
                            int countbeimpt = Convert.ToInt32(cmd_countbeimpt.ExecuteScalar());

                            //Check if the data existed?
                            SqlCommand cmd_kt = new SqlCommand("Select Distinct ROW_NUMBER() OVER (ORDER BY tmpb.KyhieuSoNgay ASC) AS [STT], b.KyhieuSoNgay, b.MasoGD, b.TenDNBH, b.SotienVATHD, b.NgayHT, b.Ghichu, b.NgaynhapHT, b.LoginName " +
                                                "From BC04 as b Inner Join " +
                                                "tmpBC04 as tmpb On b.KyhieuSoNgay = tmpb.KyhieuSoNgay and b.MasoGD = tmpb.MasoGD", conn);

                            da = new SqlDataAdapter();
                            da.SelectCommand = cmd_kt;
                            System.Data.DataTable dtb_kt = new System.Data.DataTable();
                            da.Fill(dtb_kt);

                            if (dtb_kt.Rows.Count >= 1)
                            {
                                SetupDataGridView();

                                ////Them cot STT
                                //dtb_kt.Columns.Add("STT");
                                //for (int i = 0; i < dtb_kt.Rows.Count; i++)
                                //{
                                //    dtb_kt.Rows[i]["STT"] = i + 1;
                                //}
                                dataGridView1.DataSource = dtb_kt;
                                dataGridView1.Columns["STT"].DisplayIndex = 0;

                                DialogResult diag = MessageBox.Show("" + dtb_kt.Rows.Count.ToString() + " datarow(s) existed. Please check before importing. Do you want to export existed datarow(s)?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                if (diag == DialogResult.Yes) // Co ket xuat du lieu trung
                                {
                                    ExportExcel();
                                }

                                //Import datarows into database
                                cmd.CommandText = "Insert Into BC04 Select Distinct tmpb.KyhieuSoNgay, tmpb.MasoGD, tmpb.TenDNBH, tmpb.SotienVATHD, tmpb.NgayHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC04 as tmpb " +
                                                "Where tmpb.KyhieuSoNgay Not In (Select Distinct b.KyhieuSoNgay From BC04 as b Inner Join " +
                                                "tmpBC04 as tmpb On b.KyhieuSoNgay = tmpb.KyhieuSoNgay and b.MasoGD = tmpb.MasoGD)";
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                //Import datarows into database
                                cmd.CommandText = "Insert Into BC04 Select Distinct tmpb.KyhieuSoNgay, tmpb.MasoGD, tmpb.TenDNBH, tmpb.SotienVATHD, tmpb.NgayHT, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC04 as tmpb ";
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                            }

                            SqlCommand cmd_countatmpt = new SqlCommand("SELECT COUNT(KyhieuSoNgay) FROM BC04", conn);
                            int countatimpt = Convert.ToInt32(cmd_countatmpt.ExecuteScalar());

                            int impdatarows = countatimpt - countbeimpt;

                            if (impdatarows > 0)
                            {
                                MessageBox.Show("Imported" + " " + impdatarows + " " + " datarow(s) successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                //Count datarows in the DB                     
                                txtTotalrows.Text = Convert.ToString(countatimpt.ToString("#,##0"));
                                this.Refresh();
                            }

                            #endregion
                        }
                    }
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception ex)
            {
                //-------------------------------
                this.pro_panel1.Visible = false;
                MessageBox.Show("Datarows import unsuccessfully: " + ex.Message + "", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //-------------------------------
            this.pro_panel1.Visible = false;
        }
        private void ImportBC09()
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();
            string loginName = ((Form)this.MdiParent).Controls["lblLoginName"].Text;
            //string loginName = "HQ10-0152";

            // --------Progess bar process----
            this.pro_panel1.Visible = true;
            this.pro_label.Text = "Preparing Excel...";
            this.pro_panel1.Refresh();
            // -------------------------------
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    using (var trans = conn.BeginTransaction())
                    {
                        using (SqlCommand cmd = new SqlCommand("", conn, trans))
                        {
                            cmd.CommandText = "SET ANSI_WARNINGS OFF";
                            cmd.ExecuteNonQuery();

                            //Xoa bang tam
                            cmd.CommandText = "Delete from tmpBC09";
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();

                            //Nhap du lieu tu Excel
                            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=Excel 8.0;";

                            // Tạo đối tượng kết nối voi file Excel
                            OleDbConnection Econn = new OleDbConnection(connectionString);
                            System.Data.DataTable Exceldt = new System.Data.DataTable();

                            OleDbDataAdapter dap = new OleDbDataAdapter("Select * From [Table 1$]", Econn);

                            dap.Fill(Exceldt);
                            Econn.Close();
                            //dataGridView1.DataSource = Exceldt;
                            int dong = dataGridView1.RowCount;

                            #region

                            for (int i = 12; i < Exceldt.Rows.Count - 10; i++)
                            {
                                #region

                                DataRow r = Exceldt.Rows[i];

                                string kyhieuHD = r[1].ToString().Trim();
                                string soHD = Convert.ToString(r[2].ToString());
                                string ngayHD = r[3].ToString().Trim();
                                string tenDNBH = Convert.ToString(r[4].ToString());
                                string masoDN = Convert.ToString(r[5].ToString());
                                string hoTenHK = Convert.ToString(r[6].ToString());
                                string soHC = Convert.ToString(r[7].ToString());
                                string ngayHC = r[8].ToString().Trim();
                                string quoctich = Convert.ToString(r[9].ToString());
                                string trigiaHHchuaVAT = Convert.ToString(r[10].ToString()).Replace(",", "");
                                string ngayHT = r[13].ToString().Trim();
                                string sotienVATDH = Convert.ToString(r[14].ToString()).Replace(",", "");
                                string sotienDVNHH = Convert.ToString(r[15].ToString()).Replace(",", "");

                                DateTime strngaynhapht = DateTime.Now;
                                string ngaynhapht = strngaynhapht.ToString("yyyy-MM-dd hh:mm:ss tt");

                                cmd.CommandText = "Insert Into tmpBC09 (KyhieuHD, SoHD, NgayHD, TenDNBH, MasoDN, HoTenHK, SoHC, NgayHC, Quoctich, " +
                                                  "TrigiaHHchuaVAT, NgayHT, SotienVATDH, SotienDVNHH, NgaynhapHT, LoginName) Values " +
                                                "(@KyhieuHD, @SoHD, @NgayHD, @TenDNBH, @MasoDN, @HoTenHK, @SoHC, @NgayHC, @Quoctich, " +
                                                  "@TrigiaHHchuaVAT, @NgayHT, @SotienVATDH, @SotienDVNHH, @NgaynhapHT, @LoginName)";

                                //parameters declare
                                cmd.Parameters.Add("@KyhieuHD", SqlDbType.VarChar).Value = kyhieuHD;
                                cmd.Parameters.Add("@SoHD", SqlDbType.VarChar).Value = soHD;
                                if (string.IsNullOrWhiteSpace(ngayHD))
                                {
                                    cmd.Parameters.Add("@NgayHD", SqlDbType.DateTime).Value = DBNull.Value;
                                }
                                else //date is not null
                                {
                                    IFormatProvider provider = new CultureInfo("fr-FR");

                                    DateTime pngayHD = DateTime.Parse(ngayHD, provider);
                                    var strngayHD = pngayHD.ToString("yyyy-MM-dd hh:mm:ss tt");
                                    cmd.Parameters.Add("@NgayHD", SqlDbType.DateTime).Value = strngayHD;
                                }

                                cmd.Parameters.Add("@TenDNBH", SqlDbType.NVarChar).Value = tenDNBH;
                                cmd.Parameters.Add("@MasoDN", SqlDbType.VarChar).Value = masoDN;
                                cmd.Parameters.Add("@HoTenHK", SqlDbType.NVarChar).Value = hoTenHK;
                                cmd.Parameters.Add("@SoHC", SqlDbType.VarChar).Value = soHC;
                                if (string.IsNullOrWhiteSpace(ngayHC))
                                {
                                    cmd.Parameters.Add("@NgayHC", SqlDbType.DateTime).Value = DBNull.Value;
                                }
                                else //date is not null
                                {
                                    IFormatProvider provider = new CultureInfo("fr-FR");

                                    DateTime pngayHC = DateTime.Parse(ngayHC, provider);
                                    var strngayHC = pngayHC.ToString("yyyy-MM-dd hh:mm:ss tt");
                                    cmd.Parameters.Add("@NgayHC", SqlDbType.DateTime).Value = strngayHC;
                                }
                                cmd.Parameters.Add("@Quoctich", SqlDbType.VarChar).Value = quoctich;
                                cmd.Parameters.Add("@TrigiaHHchuaVAT", SqlDbType.BigInt).Value = trigiaHHchuaVAT;
                                if (string.IsNullOrWhiteSpace(ngayHT))
                                {
                                    cmd.Parameters.Add("@NgayHT", SqlDbType.DateTime).Value = DBNull.Value;
                                }
                                else //date is not null
                                {
                                    IFormatProvider provider = new CultureInfo("fr-FR");

                                    DateTime pngayHT = DateTime.Parse(ngayHT, provider);
                                    var strngayHT = pngayHT.ToString("yyyy-MM-dd hh:mm:ss tt");
                                    cmd.Parameters.Add("@NgayHT", SqlDbType.DateTime).Value = strngayHT;
                                }
                                cmd.Parameters.Add("@SotienVATDH", SqlDbType.Int).Value = sotienVATDH;
                                cmd.Parameters.Add("@SotienDVNHH", SqlDbType.Int).Value = sotienDVNHH;
                                cmd.Parameters.Add("@NgaynhapHT", SqlDbType.DateTime).Value = ngaynhapht;
                                cmd.Parameters.Add("@LoginName", SqlDbType.VarChar).Value = loginName;

                                //Thuc hien cau lenh
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                                cmd.Parameters.Clear();

                                //--------Progess bar process----                                
                                this.pro_label.Text = "Reading Excel: " + (i - 11).ToString() + " of " + (Exceldt.Rows.Count - 22).ToString();
                                this.pro_panel1.Refresh();
                                //-------------------------------
                            }
                            cmd.CommandText = "SET ANSI_WARNINGS ON";
                            cmd.ExecuteNonQuery();
                            trans.Commit();

                            #endregion

                            SqlCommand cmd_ht = new SqlCommand("Select Distinct ROW_NUMBER() OVER (ORDER BY tmpb.NgayHT ASC) AS [STT], tmpb.KyhieuHD, tmpb.SoHD, tmpb.NgayHD, tmpb.TenDNBH, tmpb.MasoDN, tmpb.HoTenHK, tmpb.SoHC, tmpb.NgayHC, " +
                                "tmpb.Quoctich, tmpb.TrigiaHHchuaVAT, tmpb.NgayHT, tmpb.SotienVATDH, tmpb.SotienDVNHH, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC09 as tmpb " +
                                                "Where tmpb.SoHD Not In (Select Distinct b.SoHD " +
                                                "From BC09 as b " +
                                                "Inner Join tmpBC09 as tmpb " +
                                                "On b.KyhieuHD = tmpb.KyhieuHD and b.SoHD = tmpb.SoHD and b.SoHC = tmpb.SoHC and b.NgayHD = tmpb.NgayHD) " +
                                                "And tmpb.KyhieuHD Not In (Select Distinct b.KyhieuHD " +
                                                "From BC09 as b " +
                                                "Inner Join tmpBC09 as tmpb On b.KyhieuHD = tmpb.KyhieuHD and b.SoHD = tmpb.SoHD and b.SoHC = tmpb.SoHC and b.NgayHD = tmpb.NgayHD)", conn);

                            SqlDataAdapter da_ht = new SqlDataAdapter();
                            da_ht.SelectCommand = cmd_ht;
                            dt = new System.Data.DataTable();
                            da_ht.Fill(dt);

                            //Caculate datarows before import new datarows
                            SqlCommand cmd_countbeimpt = new SqlCommand("SELECT COUNT(SoHD) FROM BC09", conn);
                            int countbeimpt = Convert.ToInt32(cmd_countbeimpt.ExecuteScalar());

                            //Check if the datarows existed?
                            SqlCommand cmd_kt = new SqlCommand("Select Distinct ROW_NUMBER() OVER (ORDER BY tmpb.NgayHT ASC) AS [STT], b.KyhieuHD, b.SoHD, b.NgayHD, b.TenDNBH, " +
                                "b.MasoDN, b.HoTenHK, b.SoHC, b.NgayHC, b.Quoctich, b.TrigiaHHchuaVAT, b.NgayHT, b.SotienVATDH, b.SotienDVNHH, b.Ghichu, b.NgaynhapHT, b.LoginName " +
                                                "From BC09 as b " +
                                                "Inner Join tmpBC09 as tmpb " +
                                                "On b.KyhieuHD = tmpb.KyhieuHD and b.SoHD = tmpb.SoHD and b.NgayHD = tmpb.NgayHD and b.MasoDN = tmpb.MasoDN and b.SoHC = tmpb.SoHC", conn);
                            //ROW_NUMBER() OVER (ORDER BY b.NgayHT ASC) AS [STT],

                            da = new SqlDataAdapter();
                            da.SelectCommand = cmd_kt;
                            System.Data.DataTable dtb_kt = new System.Data.DataTable();
                            da.Fill(dtb_kt);

                            if (dtb_kt.Rows.Count > 0)
                            {
                                SetupDataGridView();

                                //Them cot STT
                                //dtb_kt.Columns.Add("STT");
                                //for (int i = 0; i < dtb_kt.Rows.Count; i++)
                                //{
                                //    dtb_kt.Rows[i]["STT"] = i + 1;
                                //}
                                dataGridView1.DataSource = dtb_kt;
                                dataGridView1.Columns["STT"].DisplayIndex = 0;

                                DialogResult diag = MessageBox.Show("" + dtb_kt.Rows.Count.ToString() + " datarow(s) existed. " +
                                    "Please check before importing. Do you want to export existed datarow(s)?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                if (diag == DialogResult.Yes) // Export dupplicated datarows
                                {
                                    ExportExcel();
                                }
                                //Import datarows
                                cmd.CommandText = "Insert Into BC09 Select Distinct tmpb.KyhieuHD, tmpb.SoHD, tmpb.NgayHD, tmpb.TenDNBH, tmpb.MasoDN, tmpb.HoTenHK, tmpb.SoHC, tmpb.NgayHC, tmpb.Quoctich, " +
                                                  "tmpb.TrigiaHHchuaVAT, tmpb.NgayHT, tmpb.SotienVATDH, tmpb.SotienDVNHH, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC09 as tmpb " +
                                                "Where tmpb.SoHD Not In (Select Distinct b.SoHD " +
                                                "From BC09 as b " +
                                                "Inner Join tmpBC09 as tmpb " +
                                                "On b.KyhieuHD = tmpb.KyhieuHD and b.SoHD = tmpb.SoHD and b.SoHC = tmpb.SoHC and b.NgayHD = tmpb.NgayHD) " +
                                                "And tmpb.KyhieuHD Not In (Select Distinct b.KyhieuHD " +
                                                "From BC09 as b " +
                                                "Inner Join tmpBC09 as tmpb " +
                                                "On b.KyhieuHD = tmpb.KyhieuHD and b.SoHD = tmpb.SoHD and b.SoHC = tmpb.SoHC and b.NgayHD = tmpb.NgayHD)";
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                //Import datarows
                                cmd.CommandText = "Insert Into BC09 Select Distinct tmpb.KyhieuHD, tmpb.SoHD, tmpb.NgayHD, tmpb.TenDNBH, tmpb.MasoDN, tmpb.HoTenHK, tmpb.SoHC, tmpb.NgayHC, tmpb.Quoctich, " +
                                                  "tmpb.TrigiaHHchuaVAT, tmpb.NgayHT, tmpb.SotienVATDH, tmpb.SotienDVNHH, tmpb.Ghichu, tmpb.NgaynhapHT, tmpb.LoginName " +
                                                "From tmpBC09 as tmpb ";
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();

                                dataGridView1.DataSource = dt;
                            }

                            SqlCommand cmd_countatmpt = new SqlCommand("SELECT COUNT(SoHD) FROM BC09", conn);
                            int countatimpt = Convert.ToInt32(cmd_countatmpt.ExecuteScalar());

                            int impdatarows = countatimpt - countbeimpt;

                            if (impdatarows > 0)
                            {
                                MessageBox.Show("Imported" + " " + impdatarows + " " + " datarow(s) successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }

                            //Count datarows in the DB                     
                            txtTotalrows.Text = Convert.ToString(countatimpt.ToString("#,##0"));
                            this.Refresh();
                            #endregion
                        }
                    }
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception ex)
            {
                this.pro_panel1.Visible = false;
                MessageBox.Show("Datarows import unsuccessfully: " + ex.Message + "", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            // --------Progess bar process----
            this.pro_panel1.Visible = false;
        }
        private void ImportUser()
        {
            Utility ut = new Utility();
            var conn = ut.OpenDB();

            string loginName = ((Form)this.MdiParent).Controls["lblLoginName"].Text;
            //string loginName = "HQ10-0152";
            dt = new System.Data.DataTable("Congchuc");

            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();

                using (var trans = conn.BeginTransaction())
                {
                    using (SqlCommand cmd = new SqlCommand("", conn, trans))
                    {
                        cmd.CommandText = "SET ANSI_WARNINGS OFF";
                        cmd.ExecuteNonQuery();


                        //Xoa bang tam
                        cmd.CommandText = "Delete from tmpCongchuc";
                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();

                        //Nhap du lieu tu Excel
                        string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtFilePath.Text.Trim() + ";Extended Properties=Excel 8.0;";

                        // Tạo đối tượng kết nối voi file Excel
                        OleDbConnection Econn = new OleDbConnection(connectionString);
                        System.Data.DataTable Exceldt = new System.Data.DataTable();

                        OleDbDataAdapter dap = new OleDbDataAdapter("Select * From [Sheet1$]", Econn);

                        dap.Fill(Exceldt);
                        //dgvNhaplieuExcel.DataSource = Exceldt;              

                        for (int i = 0; i < Exceldt.Rows.Count; i++)
                        {
                            DataRow r = Exceldt.Rows[i];

                            string sohieucc = r[0].ToString().Trim();
                            string tencc = r[1].ToString().Trim();
                            string ghichu = r[2].ToString().Trim();

                            cmd.CommandText = "Insert Into tmpCongchuc (SHCC, TenCC, Ghichu) Values (@SHCC, @TenCC, @Ghichu)";

                            //parameters declare
                            cmd.Parameters.Add("@SHCC", SqlDbType.VarChar).Value = sohieucc;
                            cmd.Parameters.Add("@TenCC", SqlDbType.NVarChar).Value = tencc;
                            cmd.Parameters.Add("@Ghichu", SqlDbType.NVarChar).Value = ghichu;

                            //Thuc hien cau lenh
                            cmd.CommandType = CommandType.Text;
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }

                        cmd.CommandText = "SET ANSI_WARNINGS ON";
                        cmd.ExecuteNonQuery();
                        trans.Commit();
                    }

                    cmd = new SqlCommand("Select Distinct ROW_NUMBER() OVER (ORDER BY tmpc.SHCC ASC) AS [STT], tmpc.SHCC, tmpc.TenCC, tmpc.Ghichu " +
                                                "From tmpCongchuc as tmpc " +
                                                "Where tmpc.SHCC Not In (Select Distinct c.SHCC From Congchuc as c Inner Join " +
                                                "tmpCongchuc as tmpc On c.SHCC = tmpc.SHCC)", conn);
                    da = new SqlDataAdapter();
                    da.SelectCommand = cmd;
                    dt = new System.Data.DataTable();
                    da.Fill(dt);

                    //Caculate datarows before import new datarows
                    SqlCommand cmd_countbeimpt = new SqlCommand("SELECT COUNT(SHCC) FROM Congchuc", conn);
                    int countbeimpt = Convert.ToInt32(cmd_countbeimpt.ExecuteScalar());

                    //Check if the datarows existed?
                    SqlCommand cmd_kt = new SqlCommand("Select Count(tmpc.SHCC) " +
                        "From tmpCongchuc as tmpc Inner Join Congchuc as c On tmpc.SHCC = c.SHCC", conn);
                    int count_kt = Convert.ToInt32(cmd_kt.ExecuteScalar());

                    if (count_kt >= 1)
                    {
                        DialogResult diag = MessageBox.Show("" + count_kt.ToString() + " datarow(s) existed. Please check before importing. Do you want to export existed datarow(s)?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (diag == DialogResult.Yes) // Export dupplicated datarows
                        {
                            ExportExcel();
                        }

                        //Import datarows
                        cmd.CommandText = "Insert Into Congchuc Select Distinct tmpc.SHCC, tmpc.TenCC, tmpc.Ghichu " +
                                                "From tmpCongchuc as tmpc " +
                                                "Where tmpc.SHCC Not In (Select Distinct c.SHCC From Congchuc as c Inner Join " +
                                                "tmpCongchuc as tmpc On c.SHCC = tmpc.SHCC)";

                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        cmd.CommandText = "Insert Into Congchuc Select Distinct tmpc.SHCC, tmpc.TenCC, tmpc.Ghichu " +
                                                "From tmpCongchuc as tmpc";

                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();
                    }

                    SetupDataGridView();
                    dataGridView1.DataSource = dt;

                    SqlCommand cmd_countatmpt = new SqlCommand("SELECT COUNT(SHCC) FROM Congchuc", conn);
                    int countatimpt = Convert.ToInt32(cmd_countatmpt.ExecuteScalar());

                    int impdatarows = countatimpt - countbeimpt;

                    MessageBox.Show("Import" + " " + impdatarows + " " + " datarow(s) successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //Count datarows in the DB                     
                    txtTotalrows.Text = Convert.ToString(countatimpt.ToString("#,##0"));

                    TextboxClear();
                    this.Refresh();
                }
                conn.Close();
                conn.Dispose();
            }
        }
        private void btnNhaplieu_Click(object sender, EventArgs e)
        {
            if (ValidInput() == true && txtFilePath.Text.Contains(".xls") && txtFilePath.Text.Contains("BC03"))
            {
                ImportBC03();
            }
            else if (ValidInput() == true && txtFilePath.Text.Contains(".xls") && txtFilePath.Text.Contains("BC04"))
            {
                ImportBC04();
            }
            else if (ValidInput() == true && txtFilePath.Text.Contains(".xls") && txtFilePath.Text.Contains("BC09"))
            {
                ImportBC09();
            }
            else
            {
                ImportUser();
            }
        }

        private void ConfigureNumericColumnStyles()
        {
            var numericStyle = new DataGridViewCellStyle
            {
                Format = "N0",
                Alignment = DataGridViewContentAlignment.MiddleRight
            };

            string[] numericColumns = { "TrigiaHHchuaVAT", "SotienVATDH", "SotienDVNHH" };

            foreach (var columnName in numericColumns)
            {
                if (dataGridView1.Columns.Contains(columnName))
                {
                    dataGridView1.Columns[columnName].DefaultCellStyle = numericStyle;
                }
            }
        }
        private void ApplyConditionalFormatting(DataGridViewCellFormattingEventArgs e)
        {
            // Skip if not the SotienVATDH column or value is null
            if (e.ColumnIndex < 0 ||
                e.RowIndex < 0 ||
                dataGridView1.Columns[e.ColumnIndex].Name != "SotienVATDH" ||
                e.Value == null)
            {
                return;
            }

            // Try to parse and apply formatting if value meets condition
            if (decimal.TryParse(e.Value.ToString(), out decimal sotienVATDH) &&
                sotienVATDH >= 20000000)
            {
                e.CellStyle.ForeColor = Color.DarkRed;
                e.CellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12, FontStyle.Bold);
            }
        }
        private void LoadData()
        {
            try
            {
                Utility ut = new Utility();
                var conn = ut.OpenDB();
                int showRows = Int32.Parse(txtshowRows.Text.Trim());

                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    if (showRows == 100 || showRows == 0)
                    {
                        cmd = new SqlCommand("Select Top 100 ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], bc09.SoHD, bc09.KyhieuHD, bc09.NgayHD, " +
                        "bc09.NgayHT, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH, " +
                        "bc09.MasoDN, bc09.TenDNBH, bc09.NgaynhapHT, bc09.LoginName " +
                        "From BC09 as bc09", conn);
                    }
                    else
                    {
                        cmd = new SqlCommand("Select Top " + showRows + " ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], bc09.SoHD, bc09.KyhieuHD, bc09.NgayHD, " +
                        "bc09.NgayHT, bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH, " +
                        "bc09.MasoDN, bc09.TenDNBH, bc09.NgaynhapHT, bc09.LoginName " +
                        "From BC09 as bc09", conn);
                    }

                    da = new SqlDataAdapter(cmd);
                    dt = new System.Data.DataTable();
                    da.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dt;

                        SetupDataGridView();

                        object tongsotienVATDH, tongtrigiaHHchuaVAT, tongsotienDVNHH;

                        tongsotienVATDH = dt.Compute("Sum(SotienVATDH)", string.Empty);
                        decimal tongsotienvatdh = decimal.Parse(string.Format("{0}", tongsotienVATDH));
                        txtTotalRefundAmt.Text = Convert.ToString(tongsotienvatdh.ToString("#,##0"));

                        tongtrigiaHHchuaVAT = dt.Compute("Sum(TrigiaHHchuaVAT)", string.Empty);
                        decimal tongtrigiahhchuavat = decimal.Parse(string.Format("{0}", tongtrigiaHHchuaVAT));
                        txtTotalValue.Text = Convert.ToString(tongtrigiahhchuavat.ToString("#,##0"));

                        tongsotienDVNHH = dt.Compute("Sum(SotienDVNHH)", string.Empty);
                        decimal tongsotiendvnhh = decimal.Parse(string.Format("{0}", tongsotienDVNHH));
                        txtTotalBankServiceFee.Text = Convert.ToString(tongsotiendvnhh.ToString("#,##0"));

                        //// Old code to count distinct passports (--> wrong)
                        //int totalPassenger = dt
                        //            .AsEnumerable()
                        //            .Select(r => r.Field<string>("SoHC"))
                        //            .Distinct()
                        //            .Count();

                        // 1. Group by Date and SoHC to get unique passports per day
                        var dailyUniqueCount = dt.AsEnumerable()
                            .Where(row => row["SoHC"] != DBNull.Value && row["NgayHT"] != DBNull.Value)
                            .Select(row => new
                            {
                                Date = Convert.ToDateTime(row["NgayHT"]).Date,
                                Passport = row["SoHC"].ToString()
                            })
                            .Distinct() // This ensures we only count a passport once per day
                            .Count();   // This gives you the final total number
                        txtTotalPassenger.Text = (+dailyUniqueCount).ToString("#,##0");

                        //Count datarows in the DB                        
                        txtTotalrows.Text = Convert.ToString(dt.Rows.Count.ToString("#,##0"));

                        //Binding data 
                        #region
                        txtSoHC.DataBindings.Clear();
                        txtSoHC.DataBindings.Add("Text", dataGridView1.DataSource, "SoHC");
                        txtSoHD.DataBindings.Clear();
                        txtSoHD.DataBindings.Add("Text", dataGridView1.DataSource, "SoHD");
                        txtNgayHD.DataBindings.Clear();
                        txtNgayHD.DataBindings.Add("Text", dataGridView1.DataSource, "NgayHD");
                        txtGoodsValue.DataBindings.Clear();
                        txtGoodsValue.DataBindings.Add("Text", dataGridView1.DataSource, "TrigiaHHchuaVAT", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
                        txtRefundAmt.DataBindings.Clear();
                        txtRefundAmt.DataBindings.Add("Text", dataGridView1.DataSource, "SotienVATDH", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
                        txtMasoDN.DataBindings.Clear();
                        txtMasoDN.DataBindings.Add("Text", dataGridView1.DataSource, "MasoDN");
                        txtTenDNBH.DataBindings.Clear();
                        txtTenDNBH.DataBindings.Add("Text", dataGridView1.DataSource, "TenDNBH");
                        #endregion

                        this.Refresh();
                    }
                    else
                    {
                        //txtTongTK.Text = "0";
                        MessageBox.Show("No datarows to show.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dataGridView1.DataSource = null;
                    }
                }
                // Close connection
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                txtTungay.Text = null;
                txtDenngay.Text = null;
                txtRfdate.Text = null;
                txtRtdate.Text = null;
                txtSolanHT.Text = null;
                txtminSotienVNDHT.Text = null;
                txtFilePath.Text = null;
                rdoSimilar.Checked = true;
                //SetNumericFieldsToZero();               

                //ToggleDataGridViewLayout();
                RestoreDataGridView();
                LoadData();

                //RestoreDataGridView();
                //RestoreDataGridViewToOriginal();
                //RestoreToExactOriginalState();
                this.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Tracuu_Load(object sender, EventArgs e)
        {
            try
            {
                if (((Form)this.MdiParent).Controls["lblStatus"].Text != "True")
                {
                    btnBrowse.Enabled = false;
                    btnNhaplieu.Enabled = false;
                    txtFilePath.Enabled = false;
                    btnDel.Enabled = false;
                    btnUndo.Enabled = false;
                }
                chbMonthlyReport.Checked = true;
                rdoSimilar.Checked = true;
                rdoOpenOffice.Checked = true;
                dtpRfdate.Value = DateTime.Now;
                txtRfdate.Text = null;
                dtpRtdate.Value = DateTime.Now;
                txtRtdate.Text = null;
                txtSolanHT.Text = null;
                btnUndo.Enabled = false;
                txtminSotienVNDHT.Text = null;

                txtshowRows.Text = "100";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";

                this.pro_panel1.Visible = false;
                LoadData();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (rdoSimilar.Checked == true)
            {
                SimilarSearch();
            }
            else if (rdoNormal.Checked == true)
            {
                NormalSearch();
            }
            else if (rdoHighValue.Checked == true)
            {
                HighValueSearch();
            }
            else if (rdoRefundManyTimes.Checked == true)
            {
                RefundManyTimesSearch();
            }
            else
            {
                DupplicateSearch();
            }
        }

        private void DupplicateSearch()
        {
            string rfDate = txtRfdate.Text.Trim();
            string rtDate = txtRtdate.Text.Trim();

            if (string.IsNullOrWhiteSpace(rfDate) || string.IsNullOrWhiteSpace(rtDate))
            {
                MessageBox.Show("Please input tax refund date.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // Parse dates (expecting dd/MM/yyyy as elsewhere in the project)
            if (!DateTime.TryParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime prfDate)
                || !DateTime.TryParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime prtDate))
            {
                MessageBox.Show("Invalid date format. Use dd/MM/yyyy.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // Make end of day to include full rtDate
            var rfParam = prfDate.Date;
            var rtParam = prtDate.Date.AddDays(60).AddTicks(-1);

            // SQL finds keys with duplicates within the date range and returns full rows for those keys.
            string sql = @"
                WITH DuplicateKeys AS
                (
                    SELECT KyhieuHD, SoHD, MasoDN, SoHC
                    FROM BC09
                    WHERE NgayHT BETWEEN @rfdate AND @rtdate
                    GROUP BY KyhieuHD, SoHD, MasoDN, SoHC
                    HAVING COUNT(*) > 1
                )
                SELECT b.KyhieuHD, b.SoHD, b.NgayHD, b.NgayHT, b.TenDNBH, b.MasoDN, b.HoTenHK, b.SoHC, b.NgayHC,
                       b.Quoctich, b.TrigiaHHchuaVAT, b.SotienVATDH, b.SotienDVNHH
                FROM BC09 b
                INNER JOIN DuplicateKeys dk
                    ON b.KyhieuHD = dk.KyhieuHD
                    AND b.SoHD = dk.SoHD
                    AND b.MasoDN = dk.MasoDN
                    AND b.SoHC = dk.SoHC
                ORDER BY b.KyhieuHD, b.SoHD, b.MasoDN, b.SoHC, b.NgayHT;
            ";

            var parameters = new Dictionary<string, object>
            {
                { "@rfdate", rfParam },
                { "@rtdate", rtParam }
            };

            var pdgv = (ProgressDataGridView)dataGridView1;

            // One-time handler: unsubscribes itself after execution.
            RunWorkerCompletedEventHandler onComplete = null;
            onComplete = (s, e) =>
            {
                pdgv.SearchCompleted -= onComplete;

                // UI updates must run on UI thread
                BeginInvoke(new System.Action(() =>
                {
                    var dt = pdgv.DataSource as System.Data.DataTable;
                    if (dt == null || dt.Rows.Count == 0)
                    {
                        MessageBox.Show("No duplicate invoice found.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtTotalrows.Text = "0";
                        dataGridView1.DataSource = null;
                        return;
                    }

                    // Set grid and formatting (same as original)
                    dataGridView1.DataSource = dt;

                    DataGridViewCellStyle style = new DataGridViewCellStyle { Format = "N0" };

                    if (dataGridView1.Columns.Contains("TrigiaHHchuaVAT"))
                        dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle = style;
                    if (dataGridView1.Columns.Contains("SotienVATDH"))
                        dataGridView1.Columns["SotienVATDH"].DefaultCellStyle = style;
                    if (dataGridView1.Columns.Contains("SotienDVNHH"))
                        dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle = style;

                    if (dataGridView1.Columns.Contains("TrigiaHHchuaVAT"))
                        dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    if (dataGridView1.Columns.Contains("SotienVATDH"))
                        dataGridView1.Columns["SotienVATDH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    if (dataGridView1.Columns.Contains("SotienDVNHH"))
                        dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    SetupDataGridView();

                    // Totals
                    object tongsotienVATDH = dt.Compute("Sum(SotienVATDH)", string.Empty);
                    decimal tongsotienvatdh = 0;
                    decimal.TryParse(string.Format("{0}", tongsotienVATDH), out tongsotienvatdh);
                    txtTotalRefundAmt.Text = tongsotienvatdh.ToString("#,##0");

                    object tongtrigiaHHchuaVAT = dt.Compute("Sum(TrigiaHHchuaVAT)", string.Empty);
                    decimal tongtrigiahhchuavat = 0;
                    decimal.TryParse(string.Format("{0}", tongtrigiaHHchuaVAT), out tongtrigiahhchuavat);
                    txtTotalValue.Text = tongtrigiahhchuavat.ToString("#,##0");

                    object tongsotienDVNHH = dt.Compute("Sum(SotienDVNHH)", string.Empty);
                    decimal tongsotiendvnhh = 0;
                    decimal.TryParse(string.Format("{0}", tongsotienDVNHH), out tongsotiendvnhh);
                    txtTotalBankServiceFee.Text = tongsotiendvnhh.ToString("#,##0");

                    int totalPassenger = dt
                        .AsEnumerable()
                        .Select(r => r.Field<string>("SoHC"))
                        .Where(s => !string.IsNullOrEmpty(s))
                        .Distinct()
                        .Count();
                    txtTotalPassenger.Text = totalPassenger.ToString("#,##0");

                    txtTotalrows.Text = dt.Rows.Count.ToString("#,##0");

                    // Data binding
                    txtSoHC.DataBindings.Clear();
                    txtSoHC.DataBindings.Add("Text", dt, "SoHC");
                    txtSoHD.DataBindings.Clear();
                    txtSoHD.DataBindings.Add("Text", dt, "SoHD");
                    txtNgayHD.DataBindings.Clear();
                    txtNgayHD.DataBindings.Add("Text", dt, "NgayHD");
                    txtGoodsValue.DataBindings.Clear();
                    txtGoodsValue.DataBindings.Add("Text", dt, "TrigiaHHchuaVAT", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
                    txtRefundAmt.DataBindings.Clear();
                    txtRefundAmt.DataBindings.Add("Text", dt, "SotienVATDH", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
                    txtMasoDN.DataBindings.Clear();
                    txtMasoDN.DataBindings.Add("Text", dt, "MasoDN");
                    txtTenDNBH.DataBindings.Clear();
                    txtTenDNBH.DataBindings.Add("Text", dt, "TenDNBH");

                    this.Refresh();
                }));
            };

            pdgv.SearchCompleted += onComplete;

            // Execute async search (ProgressDataGridView will set DataSource when done)
            pdgv.DuplicateSearch(sql, parameters, true);
        }

        //private void DupplicateSearch()
        //{
        //    Utility ut = new Utility();
        //    var conn = ut.OpenDB();

        //    DateTime strhientai = DateTime.Now;
        //    string hientai = strhientai.ToString("yyyy-MM-dd hh:mm:ss tt");

        //    string rfDate = txtRfdate.Text.Trim();
        //    string rtDate = txtRtdate.Text.Trim();

        //    try
        //    {
        //        if (conn.State == ConnectionState.Closed)
        //        {
        //            conn.Open();

        //            if (string.IsNullOrWhiteSpace(rfDate) || string.IsNullOrWhiteSpace(rtDate))
        //            {
        //                MessageBox.Show("Please input tax refund date.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                return;
        //            }

        //            System.Data.DataTable dtbResult = new System.Data.DataTable();

        //            dtbResult.Columns.Add(new DataColumn("KyhieuHD", typeof(string)));
        //            dtbResult.Columns.Add(new DataColumn("SoHD", typeof(string)));
        //            dtbResult.Columns.Add(new DataColumn("NgayHD", typeof(DateTime)));
        //            dtbResult.Columns.Add(new DataColumn("NgayHT", typeof(DateTime)));
        //            dtbResult.Columns.Add(new DataColumn("TenDNBH", typeof(string)));

        //            dtbResult.Columns.Add(new DataColumn("MasoDN", typeof(string)));
        //            dtbResult.Columns.Add(new DataColumn("HoTenHK", typeof(string)));
        //            dtbResult.Columns.Add(new DataColumn("SoHC", typeof(string)));
        //            dtbResult.Columns.Add(new DataColumn("NgayHC", typeof(DateTime)));

        //            dtbResult.Columns.Add(new DataColumn("Quoctich", typeof(string)));
        //            dtbResult.Columns.Add(new DataColumn("TrigiaHHchuaVAT", typeof(decimal)));

        //            dtbResult.Columns.Add(new DataColumn("SotienVATDH", typeof(decimal)));

        //            dtbResult.Columns.Add(new DataColumn("SotienDVNHH", typeof(decimal)));
        //            //dtbResult.Columns.Add(new DataColumn("NgaygioHT", typeof(DateTime)));

        //            IFormatProvider provider = new CultureInfo("fr-FR");

        //            DateTime prfDate = DateTime.Parse(rfDate, provider);
        //            //DateTime prfDate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //            var rfdate = prfDate.ToString("yyyy-MM-dd hh:mm:ss tt");

        //            DateTime prtDate = DateTime.Parse(rtDate, provider);
        //            //DateTime prtDate = DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //            var rtdate = prtDate.ToString("yyyy-MM-dd hh:mm:ss tt");

        //            string query = @"Select Distinct KyhieuHD, SoHD, NgayHD, TenDNBH, MasoDN, HoTenHK, SoHC, Quoctich 
        //                            From BC09
        //                            Where (NgayHT between @rfdate and @rtdate)";

        //            using (SqlCommand cmd = new SqlCommand(query, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@rfdate", rfdate);
        //                cmd.Parameters.AddWithValue("@rtdate", rtdate);

        //                da = new SqlDataAdapter(cmd);
        //                System.Data.DataTable dt = new System.Data.DataTable();
        //                da.Fill(dt);
        //                cmd.Parameters.Clear();

        //                for (int i = 0; i < dt.Rows.Count; i++)
        //                {
        //                    DataRow r = dt.Rows[i];

        //                    string kyhieuHD = r["KyhieuHD"].ToString().Trim();
        //                    string soHD = Convert.ToString(r["SoHD"].ToString());
        //                    string ngayHD = r["NgayHD"].ToString().Trim();
        //                    string soHC = Convert.ToString(r["SoHC"].ToString());

        //                    DateTime pngayHD = DateTime.Parse(ngayHD, provider);
        //                    string strngayHD = pngayHD.ToString("yyyy-MM-dd hh:mm:ss tt");
        //                    DateTime pbeginDate = pngayHD.AddDays(0);
        //                    var strbeginDate = pbeginDate.ToString("yyyy-MM-dd hh:mm:ss tt");

        //                    DateTime pngayHD_add60 = pbeginDate.AddDays(60);
        //                    var strendingDate = pngayHD_add60.ToString("yyyy-MM-dd hh:mm:ss tt");

        //                    string fstrngayHD = pngayHD.ToString("dd-MM-yyyy");
        //                    string kyhieuSoNgay = (kyhieuHD + "/" + soHD + "/" + fstrngayHD).ToString().Trim();

        //                    // Old codes
        //                    SqlCommand cmd_kt = new SqlCommand(@"SELECT KyhieuHD, SoHD, NgayHD, NgayHT, TenDNBH, MasoDN, HoTenHK, SoHC, NgayHC, 
        //                            Quoctich, TrigiaHHchuaVAT, SotienVATDH, SotienDVNHH 
        //                            FROM BC09 
        //                            WHERE KyhieuHD = @kyhieuHD AND SoHD = @soHD AND NgayHD = @ngayHD AND SoHC = @soHC 
        //                            AND (NgayHT BETWEEN @beginDate AND @endDate)", conn);

        //                    //SqlCommand cmd_kt = new SqlCommand(@"SELECT bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.TenDNBH, bc09.MasoDN, bc09.HoTenHK, bc09.SoHC, 
        //                    //        bc09.NgayHC, bc09.Quoctich, bc09.TrigiaHHchuaVAT, 
        //                    //        bc09.NgayHT, bc09.SotienVATDH, bc09.SotienDVNHH, bc03.ThoigianGD 
        //                    //        FROM BC09 as bc09 
        //                    //        LEFT JOIN BC04 as bc04 ON bc09.SoHD = SUBSTRING(
        //                    //        bc04.KyhieuSoNgay, 
        //                    //        CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
        //                    //        CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1)
        //                    //        LEFT JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD                                   
        //                    //        WHERE bc09.KyhieuHD = @kyhieuHD AND bc09.SoHD = @soHD AND bc09.NgayHD = @ngayHD AND bc09.SoHC = @soHC
        //                    //        AND bc04.KyhieuSoNgay = @kyhieuSoNgay
        //                    //        AND (bc09.NgayHT BETWEEN @beginDate AND @endDate)", conn);

        //                    cmd_kt.Parameters.AddWithValue("@kyhieuHD", kyhieuHD);
        //                    cmd_kt.Parameters.AddWithValue("@soHD", soHD);
        //                    cmd_kt.Parameters.AddWithValue("@ngayHD", strngayHD);
        //                    cmd_kt.Parameters.AddWithValue("@soHC", soHC);
        //                    cmd_kt.Parameters.AddWithValue("@beginDate", strbeginDate);
        //                    cmd_kt.Parameters.AddWithValue("@endDate", strendingDate);
        //                    cmd_kt.Parameters.AddWithValue("@kyhieuSoNgay", kyhieuSoNgay);

        //                    da = new SqlDataAdapter(cmd_kt);
        //                    System.Data.DataTable dtb = new System.Data.DataTable();
        //                    da.Fill(dtb);
        //                    dtbResult.Merge(dtb);
        //                    cmd_kt.Parameters.Clear();
        //                }
        //            }

        //            if (dtbResult.Rows.Count > 0)
        //            {
        //                var grouped = dtbResult.AsEnumerable()
        //                .GroupBy(row => new
        //                {
        //                    KyhieuHD = row.Field<string>("KyhieuHD"),
        //                    SoHD = row.Field<string>("SoHD"),
        //                    //NgayHD = row.Field<DateTime>("NgayHD"),
        //                    //TenDNBH = row.Field<string>("TenDNBH"),
        //                    MasoDN = row.Field<string>("MasoDN"),
        //                    SoHC = row.Field<string>("SoHC")
        //                })
        //                .Where(g => g.Count() > 1)
        //                .SelectMany(g => g);

        //                System.Data.DataTable duplicates;
        //                if (grouped.Any())
        //                {
        //                    duplicates = grouped.CopyToDataTable();
        //                }
        //                else
        //                {
        //                    duplicates = dtbResult.Clone(); // empty table with same schema
        //                }

        //                if (duplicates.Rows.Count > 0)
        //                {
        //                    dataGridView1.DataSource = duplicates;

        //                    DataGridViewCellStyle style = new DataGridViewCellStyle();
        //                    style.Format = "N0";
        //                    this.dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle = style;
        //                    this.dataGridView1.Columns["SotienVATDH"].DefaultCellStyle = style;
        //                    this.dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle = style;

        //                    this.dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                    this.dataGridView1.Columns["SotienVATDH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                    this.dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                    SetupDataGridView();

        //                    object tongsotienVATDH, tongtrigiaHHchuaVAT, tongsotienDVNHH;

        //                    tongsotienVATDH = duplicates.Compute("Sum(SotienVATDH)", string.Empty);
        //                    decimal tongsotienvatdh = decimal.Parse(string.Format("{0}", tongsotienVATDH));
        //                    txtTotalRefundAmt.Text = Convert.ToString(tongsotienvatdh.ToString("#,##0"));

        //                    tongtrigiaHHchuaVAT = duplicates.Compute("Sum(TrigiaHHchuaVAT)", string.Empty);
        //                    decimal tongtrigiahhchuavat = decimal.Parse(string.Format("{0}", tongtrigiaHHchuaVAT));
        //                    txtTotalValue.Text = Convert.ToString(tongtrigiahhchuavat.ToString("#,##0"));

        //                    tongsotienDVNHH = duplicates.Compute("Sum(SotienDVNHH)", string.Empty);
        //                    decimal tongsotiendvnhh = decimal.Parse(string.Format("{0}", tongsotienDVNHH));
        //                    txtTotalBankServiceFee.Text = Convert.ToString(tongsotiendvnhh.ToString("#,##0"));

        //                    int totalPassenger = duplicates
        //                                .AsEnumerable()
        //                                .Select(r => r.Field<string>("SoHC"))
        //                                .Distinct()
        //                                .Count();
        //                    txtTotalPassenger.Text = Convert.ToString(totalPassenger.ToString("#,##0"));
        //                    // Counting total datarows in DB                        
        //                    txtTotalrows.Text = Convert.ToString(duplicates.Rows.Count.ToString("#,##0"));

        //                    //MessageBox.Show("There is/are" + "" + " " + duplicates.Rows.Count / 2 + " " + " pair(s) of dupplicated invoice(s).", "Notice", 
        //                    //    MessageBoxButtons.OK, MessageBoxIcon.Information);

        //                    //Binding data 
        //                    #region
        //                    txtSoHC.DataBindings.Clear();
        //                    txtSoHC.DataBindings.Add("Text", dataGridView1.DataSource, "SoHC");
        //                    txtSoHD.DataBindings.Clear();
        //                    txtSoHD.DataBindings.Add("Text", dataGridView1.DataSource, "SoHD");
        //                    txtNgayHD.DataBindings.Clear();
        //                    txtNgayHD.DataBindings.Add("Text", dataGridView1.DataSource, "NgayHD");
        //                    txtGoodsValue.DataBindings.Clear();
        //                    txtGoodsValue.DataBindings.Add("Text", dataGridView1.DataSource, "TrigiaHHchuaVAT", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
        //                    txtRefundAmt.DataBindings.Clear();
        //                    txtRefundAmt.DataBindings.Add("Text", dataGridView1.DataSource, "SotienVATDH", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
        //                    txtMasoDN.DataBindings.Clear();
        //                    txtMasoDN.DataBindings.Add("Text", dataGridView1.DataSource, "MasoDN");
        //                    txtTenDNBH.DataBindings.Clear();
        //                    txtTenDNBH.DataBindings.Add("Text", dataGridView1.DataSource, "TenDNBH");
        //                    #endregion
        //                    this.Refresh();
        //                }
        //                else
        //                {
        //                    MessageBox.Show("No dupplicate invoice found.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    txtTotalrows.Text = "0";
        //                    this.Refresh();
        //                }
        //            }
        //            // Close connection
        //            conn.Close();
        //            conn.Dispose();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //}


        private void DisplayTotals(System.Data.DataTable dt)
        {
            // Sum for SotienVATDH (Refund Amount)
            decimal tongsotienvatdh = dt.Compute("SUM(SotienVATDH)", string.Empty) as decimal? ?? 0;
            txtTotalRefundAmt.Text = tongsotienvatdh.ToString("#,##0");

            // Sum for TrigiaHHchuaVAT (Goods Value)
            decimal tongtrigiahhchuavat = dt.Compute("SUM(TrigiaHHchuaVAT)", string.Empty) as decimal? ?? 0;
            txtTotalValue.Text = tongtrigiahhchuavat.ToString("#,##0");

            // Sum for SotienDVNHH (Bank Service Fee)
            decimal tongsotiendvnhh = dt.Compute("SUM(SotienDVNHH)", string.Empty) as decimal? ?? 0;
            txtTotalBankServiceFee.Text = tongsotiendvnhh.ToString("#,##0");
        }

        // Helper function to bind controls to the data source
        private void BindControls(System.Data.DataTable dtSource)
        {
            // A simplified way to handle clearing and binding
            ClearDataBindings();

            if (dtSource != null)
            {
                // Re-binding logic remains similar, using the local 'dt' as the source
                txtSoHC.DataBindings.Add("Text", dtSource, "SoHC");
                txtSoHD.DataBindings.Add("Text", dtSource, "SoHD");
                txtGoodsValue.DataBindings.Add("Text", dtSource, "TrigiaHHchuaVAT", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
                txtRefundAmt.DataBindings.Add("Text", dtSource, "SotienVATDH", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
                txtMasoDN.DataBindings.Add("Text", dtSource, "MasoDN");
                txtTenDNBH.DataBindings.Add("Text", dtSource, "TenDNBH");
            }
        }

        //Helper function to clear all data bindings on the specific controls
        private void ClearDataBindings()
        {
            txtSoHC.DataBindings.Clear();
            txtSoHD.DataBindings.Clear();
            txtGoodsValue.DataBindings.Clear();
            txtRefundAmt.DataBindings.Clear();
            txtMasoDN.DataBindings.Clear();
            txtTenDNBH.DataBindings.Clear();
        }

        // final RefundManyTimesSearch() function
        private void RefundManyTimesSearch()
        {
            try
            {
                // 1. Input and Data Initialization
                string rfDateText = txtRfdate.Text.Trim();
                string rtDateText = txtRtdate.Text.Trim();
                string solanHTText = txtSolanHT.Text.Trim();

                if (string.IsNullOrWhiteSpace(rfDateText) && string.IsNullOrWhiteSpace(rtDateText))
                {
                    MessageBox.Show("Please input period of search.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Try parsing the number of refunds
                if (!int.TryParse(solanHTText, out int solanHT))
                {
                    MessageBox.Show("Invalid number of refunds (SolanHT).", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Try parsing the From Date (rfDate)
                if (!DateTime.TryParseExact(rfDateText, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime rfDate))
                {
                    MessageBox.Show("Invalid 'From Date' format. Please use dd/MM/yyyy.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Try parsing the To Date (rtDate) if provided
                DateTime? rtDate = null;
                if (!string.IsNullOrWhiteSpace(rtDateText))
                {
                    if (!DateTime.TryParseExact(rtDateText, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedRtDate))
                    {
                        MessageBox.Show("Invalid 'To Date' format. Please use dd/MM/yyyy.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    rtDate = parsedRtDate;
                }

                // The date to use for the end of the range. If rtDate is null, use DateTime.Now.
                DateTime endDate = rtDate ?? DateTime.Now;

                // Build the SQL query

                string sqlQuery = @"WITH DailyCounts AS (
                                SELECT 
                                    CAST(NgayHT AS DATE) AS ExecutionDate,
                                    SoHC,
                                    COUNT(DISTINCT SoHC) AS DailyOccurrence
                                FROM [dbo].[BC09]
                                WHERE NgayHT BETWEEN @rfdate AND @endDate 
                                  AND SoHC IS NOT NULL
                                GROUP BY CAST(NgayHT AS DATE), SoHC
                            )                           
                            SELECT    
                                DISTINCT SoHC,    
                                SUM(DailyOccurrence) OVER(PARTITION BY SoHC) AS TotalOccurrenceInPeriod
                            INTO #TempDestinationTable
                            FROM DailyCounts;                            

                            SELECT b.SoHD, b.KyhieuHD, b.NgayHD, b.NgayHT, b.HoTenHK, b.SoHC, b.NgayHC, b.Quoctich,
                                   b.TrigiaHHchuaVAT, b.SotienVATDH, b.SotienDVNHH, 
                                   b.MasoDN, b.TenDNBH, b.Ghichu
                            FROM BC09 as b 
                            INNER JOIN #TempDestinationTable as tmpt ON b.SoHC = tmpt.SoHC
                            WHERE tmpt.TotalOccurrenceInPeriod >= @solanHT AND (b.NgayHT BETWEEN @rfdate AND @endDate)
                            ORDER BY b.SoHC;

                            DROP TABLE #TempDestinationTable;";

                // Old codes
                #region
                //string sqlQuery = @"SELECT SoHC, COUNT(SoHC) as SolanHT
                //            INTO #TempDestinationTable
                //            FROM BC09
                //            WHERE NgayHT BETWEEN @rfdate AND @endDate
                //            GROUP BY SoHC;

                //            SELECT b.SoHD, b.NgayHD, b.KyhieuHD, b.HoTenHK, b.SoHC, b.NgayHC, b.Quoctich,
                //                   b.NgayHT, b.TrigiaHHchuaVAT, b.SotienVATDH, b.SotienDVNHH, 
                //                   b.MasoDN, b.TenDNBH, b.Ghichu
                //            FROM BC09 as b 
                //            INNER JOIN #TempDestinationTable as tmpt ON b.SoHC = tmpt.SoHC
                //            WHERE tmpt.SolanHT >= @solanHT AND (b.NgayHT BETWEEN @rfdate AND @endDate)
                //            ORDER BY b.SoHC;

                //            DROP TABLE #TempDestinationTable;";
                #endregion

                // Create parameters dictionary
                Dictionary<string, object> parameters = new Dictionary<string, object>
            {
                { "@rfdate", rfDate },
                { "@endDate", endDate },
                { "@solanHT", solanHT }
            };

                // Execute the search using ProgressDataGridView
                ((ProgressDataGridView)dataGridView1).RefundManyTimesSearch(sqlQuery, parameters, true);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ProgressDataGridView_SearchCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show($"Error during search: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ClearAllData();
                return;
            }

            try
            {
                System.Data.DataTable dt = ((ProgressDataGridView)dataGridView1).DataSource as System.Data.DataTable;

                if (dt != null && dt.Rows.Count > 0)
                {
                    ProcessAndDisplayResults(dt);
                }
                else
                {
                    MessageBox.Show("No datarows to show.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ClearAllData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ClearAllData();
            }
        }

        private void ProcessAndDisplayResults(System.Data.DataTable dt)
        {
            // Add STT column (Serial Number)
            dt.Columns.Add("STT", typeof(int));
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["STT"] = i + 1;
            }

            // Set up DataGridView
            dataGridView1.DataSource = dt;
            dataGridView1.Columns["STT"].DisplayIndex = 0;

            SetupDataGridView();

            // Calculate Totals and display
            DisplayTotals(dt);

            // Count unique SoHC
            int totalPassenger = dt.AsEnumerable()
                                   .Select(r => r.Field<string>("SoHC"))
                                   .Where(sohc => !string.IsNullOrEmpty(sohc))
                                   .Distinct()
                                   .Count();
            txtTotalPassenger.Text = totalPassenger.ToString("#,##0");

            txtTotalrows.Text = dt.Rows.Count.ToString("#,##0");

            this.Refresh();

            // Binding data
            BindControls(dt);
        }

        private void ClearAllData()
        {
            txtTotalrows.Text = "0";
            dataGridView1.DataSource = null;

            // Clear total fields
            txtTotalRefundAmt.Text = "0";
            txtTotalValue.Text = "0";
            txtTotalBankServiceFee.Text = "0";
            txtTotalPassenger.Text = "0";

            // Clear data bindings
            ClearDataBindings();
        }

        private void NormalSearch()
        {
            try
            {
                // Validate required input
                string rfDate = txtRfdate.Text.Trim();
                if (string.IsNullOrWhiteSpace(rfDate))
                {
                    MessageBox.Show("Please input tax refund date.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Get input values
                string soHC = txtSoHC.Text.Trim();
                string soHD = txtSoHD.Text.Trim();
                string tuNgay = txtTungay.Text.Trim();
                string denNgay = txtDenngay.Text.Trim();
                string rtDate = txtRtdate.Text.Trim();
                string masoDN = txtMasoDN.Text.Trim();

                // Prepare date parameters
                DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                string formattedRfDate = prfdate.ToString("yyyy-MM-dd HH:mm:ss");
                string formattedCurrentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string formattedRtDate = string.IsNullOrWhiteSpace(rtDate)
                    ? formattedCurrentDate
                    : DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd HH:mm:ss");

                // Build base query
                string commandText = @"
                        SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
                               bc09.SoHD, bc09.KyhieuHD, bc09.NgayHD, bc09.NgayHT, bc09.HoTenHK, 
                               bc09.SoHC, bc09.NgayHC, bc09.Quoctich,  
                               bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH, 
                               bc09.MasoDN, bc09.TenDNBH, bc09.Ghichu, bc09.NgaynhapHT, bc09.LoginName 
                        FROM BC09 as bc09 
                        WHERE bc09.NgayHT BETWEEN @RfDate AND @RtDate";

                // Initialize parameters dictionary
                var parameters = new Dictionary<string, object>
                    {
                        { "@RfDate", formattedRfDate },
                        { "@RtDate", formattedRtDate }
                    };

                // Add conditions and parameters dynamically
                if (!string.IsNullOrWhiteSpace(soHC))
                {
                    commandText += " AND bc09.SoHC = @SoHC";
                    parameters.Add("@SoHC", soHC);
                }

                if (!string.IsNullOrWhiteSpace(soHD))
                {
                    commandText += " AND bc09.SoHD = @SoHD";
                    parameters.Add("@SoHD", soHD);
                }

                if (!string.IsNullOrWhiteSpace(masoDN))
                {
                    commandText += " AND bc09.MasoDN = @MasoDN";
                    parameters.Add("@MasoDN", masoDN);
                }

                // Handle invoice date range
                if (!string.IsNullOrWhiteSpace(tuNgay) || !string.IsNullOrWhiteSpace(denNgay))
                {
                    if (string.IsNullOrWhiteSpace(tuNgay) || string.IsNullOrWhiteSpace(denNgay))
                    {
                        MessageBox.Show("Please input both from and to invoice dates.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    DateTime ptuNgay = DateTime.ParseExact(tuNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    DateTime pdenNgay = DateTime.ParseExact(denNgay, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                    commandText += " AND bc09.NgayHD BETWEEN @TuNgay AND @DenNgay";
                    parameters.Add("@TuNgay", ptuNgay.ToString("yyyy-MM-dd HH:mm:ss"));
                    parameters.Add("@DenNgay", pdenNgay.ToString("yyyy-MM-dd HH:mm:ss"));
                }

                // Execute query through NormalSeach
                ((ProgressDataGridView)dataGridView1).NormalSearch(commandText, parameters, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SimilarSearch()
        {
            string soHC = txtSoHC.Text.Trim();
            string soHD = txtSoHD.Text.Trim();
            string masoDN = txtMasoDN.Text.Trim();

            try
            {
                // Validate input first
                if (string.IsNullOrWhiteSpace(soHC) && string.IsNullOrWhiteSpace(soHD) && string.IsNullOrWhiteSpace(masoDN))
                {
                    MessageBox.Show("No datarows. Please check input data.", "Notice",
                                  MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Build base query
                string commandText = @"SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHD DESC) AS [STT], 
                                bc09.SoHD, bc09.KyhieuHD, bc09.NgayHD, bc09.NgayHT,  
                                bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, 
                                bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH, bc09.MasoDN, bc09.TenDNBH, 
                                bc09.Ghichu, bc09.NgaynhapHT, bc09.LoginName  
                                FROM BC09 as bc09 
                                WHERE 1=1";

                // Create parameters dictionary
                var parameters = new Dictionary<string, object>();

                // Add conditions based on input
                if (!string.IsNullOrWhiteSpace(soHC))
                {
                    commandText += " AND bc09.SoHC LIKE '%' + @soHC + '%'";
                    parameters.Add("@soHC", soHC);
                }

                if (!string.IsNullOrWhiteSpace(soHD))
                {
                    commandText += " AND bc09.SoHD LIKE '%' + @soHD + '%'";
                    parameters.Add("@soHD", soHD);
                }

                if (!string.IsNullOrWhiteSpace(masoDN))
                {
                    commandText += " AND bc09.MasoDN LIKE '%' + @masoDN + '%'";
                    parameters.Add("@masoDN", masoDN);
                }

                // Execute query through ProgressDataGridView's SimilarSearch method
                ((ProgressDataGridView)dataGridView1).SimilarSearch(commandText, parameters, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //private void HighValueSearch()
        //{
        //    Utility ut = new Utility();
        //    var conn = ut.OpenDB();

        //    try
        //    {
        //        string rfDate = txtRfdate.Text.Trim();
        //        string rtDate = txtRtdate.Text.Trim();
        //        DateTime strhientai = DateTime.Now;
        //        string hientai = strhientai.ToString("yyyy-MM-dd hh:mm:ss tt");
        //        string[] strsotienVNDHT = txtminsotienVNDHT.Text.Split(',');
        //        string sotienVNDHT = strsotienVNDHT[0].Replace(".", "");

        //        if (conn.State == ConnectionState.Closed)
        //        {
        //            conn.Open();

        //            if (string.IsNullOrWhiteSpace(sotienVNDHT))
        //            {
        //                MessageBox.Show("No datarows. Please input refund amount again.", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
        //            }
        //            else
        //            {
        //                if (string.IsNullOrWhiteSpace(rfDate) && string.IsNullOrWhiteSpace(rtDate))
        //                {
        //                    cmd = new SqlCommand("SELECT DISTINCT bc04.KyhieuSoNgay, bc03.ThoigianGD as NgayHT, bc03.MasoGD, bc03.HotenHK, bc03.SotienVNDHT " +
        //                    "FROM BC03 as bc03 " +
        //                    "INNER JOIN BC04 as bc04 ON bc03.MasoGD = bc04.MasoGD " +
        //                    "WHERE bc03.SotienVNDHT > N'" + sotienVNDHT + "' ", conn);
        //                }
        //                else if (string.IsNullOrWhiteSpace(rtDate))
        //                {
        //                    DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //                    var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");

        //                    cmd = new SqlCommand("SELECT DISTINCT bc04.KyhieuSoNgay, bc03.ThoigianGD as NgayHT, bc03.MasoGD, bc03.HotenHK, bc03.SotienVNDHT " +
        //                   "FROM BC03 as bc03 " +
        //                   "INNER JOIN BC04 as bc04 ON bc03.MasoGD = bc04.MasoGD " +
        //                   "WHERE bc03.SotienVNDHT > N'" + sotienVNDHT + "' and (NgayHT between '" + rfdate + "' and '" + hientai + "')", conn);
        //                }
        //                else
        //                {
        //                    DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //                    var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");
        //                    DateTime prtdate = DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //                    var rtdate = prtdate.ToString("yyyy-MM-dd hh:mm:ss tt");

        //                    cmd = new SqlCommand("SELECT DISTINCT bc04.KyhieuSoNgay, bc03.ThoigianGD as NgayHT, bc03.MasoGD, bc03.HotenHK, bc03.SotienVNDHT " +
        //                  "FROM BC03 as bc03 " +
        //                  "INNER JOIN BC04 as bc04 ON bc03.MasoGD = bc04.MasoGD " +
        //                  "WHERE bc03.SotienVNDHT > N'" + sotienVNDHT + "' and (NgayHT Between '" + rfdate + "' and '" + rtdate + "')", conn);
        //                }

        //                da = new SqlDataAdapter(cmd);
        //                System.Data.DataTable dt_kt = new System.Data.DataTable();
        //                da.Fill(dt_kt);
        //                //dgvNhaplieuExcel.DataSource = dt_kt;

        //                if (dt_kt.Rows.Count > 0)
        //                {
        //                    dt = new System.Data.DataTable();
        //                    IFormatProvider provider = new CultureInfo("fr-FR");

        //                    for (int i = 0; i < dt_kt.Rows.Count; i++)
        //                    {
        //                        DataRow r = dt_kt.Rows[i];

        //                        string kyhieuSoNgay = r[0].ToString().Trim();

        //                        string[] listwordFields = kyhieuSoNgay.Split('/');

        //                        string kyhieuHD = listwordFields[0].ToString();

        //                        if (listwordFields.Count() == 3)
        //                        {
        //                            string soHD = listwordFields[1].ToString();
        //                            string ngayHD = listwordFields[2].ToString();
        //                            //IFormatProvider provider = new CultureInfo("fr-FR");

        //                            DateTime pngayHD = DateTime.Parse(ngayHD, provider);
        //                            var strngayHD = pngayHD.ToString("yyyy-MM-dd hh:mm:ss tt");

        //                            cmd = new SqlCommand("Select bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.TenDNBH, bc09.MasoDN, " +
        //                                    "bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH " +
        //                                    "From BC09 as bc09 " +
        //                                    "Where bc09.SoHD = N'" + soHD + "' And bc09.KyhieuHD = N'" + kyhieuHD + "' And bc09.NgayHD = N'" + strngayHD + "' " +
        //                                    "And bc09.SotienVATDH  > N'" + sotienVNDHT + "'", conn);
        //                        }
        //                        else if (listwordFields.Count() == 2)
        //                        {
        //                            string soHD = listwordFields[1].ToString();
        //                            string ngayHD = DateTime.Now.ToString();

        //                            DateTime pngayHD = DateTime.Parse(ngayHD, provider);
        //                            var strngayHD = pngayHD.ToString("yyyy-MM-dd hh:mm:ss tt");

        //                            cmd = new SqlCommand("Select bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.TenDNBH, bc09.MasoDN, " +
        //                                    "bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH " +
        //                                    "From BC09 as bc09 " +
        //                                    "Where bc09.SoHD = N'" + soHD + "' And bc09.KyhieuHD = N'" + kyhieuHD + "' And bc09.NgayHD = N'" + strngayHD + "' " +
        //                                    "And bc09.SotienVATDH  > N'" + sotienVNDHT + "'", conn);
        //                        }
        //                        else
        //                        {
        //                            string soHD = listwordFields[2].ToString();
        //                            string ngayHD = listwordFields[3].ToString();

        //                            DateTime pngayHD = DateTime.Parse(ngayHD, provider);
        //                            var strngayHD = pngayHD.ToString("yyyy-MM-dd hh:mm:ss tt");

        //                            cmd = new SqlCommand("Select bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.TenDNBH, bc09.MasoDN, " +
        //                                    "bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH " +
        //                                    "From BC09 as bc09 " +
        //                                    "Where bc09.SoHD = N'" + soHD + "' and bc09.NgayHD = N'" + strngayHD + "' " +
        //                                    "And bc09.SotienVATDH  > N'" + sotienVNDHT + "'", conn);
        //                        }

        //                        da = new SqlDataAdapter(cmd);
        //                        da.Fill(dt);
        //                    }

        //                    if (dt.Rows.Count > 0)
        //                    {
        //                        dt.Columns.Add("STT");
        //                        for (int i = 0; i < dt.Rows.Count; i++)
        //                        {
        //                            dt.Rows[i]["STT"] = i + 1;
        //                        }
        //                        dataGridView1.DataSource = dt;
        //                        dataGridView1.Columns["STT"].DisplayIndex = 0;

        //                        DataGridViewCellStyle style = new DataGridViewCellStyle();
        //                        style.Format = "N0";
        //                        this.dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle = style;
        //                        this.dataGridView1.Columns["SotienVATDH"].DefaultCellStyle = style;
        //                        this.dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle = style;

        //                        this.dataGridView1.Columns["TrigiaHHchuaVAT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                        this.dataGridView1.Columns["SotienVATDH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                        this.dataGridView1.Columns["SotienDVNHH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        //                        SetupDataGridView();

        //                        //Data binding
        //                        txtSoHC.DataBindings.Clear();
        //                        txtSoHC.DataBindings.Add("Text", dataGridView1.DataSource, "SoHC");
        //                        txtSoHD.DataBindings.Clear();
        //                        txtSoHD.DataBindings.Add("Text", dataGridView1.DataSource, "SoHD");
        //                        txtGoodsValue.DataBindings.Clear();
        //                        txtGoodsValue.DataBindings.Add("Text", dataGridView1.DataSource, "TrigiaHHchuaVAT", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");
        //                        txtRefundAmt.DataBindings.Clear();
        //                        txtRefundAmt.DataBindings.Add("Text", dataGridView1.DataSource, "SotienVATDH", true, DataSourceUpdateMode.OnPropertyChanged, string.Empty, "#,##0");

        //                        txtMasoDN.DataBindings.Clear();
        //                        txtMasoDN.DataBindings.Add("Text", dataGridView1.DataSource, "MasoDN");
        //                        txtTenDNBH.DataBindings.Clear();
        //                        txtTenDNBH.DataBindings.Add("Text", dataGridView1.DataSource, "TenDNBH");

        //                        object tongsotienVATDH, tongtrigiaHHchuaVAT, tongsotienDVNHH;

        //                        tongsotienVATDH = dt.Compute("Sum(SotienVATDH)", string.Empty);
        //                        decimal tongsotienvatdh = decimal.Parse(string.Format("{0}", tongsotienVATDH));
        //                        txtTotalRefundAmt.Text = Convert.ToString(tongsotienvatdh.ToString("#,##0"));

        //                        tongtrigiaHHchuaVAT = dt.Compute("Sum(TrigiaHHchuaVAT)", string.Empty);
        //                        decimal tongtrigiahhchuavat = decimal.Parse(string.Format("{0}", tongtrigiaHHchuaVAT));
        //                        txtTotalValue.Text = Convert.ToString(tongtrigiahhchuavat.ToString("#,##0"));

        //                        tongsotienDVNHH = dt.Compute("Sum(SotienDVNHH)", string.Empty);
        //                        decimal tongsotiendvnhh = decimal.Parse(string.Format("{0}", tongsotienDVNHH));
        //                        txtTotalBankServiceFee.Text = Convert.ToString(tongsotiendvnhh.ToString("#,##0"));

        //                        int totalPassenger = dt
        //                                    .AsEnumerable()
        //                                    .Select(r => r.Field<string>("SoHC"))
        //                                    .Distinct()
        //                                    .Count();
        //                        txtTotalPassenger.Text = Convert.ToString(totalPassenger.ToString("#,##0"));
        //                        //Count total datarows in DB                        
        //                        txtTotalrows.Text = Convert.ToString(dt.Rows.Count.ToString("#,##0"));
        //                        this.Refresh();
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("No datarows to show.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                        dataGridView1.DataSource = null;
        //                    }
        //                }
        //            }
        //        }
        //        //Dong ket noi
        //        conn.Close();
        //        conn.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //}

        private void HighValueSearch()
        {
            try
            {
                string rfDate = txtRfdate.Text.Trim();
                string rtDate = txtRtdate.Text.Trim();
                DateTime strhientai = DateTime.Now;
                string hientai = strhientai.ToString("yyyy-MM-dd hh:mm:ss tt");
                string[] strsotienVNDHT = txtminSotienVNDHT.Text.Split(',');
                string sotienVNDHT = strsotienVNDHT[0].Replace(".", "");

                if (string.IsNullOrWhiteSpace(sotienVNDHT))
                {
                    MessageBox.Show("No datarows. Please input refund amount again.", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    return;
                }

                // Build the appropriate SQL command based on date inputs
                string commandText;
                Dictionary<string, object> parameters = new Dictionary<string, object>();

                if (string.IsNullOrWhiteSpace(rfDate) && string.IsNullOrWhiteSpace(rtDate))
                {
                    commandText = @"SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT) AS [STT], 
                                    bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.NgayHT, 
                                    bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.TrigiaHHchuaVAT, 
                                    bc09.SotienVATDH, bc09.SotienDVNHH, bc09.MasoDN, bc09.TenDNBH 
                                    FROM BC09 as bc09      
                                    WHERE bc09.SotienVATDH > @sotienVNDHT 
                                    ORDER BY bc09.NgayHT";
                }
                else if (string.IsNullOrWhiteSpace(rtDate))
                {
                    DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");

                    commandText = @"SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT) AS [STT], 
                                    bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.NgayHT, 
                                    bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.TrigiaHHchuaVAT, 
                                    bc09.SotienVATDH, bc09.SotienDVNHH, bc09.MasoDN, bc09.TenDNBH  
                                    FROM BC09 as bc09      
                                    WHERE bc09.SotienVATDH > @sotienVNDHT and (NgayHT between @rfdate and @hientai) ORDER BY bc09.NgayHT";

                    parameters.Add("@rfdate", rfdate);
                    parameters.Add("@hientai", hientai);
                }
                else
                {
                    DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");
                    DateTime prtdate = DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    var rtdate = prtdate.ToString("yyyy-MM-dd hh:mm:ss tt");

                    commandText = @"SELECT DISTINCT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT) AS [STT], 
                                    bc09.KyhieuHD, bc09.SoHD, bc09.NgayHD, bc09.NgayHT, 
                                    bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.TrigiaHHchuaVAT, 
                                    bc09.SotienVATDH, bc09.SotienDVNHH, bc09.MasoDN, bc09.TenDNBH  
                                    FROM BC09 as bc09      
                                    WHERE bc09.SotienVATDH > @sotienVNDHT and (NgayHT Between @rfdate and @rtdate) ORDER BY bc09.NgayHT";

                    parameters.Add("@rfdate", rfdate);
                    parameters.Add("@rtdate", rtdate);
                }

                parameters.Add("@sotienVNDHT", sotienVNDHT);

                // Execute query through ProgressDataGridView's HighValueSearch method
                ((ProgressDataGridView)dataGridView1).HighValueSearch(commandText, parameters, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DataGridSearch_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Get the DataTable from the DataGridView
            var dt = dataGridView1.DataSource as System.Data.DataTable;

            if (dt != null && dt.Rows.Count > 0)
            {
                SetupDataGridView();
                SetupDataBindings();
                CalculateAndDisplayTotals(dt);
            }
            else
            {
                ClearResultFields();
            }
        }

        private void CalculateAndDisplayTotals(System.Data.DataTable dt)
        {
            // Calculate and display totals
            decimal tongsotienvatdh = Convert.ToDecimal(dt.Compute("Sum(SotienVATDH)", string.Empty));
            txtTotalRefundAmt.Text = tongsotienvatdh.ToString("#,##0");

            decimal tongtrigiahhchuavat = Convert.ToDecimal(dt.Compute("Sum(TrigiaHHchuaVAT)", string.Empty));
            txtTotalValue.Text = tongtrigiahhchuavat.ToString("#,##0");

            decimal tongsotiendvnhh = Convert.ToDecimal(dt.Compute("Sum(SotienDVNHH)", string.Empty));
            txtTotalBankServiceFee.Text = tongsotiendvnhh.ToString("#,##0");

            //int totalPassenger = dt.AsEnumerable()
            //                    .Select(r => r.Field<string>("SoHC"))
            //                    .Where(x => !string.IsNullOrWhiteSpace(x))
            //                    .Distinct()
            //                    .Count();

            // 1. Group by Date and SoHC to get unique passports per day
            var dailyUniqueCount = dt.AsEnumerable()
                .Where(row => row["SoHC"] != DBNull.Value && row["NgayHT"] != DBNull.Value)
                .Select(row => new
                {
                    Date = Convert.ToDateTime(row["NgayHT"]).Date,
                    Passport = row["SoHC"].ToString()
                })
                .Distinct() // This ensures we only count a passport once per day
                .Count();   // This gives you the final total number

            txtTotalPassenger.Text = (+dailyUniqueCount).ToString("#,##0");

            // Count total datarows                        
            txtTotalrows.Text = dataGridView1.Rows.Count.ToString("#,##0");
        }

        private void SimilarSeach()
        {
            string soHC = txtSoHC.Text.Trim();
            string soHD = txtSoHD.Text.Trim();

            try
            {
                // Validate input first
                if (string.IsNullOrWhiteSpace(soHC) && string.IsNullOrWhiteSpace(soHD))
                {
                    MessageBox.Show("No datarows. Please check again.", "Notice",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    return;
                }

                // Build base query
                string commandText = @"SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHD DESC) AS [STT], 
                  bc09.SoHD, bc09.KyhieuHD, bc09.NgayHD, bc09.NgayHT, bc09.TenDNBH, bc09.MasoDN, 
                  bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, 
                  bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH, bc09.Ghichu, bc09.NgaynhapHT, bc09.LoginName  
                  FROM BC09 as bc09 
                  WHERE 1=1";

                // Create parameters dictionary
                var parameters = new Dictionary<string, object>();

                // Add conditions based on input
                if (!string.IsNullOrWhiteSpace(soHC))
                {
                    commandText += " AND bc09.SoHC LIKE '%' + @soHC + '%'";
                    parameters.Add("@soHC", soHC);
                }

                if (!string.IsNullOrWhiteSpace(soHD))
                {
                    commandText += " AND bc09.SoHD LIKE '%' + @soHD + '%'";
                    parameters.Add("@soHD", soHD);
                }

                // Execute query through ProgressDataGridView's SimilarSeach method
                ((ProgressDataGridView)dataGridView1).SimilarSearch(commandText, parameters, true);

                SetupDataGridView();
                //SetupDataBindings();
                // Handle the results when DataThread_RunWorkerCompleted is called
                // Note: The actual data binding will happen automatically through the DataSource setting
                // in the DataThread_DoWork method of ProgressDataGridView
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void dtpTungay_ValueChanged(object sender, EventArgs e)
        {
            if (txtTungay.Text != null)
            {
                txtTungay.Text = dtpTungay.Value.ToString("dd/MM/yyyy");
                string strtungay = dtpTungay.Text;
                DateTime ptungay = DateTime.ParseExact(strtungay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var tungay = ptungay.ToString("yyyy-MM-dd");
            }
            else
            {
                DateTime strngayhientai = DateTime.Now;
                string ngayhientai = strngayhientai.ToString("yyyy-MM-dd hh:mm:ss tt");
                txtTungay.Text = strngayhientai.ToString("dd/MM/yyyy");
            }
        }

        private void dtpDenngay_ValueChanged(object sender, EventArgs e)
        {
            if (txtDenngay.Text != null)
            {
                txtDenngay.Text = dtpDenngay.Value.ToString("dd/MM/yyyy");
                string strdenngay = dtpDenngay.Text;
                DateTime pdenngay = DateTime.ParseExact(strdenngay, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                var denngay = pdenngay.ToString("yyyy-MM-dd");
            }
            else
            {
                DateTime strngayhientai = DateTime.Now;
                string ngayhientai = strngayhientai.ToString("yyyy-MM-dd hh:mm:ss tt");
                txtDenngay.Text = strngayhientai.ToString("dd/MM/yyyy");
            }
        }

        private void dtpRfdate_ValueChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    //if (txtRfdate.Text != null)
            //    //{
            //    //    txtRfdate.Text = dtpRfdate.Value.ToString("dd/MM/yyyy");
            //    //    string strrfdate = dtpRfdate.Text;
            //    //    DateTime prfdate = DateTime.ParseExact(strrfdate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            //    //    var rfdate = prfdate.ToString("yyyy-MM-dd");
            //    //}
            //    //else
            //    //{
            //    //    DateTime strngayhientai = DateTime.Now;
            //    //    string ngayhientai = strngayhientai.ToString("yyyy-MM-dd hh:mm:ss tt");
            //    //    txtRfdate.Text = strngayhientai.ToString("dd/MM/yyyy");
            //    //}

            //    // Direct assignment without unnecessary conversions
            //    txtRfdate.Text = dtpRfdate.Value.ToString("dd/MM/yyyy");

            //    // If you need the date in different formats, use the Value property directly
            //    DateTime selectedDate = dtpRfdate.Value;
            //    string displayFormat = selectedDate.ToString("dd/MM/yyyy");
            //    string databaseFormat = selectedDate.ToString("yyyy-MM-dd");

            //    // Use the formats as needed
            //    Console.WriteLine($"Display: {displayFormat}, Database: {databaseFormat}");
            //}
            ////catch (Exception ex)
            ////{
            ////    MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            ////}
            //catch (FormatException fex)
            //{
            //    MessageBox.Show($"Date format error: {fex.Message}", "Format Error",
            //                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"Error: {ex.Message}", "Notice",
            //                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
        }
        // Method to reset to current date
        private void ResetToCurrentDate()
        {
            dtpRfdate.Value = DateTime.Now;
            txtRfdate.Text = DateTime.Now.ToString("dd/MM/yyyy");
        }

        private void dtpRtdate_ValueChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    if (txtRtdate.Text != null)
            //    {
            //        txtRtdate.Text = dtpRtdate.Value.ToString("dd/MM/yyyy");
            //        string strrtdate = dtpRtdate.Text;
            //        DateTime prtdate = DateTime.ParseExact(strrtdate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            //        var rtdate = prtdate.ToString("yyyy-MM-dd");
            //    }
            //    else
            //    {
            //        DateTime strngayhientai = DateTime.Now;
            //        //var pngayhientai = DateTime.ParseExact(strngayhientai, "dd/MM/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);
            //        string ngayhientai = strngayhientai.ToString("yyyy-MM-dd hh:mm:ss tt");
            //        txtRtdate.Text = strngayhientai.ToString("dd/MM/yyyy");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
        }

        // Method to set a specific date
        private void SetSpecificDate(DateTime date)
        {
            dtpRfdate.Value = date;
            txtRfdate.Text = date.ToString("dd/MM/yyyy");
        }

        private void ClearTextFields()
        {
            txtminSotienVNDHT.Text = string.Empty;
            txtNgayHD.Text = string.Empty;
            txtSoHC.Text = string.Empty;
            txtMasoDN.Text = string.Empty;
            txtTenDNBH.Text = string.Empty;
        }

        private void SetNumericFieldsToZero()
        {
            txtTotalRefundAmt.Text = "0";
            txtTotalValue.Text = "0";
            txtTotalPassenger.Text = "0";
            txtRefundAmt.Text = "0";
            txtGoodsValue.Text = "0";
            txtTotalBankServiceFee.Text = "0";
        }
        private void rdoManyTimes_CheckedChanged(object sender, EventArgs e)
        {

            if (rdoRefundManyTimes.Checked)
            {
                txtSolanHT.Text = "10";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalPassenger.Text = "0";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";

                txtSoHC.Text = null;
                txtminSotienVNDHT.Text = null;
                txtNgayHD.Text = null;
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
            }
            else
            {
                txtSolanHT.Text = null;
            }

            //if (rdoRefundManyTimes.Checked)
            //{
            //    // Set default values when checked
            //    txtSolanHT.Text = "10";
            //    ClearTextFields();
            //    SetNumericFieldsToZero();
            //    this.Refresh();
            //}
            //else
            //{
            //    // Only clear this field when unchecked
            //    txtSolanHT.Text = null;
            //}
        }

        private void rdoHighValue_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoHighValue.Checked)
            {
                txtminSotienVNDHT.Text = "100000000";
                FormatTextBoxForDisplay(txtminSotienVNDHT);
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalPassenger.Text = "0";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                txtSolanHT.Text = null;
            }
            else
            {
                //txtsotienVNDHT.Text = string.Format("{0:#,##0.00}", 0, 00);
                txtminSotienVNDHT.Text = null;
            }
        }

        // Create a helper method for formatting
        private void FormatTextBoxForDisplay(System.Windows.Forms.TextBox t)
        {
            // Get the unformatted text
            string unformattedText = t.Text.Replace(".", "").Replace(",", "");

            if (decimal.TryParse(unformattedText, out decimal value))
            {
                // Use Vietnamese CultureInfo for 'N0' formatting to get dot separators
                System.Globalization.CultureInfo viVN = System.Globalization.CultureInfo.GetCultureInfo("vi-VN");
                t.Text = value.ToString("N0", viVN);
            }
            else
            {
                // Handle case where the default value is not a valid number
                t.Text = "0";
            }
        }
        private void rdoNormal_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoNormal.Checked == true)
            {
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalPassenger.Text = "0";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtminSotienVNDHT.Text = null;
                txtminSotienVNDHT.Text = null;
                txtSolanHT.Text = null;
                txtNgayHD.Text = null;
                txtSoHC.Text = null;
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
            }
        }
        private void rdoSimilar_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoSimilar.Checked == true)
            {
                txtRfdate.Text = string.Empty;
                txtRtdate.Text = string.Empty;

                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalPassenger.Text = "0";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";

                dtpRtdate.CustomFormat = null;
                dtpRtdate.CustomFormat = null;
                txtminSotienVNDHT.Text = null;
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                txtSolanHT.Text = null;
            }
        }

        private void rdoDupplicate_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoDupplicate.Checked == true)
            {
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalPassenger.Text = "0";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtSoHC.Text = null;
                txtminSotienVNDHT.Text = null;
                txtNgayHD.Text = null;
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                txtSolanHT.Text = null;
            }
        }

        private void txtMasoDN_TextChanged(object sender, EventArgs e)
        {
            if (txtMasoDN.Text.Length == 0)
            {
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtSoHC.Text = null;
                txtSoHD.Text = null;
                txtNgayHD.Text = null;
                txtTenDNBH.Text = null;
                txtSolanHT.Text = null;
            }
        }

        private void txtSoHC_TextChanged(object sender, EventArgs e)
        {
            if (txtSoHC.Text.Length == 0)
            {
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtTotalrows.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                dataGridView1.DataSource = null;
                txtSoHD.Text = null;
                txtNgayHD.Text = null;
                txtSolanHT.Text = null;
            }
        }

        private void txtSoHD_TextChanged(object sender, EventArgs e)
        {
            if (txtSoHD.Text.Length == 0)
            {
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtMasoDN.Text = null;
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
                txtSolanHT.Text = null;
            }
        }
        private void txtRfdate_TextChanged(object sender, EventArgs e)
        {
            if (txtRfdate.Text.Length == 0 && rdoRefundManyTimes.Checked == true)
            {

                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtTotalrows.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
            }
            else
            {
                //txtSolanHT.Text = "10";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtTotalrows.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
            }
        }

        private void txtRtdate_TextChanged(object sender, EventArgs e)
        {
            if (txtRtdate.Text.Length == 0 && rdoRefundManyTimes.Checked == true)
            {
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtTotalrows.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
            }
            else
            {
                //txtSolanHT.Text = "10";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtTotalrows.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
            }
        }

        private void txtTungay_TextChanged(object sender, EventArgs e)
        {
            if (rdoRefundManyTimes.Checked == true)
            {
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
            }
            else
            {
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
                txtSolanHT.Text = "10";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
            }
        }

        private void txtDenngay_TextChanged(object sender, EventArgs e)
        {
            if (rdoRefundManyTimes.Checked == true)
            {
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
            }
            else
            {
                txtSoHC.Text = null;
                txtNgayHD.Text = null;
                txtSolanHT.Text = "10";
                txtRefundAmt.Text = "0";
                txtGoodsValue.Text = "0";
                txtTotalRefundAmt.Text = "0";
                txtTotalValue.Text = "0";
                txtTotalBankServiceFee.Text = "0";
                txtTotalPassenger.Text = "0";
                txtMasoDN.Text = null;
                txtTenDNBH.Text = null;
            }
        }

        private void TextboxClear()
        {
            txtSoHC.Text = null;
            txtNgayHD.Text = null;
            txtSolanHT.Text = "10";
            txtRefundAmt.Text = "0";
            txtGoodsValue.Text = "0";
            txtTotalRefundAmt.Text = "0";
            txtTotalValue.Text = "0";
            txtTotalBankServiceFee.Text = "0";
            txtTotalPassenger.Text = "0";
            txtMasoDN.Text = null;
            txtTenDNBH.Text = null;
        }
        private decimal SafeDivideWithPercentage(decimal current, decimal previous)
        {
            if (previous == 0)
            {
                return current == 0 ? 100 : decimal.MaxValue;
            }
            return (current / previous) * 100;
        }
        // Models
        public class DateRange
        {
            public DateTime StartDate { get; set; }
            public DateTime EndDate { get; set; }
            public string StartDateString => StartDate.ToString("yyyy-MM-dd hh:mm:ss tt");
            public string EndDateString => EndDate.ToString("yyyy-MM-dd hh:mm:ss tt");
            public string FirstDayOfYear => new DateTime(StartDate.Year, 1, 1).ToString("yyyy-MM-dd hh:mm:ss tt");

            public DateRange PreviousYear => new DateRange
            {
                StartDate = StartDate.AddDays(-365),
                EndDate = EndDate.AddDays(-365)
            };
        }

        // Data Service
        public interface IReportService
        {
            System.Data.DataTable GetReportData(DateRange dateRange);
            ReportData CalculateReportData(System.Data.DataTable sourceData, DateRange dateRange);
        }

        public class ReportService : IReportService
        {
            private readonly Utility _utility;

            public ReportService()
            {
                _utility = new Utility();
            }

            //public System.Data.DataTable GetReportData(DateRange dateRange)
            //{
            //    using (var conn = _utility.OpenDB())
            //    {
            //        conn.Open();

            //        using (var cmd = new SqlCommand("ReportSearch_sp", conn))
            //        {
            //            cmd.CommandType = CommandType.StoredProcedure;
            //            cmd.Parameters.Add("@previousrfdate", SqlDbType.DateTime).Value = dateRange.PreviousYear.StartDateString;
            //            cmd.Parameters.Add("@rtdate", SqlDbType.DateTime).Value = dateRange.EndDateString;

            //            using (var da = new SqlDataAdapter(cmd))
            //            {
            //                var dt = new System.Data.DataTable();
            //                da.Fill(dt);
            //                return dt;
            //            }
            //        }
            //    }
            //}

            public System.Data.DataTable GetReportData(DateRange dateRange)
            {
                string sqlQuery = @"SELECT                    
                                                bc09.SoHD, 
                                                bc09.SoHC, 
                                                bc09.NgayHT, 
                                                bc09.TrigiaHHchuaVAT, 
                                                bc09.SotienVATDH, 
                                                bc09.SotienDVNHH 
                                            FROM BC09 AS bc09 
                                            WHERE bc09.NgayHT BETWEEN @StartDate AND @EndDate 
                                            ORDER BY bc09.NgayHT";

                using (var connection = _utility.OpenDB())
                {
                    connection.Open();

                    using (var command = new SqlCommand(sqlQuery, connection))
                    {
                        // Using SqlDbType for type safety
                        command.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = dateRange.PreviousYear.StartDate;
                        command.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = dateRange.EndDate;

                        using (var dataAdapter = new SqlDataAdapter(command))
                        {
                            var dataTable = new System.Data.DataTable();
                            dataAdapter.Fill(dataTable);
                            return dataTable;
                        }
                    }
                }
            }
            public ReportData CalculateReportData(System.Data.DataTable sourceData, DateRange dateRange)
            {
                if (sourceData.Rows.Count == 0)
                    return null;

                var reportData = new ReportData();

                // Current year calculations
                reportData.TotalGoodsValue = CalculateSum(sourceData, "TrigiaHHchuaVAT",
                    dateRange.StartDateString, dateRange.EndDateString);
                reportData.AccumulatedGoodsValue = CalculateSum(sourceData, "TrigiaHHchuaVAT",
                    dateRange.FirstDayOfYear, dateRange.EndDateString);

                reportData.TotalVATRefund = CalculateSum(sourceData, "SotienVATDH",
                    dateRange.StartDateString, dateRange.EndDateString);
                reportData.AccumulatedVATRefund = CalculateSum(sourceData, "SotienVATDH",
                    dateRange.FirstDayOfYear, dateRange.EndDateString);

                reportData.TotalServiceFee = CalculateSum(sourceData, "SotienDVNHH",
                    dateRange.StartDateString, dateRange.EndDateString);
                reportData.AccumulatedServiceFee = CalculateSum(sourceData, "SotienDVNHH",
                    dateRange.FirstDayOfYear, dateRange.EndDateString);

                reportData.PassengerTurns = CountDistinctPassengers(sourceData,
                    dateRange.StartDateString, dateRange.EndDateString);
                reportData.AccumulatedPassengerTurns = CountDistinctPassengers(sourceData,
                    dateRange.FirstDayOfYear, dateRange.EndDateString);

                // Previous year calculations
                reportData.PreviousYearGoodsValue = CalculateSum(sourceData, "TrigiaHHchuaVAT",
                    dateRange.PreviousYear.StartDateString, dateRange.PreviousYear.EndDateString);
                reportData.PreviousYearVATRefund = CalculateSum(sourceData, "SotienVATDH",
                    dateRange.PreviousYear.StartDateString, dateRange.PreviousYear.EndDateString);
                reportData.PreviousYearPassengerTurns = CountDistinctPassengers(sourceData,
                    dateRange.PreviousYear.StartDateString, dateRange.PreviousYear.EndDateString);

                return reportData;
            }

            private decimal CalculateSum(System.Data.DataTable data, string columnName, string startDate, string endDate)
            {
                return data.AsEnumerable()
                    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(startDate) &&
                               Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(endDate))
                    .Sum(x => x.Field<decimal>(columnName));
            }

            private decimal CountDistinctPassengers(System.Data.DataTable data, string startDate, string endDate)
            {
                //return data.AsEnumerable()
                //    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(startDate) &&
                //               Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(endDate))
                //    .Select(r => r.Field<string>("SoHC"))
                //    .Distinct()
                //    .Count();

                return data.AsEnumerable()
                    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(startDate) &&
                                Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(endDate))
                    .GroupBy(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")).Date)  // Group by date (ignoring time)
                    .Select(group => new
                    {
                        Date = group.Key,
                        DistinctCount = group.Select(r => r.Field<string>("SoHC")).Distinct().Count()
                    })
                    .Sum(x => x.DistinctCount);
            }
        }

        private async void btnReport_Click(object sender, EventArgs e)
        {
            try
            {
                RestoreDataGridView();

                if (!ValidateInput())
                    return;

                await GenerateReportAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private bool ValidateInput()
        {
            if (rdoNormal.Checked == false)
                return true;

            if (string.IsNullOrWhiteSpace(txtRfdate.Text) || string.IsNullOrWhiteSpace(txtRtdate.Text))
            {
                MessageBox.Show("Please enter both start and end dates.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!DateTime.TryParse(txtRfdate.Text, out DateTime startDate) ||
                !DateTime.TryParse(txtRtdate.Text, out DateTime endDate))
            {
                MessageBox.Show("Please enter valid dates in the correct format.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (startDate > endDate)
            {
                MessageBox.Show("Start date cannot be after end date.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private async Task GenerateReportAsync()
        {
            rdoNormal.Checked = true;
            isDataGridViewInCustomLayout = false;

            var dateRange = ParseDateRange();
            var reportService = new ReportService();

            // Get data
            var sourceData = await Task.Run(() => reportService.GetReportData(dateRange));

            if (sourceData.Rows.Count == 0)
            {
                MessageBox.Show("No datarows to show.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                dataGridView1.DataSource = null;
                return;
            }

            // Calculate report data
            var reportData = reportService.CalculateReportData(sourceData, dateRange);

            // Create and display result table
            var resultTable = CreateResultTable(reportData);
            dataGridView1.DataSource = resultTable;

            // Configure UI
            ConfigureGridView();
            SetupFixedDatagridview();

            LoadChartDataAsAreaChart();

            //LoadChartData();
            //LoadChartDataAsStackedColumn();
            //LoadChartDataAsSplineChart();
            //LoadChartDataAsComboChart();

            txtTotalrows.Text = resultTable.Rows.Count.ToString("#,##0");
            isDataGridViewInCustomLayout = true;

            this.Refresh();
        }

        private DateRange ParseDateRange()
        {
            var startDate = DateTime.ParseExact(txtRfdate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var endDate = DateTime.ParseExact(txtRtdate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

            return new DateRange
            {
                StartDate = startDate,
                EndDate = endDate
            };
        }

        private System.Data.DataTable CreateResultTable(ReportData data)
        {
            var resultTable = new System.Data.DataTable();

            // Define columns
            resultTable.Columns.AddRange(new[]
            {
                new DataColumn("STT", typeof(string)),
                new DataColumn("Motahanghoa", typeof(string)),
                new DataColumn("TongTrigiaHHchuaVAT (VND)", typeof(decimal)),
                new DataColumn("TongLKTrigiaHHchuaVAT (VND)", typeof(decimal)),
                new DataColumn("TongSotienVATDH (VND)", typeof(decimal)),
                new DataColumn("TongLKSotienVATDH (VND)", typeof(decimal)),
                new DataColumn("TongSotienDVNHH (VND)", typeof(decimal)),
                new DataColumn("TongLKSotienDVNHH (VND)", typeof(decimal)),
                new DataColumn("TongLuotHK (HK)", typeof(decimal)),
                new DataColumn("TongLKLuotHK (HK)", typeof(decimal)),
                new DataColumn("SosanhTongTrigiaHHchuaVAT (%)", typeof(string)),
                new DataColumn("SosanhTongSotienVATDH (%)", typeof(string)),
                new DataColumn("SosanhTongLuotHK (%)", typeof(string))
            });

            // Add data row
            var row = resultTable.NewRow();
            row["STT"] = "1";
            row["Motahanghoa"] = "Giày, dép, quần, áo, valy, túi xách, mỹ phẩm, đồng hồ, tư trang, vật dụng cá nhân";
            row["TongTrigiaHHchuaVAT (VND)"] = data.TotalGoodsValue;
            row["TongLKTrigiaHHchuaVAT (VND)"] = data.AccumulatedGoodsValue;
            row["TongSotienVATDH (VND)"] = data.TotalVATRefund;
            row["TongLKSotienVATDH (VND)"] = data.AccumulatedVATRefund;
            row["TongSotienDVNHH (VND)"] = data.TotalServiceFee;
            row["TongLKSotienDVNHH (VND)"] = data.AccumulatedServiceFee;
            row["TongLuotHK (HK)"] = data.PassengerTurns;
            row["TongLKLuotHK (HK)"] = data.AccumulatedPassengerTurns;
            row["SosanhTongTrigiaHHchuaVAT (%)"] = data.GoodsValueComparison.ToString("#,##0.00");
            row["SosanhTongSotienVATDH (%)"] = data.VATRefundComparison.ToString("#,##0.00");
            row["SosanhTongLuotHK (%)"] = data.PassengerTurnsComparison.ToString("#,##0.00");

            resultTable.Rows.Add(row);
            return resultTable;
        }

        private void ConfigureGridView()
        {
            DataGridViewHelper.ConfigureReportGridView(dataGridView1);
        }

        //private void btnReport_Click(object sender, EventArgs e)
        //{       
        //    Utility ut = new Utility();
        //    var conn = ut.OpenDB();
        //    string rfDate = txtRfdate.Text.Trim();
        //    string rtDate = txtRtdate.Text.Trim();
        //    rdoNormal.Checked = true;
        //    isDataGridViewInCustomLayout = false;

        //    try
        //    {
        //        if (rdoNormal.Checked == false)
        //            return;

        //        if (conn.State == ConnectionState.Closed)
        //        {
        //            conn.Open();

        //            if (string.IsNullOrWhiteSpace(rfDate) || string.IsNullOrWhiteSpace(rtDate))
        //            {
        //                MessageBox.Show("Please enter both start and end dates.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                return;
        //            }

        //            if (!DateTime.TryParse(txtRfdate.Text, out DateTime startDate) || !DateTime.TryParse(txtRtdate.Text, out DateTime endDate))
        //            {
        //                MessageBox.Show("Please enter valid dates in the correct format.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                return;
        //            }

        //            if (startDate > endDate)
        //            {
        //                MessageBox.Show("Start date cannot be after end date.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //                return;
        //            }

        //            DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //            var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");
        //            DateTime ppreviousrfdate = prfdate.AddDays(-365);
        //            var previousrfdate = ppreviousrfdate.ToString("yyyy-MM-dd hh:mm:ss tt");

        //            int year = prfdate.Year;
        //            DateTime firstdoy = new DateTime(year, 1, 1);
        //            var strfirstdoy = firstdoy.ToString("yyyy-MM-dd hh:mm:ss tt");

        //            DateTime prtdate = DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
        //            var rtdate = prtdate.ToString("yyyy-MM-dd hh:mm:ss tt");
        //            DateTime ppreviousrtdate = prtdate.AddDays(-365);
        //            var previousrtdate = ppreviousrtdate.ToString("yyyy-MM-dd hh:mm:ss tt");

        //            // Old codes: stored procedure approach
        //            #region
        //            //using (SqlCommand cmd = new SqlCommand("ReportSearch_sp", conn))
        //            //{
        //            //    cmd.CommandType = CommandType.StoredProcedure;

        //            //    cmd.Parameters.Add("@previousrfdate", SqlDbType.DateTime).Value = previousrfdate;
        //            //    cmd.Parameters.Add("@rtdate", SqlDbType.DateTime).Value = rtdate;

        //            //    cmd.ExecuteNonQuery();

        //            //    da = new SqlDataAdapter(cmd);
        //            //    dt = new System.Data.DataTable();
        //            //    da.Fill(dt);
        //            //}
        //            #endregion

        //            string sqlQuery = @"
        //                            SELECT 
        //                                bc09.SoHD, 
        //                                bc09.SoHC, 
        //                                bc09.NgayHT, 
        //                                bc09.TrigiaHHchuaVAT, 
        //                                bc09.SotienVATDH, 
        //                                bc09.SotienDVNHH 
        //                            FROM BC09 AS bc09 
        //                            WHERE bc09.NgayHT BETWEEN @StartDate AND @EndDate 
        //                            ORDER BY bc09.NgayHT";

        //            SqlCommand cmd = new SqlCommand(sqlQuery, conn);
        //            cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = previousrfdate;
        //            cmd.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = rtdate;
        //            SqlDataAdapter da = new SqlDataAdapter(cmd);
        //            System.Data.DataTable dt = new System.Data.DataTable();
        //            da.Fill(dt);                   

        //            if (dt.Rows.Count > 0)
        //            {
        //                System.Data.DataTable resultTable = new System.Data.DataTable();

        //                resultTable.Columns.Add(new DataColumn("STT", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("Motahanghoa", typeof(string)));
        //                resultTable.Columns.Add(new DataColumn("TongTrigiaHHchuaVAT (VND)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("TongLKTrigiaHHchuaVAT (VND)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("TongSotienVATDH (VND)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("TongLKSotienVATDH (VND)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("TongSotienDVNHH (VND)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("TongLKSotienDVNHH (VND)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("TongLuotHK (HK)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("TongLKLuotHK (HK)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("SosanhTongTrigiaHHchuaVAT (%)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("SosanhTongSotienVATDH (%)", typeof(decimal)));
        //                resultTable.Columns.Add(new DataColumn("SosanhTongLuotHK (%)", typeof(decimal)));

        //                //Data of the current searching year
        //                var tongTGHHchuaVAT = dt.AsEnumerable()
        //                 .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(rfdate) &&
        //                 Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                 .Sum(x => x.Field<decimal>("TrigiaHHchuaVAT"));
        //                var tongLKTGHHchuaVAT = dt.AsEnumerable()
        //                    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(strfirstdoy) &&
        //                    Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                    .Sum(x => x.Field<decimal>("TrigiaHHchuaVAT"));
        //                var tongSotienVATDH = dt.AsEnumerable()
        //                .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(rfdate) &&
        //                Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                .Sum(x => x.Field<decimal>("SotienVATDH"));
        //                var tongLKSotienVATDH = dt.AsEnumerable()
        //                    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(strfirstdoy) &&
        //                    Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                    .Sum(x => x.Field<decimal>("SotienVATDH"));
        //                var tongSotienDVNHH = dt.AsEnumerable()
        //                .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(rfdate) &&
        //                Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                .Sum(x => x.Field<decimal>("SotienDVNHH"));
        //                var tongLKSotienDVNHH = dt.AsEnumerable()
        //                    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(strfirstdoy) &&
        //                    Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                    .Sum(x => x.Field<decimal>("SotienDVNHH"));

        //                //decimal tongLuotHK = dt.AsEnumerable()
        //                //   .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(rfdate) &&
        //                //   Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                //   .Select(r => r.Field<string>("SoHC"))
        //                //   .Distinct()
        //                //   .Count();
        //                //decimal tongLKLuotHK = dt.AsEnumerable()
        //                //    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(strfirstdoy) &&
        //                //    Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                //    .Select(r => r.Field<string>("SoHC"))
        //                //    .Distinct()
        //                //    .Count();

        //                decimal tongLuotHK = dt.AsEnumerable()
        //                                    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(rfdate) &&
        //                                                Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                                    .GroupBy(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")).Date)  // Group by date (ignoring time)
        //                                    .Select(group => new {
        //                                        Date = group.Key,
        //                                        DistinctCount = group.Select(r => r.Field<string>("SoHC")).Distinct().Count()
        //                                    })
        //                                    .Sum(x => x.DistinctCount);

        //                decimal tongLKLuotHK = dt.AsEnumerable()
        //                                    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(strfirstdoy) &&
        //                                                Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(rtdate))
        //                                    .GroupBy(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")).Date)  // Group by date (ignoring time)
        //                                    .Select(group => new {
        //                                        Date = group.Key,
        //                                        DistinctCount = group.Select(r => r.Field<string>("SoHC")).Distinct().Count()
        //                                    })
        //                                    .Sum(x => x.DistinctCount);

        //                //Data of previous year
        //                var previoustongTGHHchuaVAT = dt.AsEnumerable()
        //                 .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(previousrfdate) &&
        //                 Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(previousrtdate))
        //                 .Sum(x => x.Field<decimal>("TrigiaHHchuaVAT"));
        //                decimal comparetongTGHHchuaVAT = SafeDivideWithPercentage(tongTGHHchuaVAT, previoustongTGHHchuaVAT);
        //                var previoustongSotienVATDH = dt.AsEnumerable()
        //                .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(previousrfdate) &&
        //                 Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(previousrtdate))
        //                 .Sum(x => x.Field<decimal>("SotienVATDH"));
        //                decimal comparetongSotienVATDH = SafeDivideWithPercentage(tongSotienVATDH, previoustongSotienVATDH);

        //                //decimal previoustongLuotHK = dt.AsEnumerable()
        //                //    .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(previousrfdate) &&
        //                //    Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(previousrtdate))
        //                //    .Select(r => r.Field<string>("SoHC"))
        //                //    .Distinct()
        //                //    .Count();

        //                decimal previoustongLuotHK = dt.AsEnumerable()
        //                                     .Where(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")) >= Convert.ToDateTime(previousrfdate) &&
        //                                                 Convert.ToDateTime(x.Field<DateTime>("NgayHT")) <= Convert.ToDateTime(previousrtdate))
        //                                     .GroupBy(x => Convert.ToDateTime(x.Field<DateTime>("NgayHT")).Date)  // Group by date (ignoring time)
        //                                     .Select(group => new {
        //                                         Date = group.Key,
        //                                         DistinctCount = group.Select(r => r.Field<string>("SoHC")).Distinct().Count()
        //                                     })
        //                                     .Sum(x => x.DistinctCount);

        //                decimal comparetongLuotHK = SafeDivideWithPercentage(tongLuotHK, previoustongLuotHK);

        //                DataRow dr = resultTable.NewRow();

        //                dr["STT"] = 1.ToString(("#,##0"));
        //                dr["Motahanghoa"] = "Giày, dép, quần, áo, valy, túi xách, mỹ phẩm, đồng hồ, tư trang, vật dụng cá nhân";

        //                dr["TongTrigiaHHchuaVAT (VND)"] = tongTGHHchuaVAT;
        //                dr["TongLKTrigiaHHchuaVAT (VND)"] = tongLKTGHHchuaVAT;

        //                dr["TongSotienVATDH (VND)"] = tongSotienVATDH;
        //                dr["TongLKSotienVATDH (VND)"] = tongLKSotienVATDH;

        //                dr["TongSotienDVNHH (VND)"] = tongSotienDVNHH;
        //                dr["TongLKSotienDVNHH (VND)"] = tongLKSotienDVNHH;

        //                dr["TongLuotHK (HK)"] = tongLuotHK;
        //                dr["TongLKLuotHK (HK)"] = tongLKLuotHK;

        //                dr["SosanhTongTrigiaHHchuaVAT (%)"] = comparetongTGHHchuaVAT.ToString("#,##0.00");
        //                dr["SosanhTongSotienVATDH (%)"] = comparetongSotienVATDH.ToString("#,##0.00");
        //                dr["SosanhTongLuotHK (%)"] = comparetongLuotHK.ToString("#,##0.00");

        //                resultTable.Rows.Add(dr);
        //                dataGridView1.DataSource = resultTable;

        //                DataGridViewCellStyle style = new DataGridViewCellStyle();
        //                style.Format = "N0";

        //                //this.dgvNhaplieuExcel.Columns["STT"].DefaultCellStyle = style;
        //                //this.dataGridView1.Columns["Motahanghoa"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongTrigiaHHchuaVAT (VND)"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongLKTrigiaHHchuaVAT (VND)"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongSotienVATDH (VND)"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongLKSotienVATDH (VND)"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongSotienDVNHH (VND)"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongLKSotienDVNHH (VND)"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongLuotHK (HK)"].DefaultCellStyle = style;
        //                this.dataGridView1.Columns["TongLKLuotHK (HK)"].DefaultCellStyle = style;

        //                this.dataGridView1.Columns["STT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //                this.dataGridView1.Columns["Motahanghoa"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //                this.dataGridView1.Columns["TongTrigiaHHchuaVAT (VND)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["TongLKTrigiaHHchuaVAT (VND)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["TongSotienVATDH (VND)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["TongLKSotienVATDH (VND)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["TongSotienDVNHH (VND)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["TongLKSotienDVNHH (VND)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["TongLuotHK (HK)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["TongLKLuotHK (HK)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //                this.dataGridView1.Columns["SosanhTongTrigiaHHchuaVAT (%)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //                this.dataGridView1.Columns["SosanhTongSotienVATDH (%)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //                this.dataGridView1.Columns["SosanhTongLuotHK (%)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

        //                // Old code
        //                #region
        //                //// Define the display names (Header Text) for the columns
        //                //this.dataGridView1.Columns["STT"].HeaderText = "No.";
        //                //this.dataGridView1.Columns["Motahanghoa"].HeaderText = "Item Description"; // Example translation
        //                //this.dataGridView1.Columns["TongTrigiaHHchuaVAT (VND)"].HeaderText = "Total Goods Value (ex VAT)"; // Example translation
        //                //this.dataGridView1.Columns["TongLKTrigiaHHchuaVAT (VND)"].HeaderText = "Total Accumulated Goods Value (ex VAT)";

        //                //this.dataGridView1.Columns["TongSotienVATDH (VND)"].HeaderText = "Total VAT Refundable Amount";
        //                //this.dataGridView1.Columns["TongLKSotienVATDH (VND)"].HeaderText = "Total Accumulated VAT Refund";

        //                //this.dataGridView1.Columns["TongSotienDVNHH (VND)"].HeaderText = "Total Service Fee";
        //                //this.dataGridView1.Columns["TongLKSotienDVNHH (VND)"].HeaderText = "Total Accumulated Service Fee";

        //                //this.dataGridView1.Columns["TongLuotHK (HK)"].HeaderText = "Total Passenger Turns";
        //                //this.dataGridView1.Columns["TongLKLuotHK (HK)"].HeaderText = "Total Accumulated Passenger Turns";

        //                //this.dataGridView1.Columns["SosanhTongTrigiaHHchuaVAT (%)"].HeaderText = "Goods Value Comparison (%)";
        //                //this.dataGridView1.Columns["SosanhTongSotienVATDH (%)"].HeaderText = "VAT Refund Comparison (%)";
        //                //this.dataGridView1.Columns["SosanhTongLuotHK (%)"].HeaderText = "Passenger Turn Comparison (%)";
        //                #endregion

        //                // Define all custom mappings here. Use the actual DataTable column name as the Key.
        //                var columnHeaderMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        //                    {
        //                        { "STT", "No." },
        //                        { "Motahanghoa", "Description Of Gooods" },
        //                        { "TongTrigiaHHchuaVAT (VND)", "Total Goods Value (ex VAT)" },
        //                        { "TongLKTrigiaHHchuaVAT (VND)", "Total Accumulated Goods Value (ex VAT)" },

        //                        { "TongSotienVATDH (VND)", "Total VAT Refundable Amount" },
        //                        { "TongLKSotienVATDH (VND)", "Total Accumulated VAT Refund" },

        //                        { "TongSotienDVNHH (VND)", "Total Service Fee" },
        //                        { "TongLKSotienDVNHH (VND)", "Total Accumulated Service Fee" },

        //                        { "TongLuotHK (HK)", "Total Passenger Turns" },
        //                        { "TongLKLuotHK (HK)", "Total Accumulated Passenger Turns" },

        //                        { "SosanhTongTrigiaHHchuaVAT (%)", "Goods Value Comparison (%)" },
        //                        { "SosanhTongSotienVATDH (%)", "VAT Refund Comparison (%)" },
        //                        { "SosanhTongLuotHK (%)", "Passenger Turn Comparison (%)" }
        //                    };

        //                foreach (DataGridViewColumn column in dataGridView1.Columns)
        //                {
        //                    if (columnHeaderMap.TryGetValue(column.Name, out var displayHeader))
        //                    {
        //                        column.HeaderText = displayHeader;
        //                    }
        //                }

        //                // Count total datarows         
        //                txtTotalrows.Text = Convert.ToString(resultTable.Rows.Count.ToString("#,##0"));

        //                SetupFixedDatagridview();
        //                //ToggleDataGridViewLayout();
        //                LoadChartData();

        //                // Set font for headers           
        //                var headerFont = new System.Drawing.Font("Microsoft Sans Serif", 11f, FontStyle.Bold);
        //                dataGridView1.EnableHeadersVisualStyles = false; // Disable system styles
        //                dataGridView1.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        //                {
        //                    Font = headerFont,
        //                    ForeColor = Color.DeepSkyBlue, // FromArgb(51, 122, 183),
        //                    BackColor = Color.AntiqueWhite, // FromArgb(51, 122, 183), // Bootstrap primary blue
        //                    Alignment = DataGridViewContentAlignment.MiddleCenter,
        //                    Padding = new Padding(3)
        //                };

        //                // Optional: Add gradient background
        //                #region
        //                //dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
        //                //dataGridView1.ColumnHeadersHeight = 30;
        //                //dataGridView1.Paint += (sender, e) =>
        //                //{
        //                //    var headerBounds = new System.Drawing.Rectangle(
        //                //        0, 0,
        //                //        dataGridView1.Width,
        //                //        dataGridView1.ColumnHeadersHeight);

        //                //    using (var brush = new LinearGradientBrush(
        //                //        headerBounds,
        //                //        Color.WhiteSmoke, //FromArgb(70, 130, 180),
        //                //        Color.AntiqueWhite, //FromArgb(40, 90, 150),
        //                //        30f))
        //                //    {
        //                //        e.Graphics.FillRectangle(brush, headerBounds);
        //                //    }
        //                //};
        //                #endregion

        //                // Set font for the entire grid
        //                dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11f, FontStyle.Regular);

        //                // Set up for the entire grid
        //                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //                dataGridView1.DefaultCellStyle.BackColor = Color.FloralWhite;
        //                dataGridView1.DefaultCellStyle.ForeColor = Color.DarkBlue;
        //                dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        //                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

        //                dataGridView1.GridColor = Color.Silver;
        //                dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

        //                dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
        //                dataGridView1.MultiSelect = true;

        //                this.Refresh();
        //                isDataGridViewInCustomLayout = true;
        //            }
        //            else
        //            {
        //                MessageBox.Show("No datarows to show.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                dataGridView1.DataSource = null;
        //            }
        //        }
        //        //Close connection
        //        conn.Close();
        //        conn.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //}

        private void SetupFixedDatagridview()
        {
            try
            {
                // 1. Capture the "Original" state before doing anything
                if (!isDataGridViewInCustomLayout)
                {
                    originalHeight = dataGridView1.Height;
                    originalWidth = dataGridView1.Width;
                    originalLocation = dataGridView1.Location;
                    originalDock = dataGridView1.Dock;
                    originalAnchor = dataGridView1.Anchor;
                }

                CleanupPreviousLayout();
                CreateGridContainer();
                isDataGridViewInCustomLayout = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting up data grid view: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                RestoreDefaultLayout();
            }
        }

        private void CleanupPreviousLayout()
        {
            if (gridContainer != null && this.Controls.Contains(gridContainer))
            {
                this.Controls.Remove(gridContainer);
                gridContainer.Dispose();
                gridContainer = null;
            }
        }

        private void CreateGridContainer()
        {
            gridContainer = new Panel { Dock = DockStyle.Fill };
            this.Controls.Add(gridContainer);

            CreateHeaderPanel();
            CreateBodyPanel();
        }

        private void CreateHeaderPanel()
        {
            var headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = dataGridView1.ColumnHeadersHeight
            };
            gridContainer.Controls.Add(headerPanel);
        }

        private void CreateBodyPanel()
        {
            var bodyPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };
            gridContainer.Controls.Add(bodyPanel);

            ConfigureDataGridViewInPanel(bodyPanel);
            SetupScrollSync(bodyPanel);
        }

        private void ConfigureDataGridViewInPanel(Panel bodyPanel)
        {
            // Remove from current parent
            if (dataGridView1.Parent != null)
            {
                dataGridView1.Parent.Controls.Remove(dataGridView1);
            }

            // Configure dimensions and position
            dataGridView1.Top = 288;
            dataGridView1.Left = 28;
            dataGridView1.Width = Math.Max(1112, bodyPanel.ClientSize.Width - 56);
            dataGridView1.Height = Math.Max(100, bodyPanel.ClientSize.Height - 1000);
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;

            // Add to panel
            bodyPanel.Controls.Add(dataGridView1);
        }

        private void SetupScrollSync(Panel bodyPanel)
        {
            var headerPanel = gridContainer.Controls[0] as Panel;
            if (headerPanel == null) return;

            bodyPanel.Scroll += (s, e) =>
            {
                if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                {
                    headerPanel.AutoScrollPosition = new System.Drawing.Point(e.NewValue, 0);
                }
            };
        }

        private void RestoreDefaultLayout()
        {
            // Restore dataGridView to form if something went wrong
            if (dataGridView1.Parent != gridContainer)
            {
                var currentParent = dataGridView1.Parent;
                currentParent?.Controls.Remove(dataGridView1);
                this.Controls.Add(dataGridView1);
                dataGridView1.Dock = DockStyle.Fill;
            }
        }

        public void ExportReport(System.Data.DataTable dataTable, string sheetName, string title)
        {
            Microsoft.Office.Interop.Excel.Application oExcel = null;
            Microsoft.Office.Interop.Excel.Workbook oBook = null;
            Microsoft.Office.Interop.Excel.Worksheet oSheet = null;

            try
            {
                // Parse dates once at the beginning
                DateTime now = DateTime.Now;
                // Helper: try parse textbox, else use corresponding DateTimePicker
                bool TryGetDate(System.Windows.Forms.TextBox tb, DateTimePicker dp, out DateTime dt)
                {
                    var s = tb.Text?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(s)
                        && DateTime.TryParseExact(s, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                        return true;

                    dt = dp?.Value ?? DateTime.Now;
                    return false; // indicates textbox was invalid/empty (fallback used)
                }

                DateTime prfdate, prtdate;
                bool rfOk = TryGetDate(txtRfdate, dtpRfdate, out prfdate);
                bool rtOk = TryGetDate(txtRtdate, dtpRtdate, out prtdate);

                // If you require the user to explicitly input dates, block and notify:
                if (!rfOk || !rtOk)
                {
                    MessageBox.Show("Please enter valid From/To dates in format dd/MM/yyyy or pick them from the calendar to export file.", "Invalid date", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //DateTime prfdate = DateTime.ParseExact(txtRfdate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                string rfdate = prfdate.ToString("dd-MM-yyyy");
                DateTime ptfdate = DateTime.ParseExact(txtRtdate.Text.Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                string rtdate = ptfdate.ToString("dd-MM-yyyy");

                // Initialize Excel application
                oExcel = new Microsoft.Office.Interop.Excel.Application();
                oExcel.Visible = true;
                oExcel.DisplayAlerts = false;
                oExcel.Application.SheetsInNewWorkbook = 1;

                // Create workbook and worksheet
                oBook = (Microsoft.Office.Interop.Excel.Workbook)(oExcel.Workbooks.Add(Type.Missing));
                oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oBook.Worksheets[1];
                oSheet.Name = sheetName;

                // Setup page formatting
                SetupPageFormatting(oSheet);

                // Create headers and titles
                CreateHeaders(oSheet, title, now, rfdate, rtdate);

                // Create column headers
                CreateColumnHeaders(oSheet);

                // Create footer signatures
                CreateFooterSignatures(oSheet);

                // Format header row
                FormatHeaderRow(oSheet);

                // Export data from DataTable
                ExportDataTableToExcel(dataTable, oSheet);

                // Save the file
                SaveExcelFile(oBook, oExcel, rfdate, rtdate, now);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                // Proper COM cleanup
                CleanupExcelObjects(oSheet, oBook, oExcel);
            }
        }

        // Helper methods
        private void SetupPageFormatting(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            var pageSetup = worksheet.PageSetup;
            pageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            pageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
            pageSetup.TopMargin = 20;
            pageSetup.BottomMargin = 20;
            pageSetup.LeftMargin = 30;
            pageSetup.RightMargin = 20;
        }

        private void CreateHeaders(Microsoft.Office.Interop.Excel.Worksheet worksheet, string title, DateTime now, string rfdate, string rtdate)
        {
            // Header row 1
            var header1 = worksheet.Range["A1", "C1"];
            SetCellFormat(header1, "CHI CỤC HẢI QUAN KHU VỰC II", true, false, 13,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            header1.MergeCells = true;

            // Appendix
            var appendix = worksheet.Range["H1", "H1"];
            SetCellFormat(appendix, "Phụ lục I", true, true, 14,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            appendix.MergeCells = true;

            // Header row 2
            var header2 = worksheet.Range["A2", "C2"];
            header2.Value2 = "HẢI QUAN CỬA KHẨU\r\nSÂN BAY QUỐC TẾ TÂN SƠN NHẤT";
            SetCellFormat(header2, null, true, true, 13,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            header2.MergeCells = true;
            header2.WrapText = true;
            header2.RowHeight = 29;

            // Draw line shape
            Microsoft.Office.Interop.Excel.Shape lineShape = worksheet.Shapes.AddLine(99, 55, 171, 55);
            lineShape.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
            lineShape.Line.Weight = 1;

            // Report title
            var reportTitle = worksheet.Range["A5", "J5"];
            SetCellFormat(reportTitle, title, true, true, 14,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            reportTitle.MergeCells = true;

            // Report subtitle
            var reportSubtitle = worksheet.Range["A6", "J6"];
            string subtitle = $"(Đính kèm công văn số:               /BC-CKTSN ngày         /{now.Month:00}/{now.Year:0000})";
            SetCellFormat(reportSubtitle, subtitle, false, false, 14,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            reportSubtitle.MergeCells = true;
        }

        private void CreateColumnHeaders(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            // Define column configurations
            var columnConfigs = new[]
            {
                new { Range = "A8:A9", Text = "STT", Width = 4, RowHeight = 0, Merge = true },
                new { Range = "B8:B9", Text = "MẶT HÀNG", Width = 20, RowHeight = 0, Merge = true },
                new { Range = "C8:D8", Text = "TRỊ GIÁ HÀNG HOÀN THUẾ (ĐỒNG)", Width = 30, RowHeight = 60, Merge = true },
                new { Range = "C9:C9", Text = "Trong tháng", Width = 15, RowHeight = 0, Merge = false },
                new { Range = "D9:D9", Text = "Lũy kế đến kỳ báo cáo", Width = 15, RowHeight = 0, Merge = false },
                new { Range = "E8:F8", Text = "SỐ TIỀN THUẾ GTGT ĐƯỢC HOÀN (ĐỒNG)", Width = 30, RowHeight = 60, Merge = true },
                new { Range = "E9:E9", Text = "Trong tháng", Width = 15, RowHeight = 0, Merge = false },
                new { Range = "F9:F9", Text = "Lũy kế đến kỳ báo cáo", Width = 15, RowHeight = 0, Merge = false },
                new { Range = "G8:H8", Text = "SỐ TIỀN DỊCH VỤ NGÂN HÀNG ĐƯỢC HƯỞNG (ĐỒNG)", Width = 24, RowHeight = 60, Merge = true },
                new { Range = "G9:G9", Text = "Trong tháng", Width = 12, RowHeight = 0, Merge = false },
                new { Range = "H9:H9", Text = "Lũy kế đến kỳ báo cáo", Width = 12, RowHeight = 0, Merge = false },
                new { Range = "I8:J8", Text = "SỐ LƯỢNG NGƯỜI NƯỚC NGOÀI ĐÃ ĐƯỢC HOÀN THUẾ (NGƯỜI)", Width = 16, RowHeight = 100, Merge = true },
                new { Range = "I9:I9", Text = "Trong tháng", Width = 8, RowHeight = 52, Merge = false },
                new { Range = "J9:J9", Text = "Lũy kế đến kỳ báo cáo", Width = 8, RowHeight = 52, Merge = false }
            };

            foreach (var config in columnConfigs)
            {
                var range = worksheet.Range[config.Range.Split(':')[0], config.Range.Split(':')[1]];
                range.Value2 = config.Text;
                range.Font.Name = "Times New Roman";
                range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                if (config.Merge)
                    range.MergeCells = true;

                if (config.Width > 0)
                    range.ColumnWidth = config.Width;

                if (config.RowHeight > 0)
                    range.RowHeight = config.RowHeight;

                range.WrapText = true;
            }
        }

        private void CreateFooterSignatures(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            // Maker
            var maker = worksheet.Range["B13", "C13"];
            SetCellFormat(maker, "NGƯỜI LẬP", true, true, 13,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            maker.MergeCells = true;

            var makerName = worksheet.Range["B18", "C18"];
            string userName = ((Form)this.MdiParent).Controls["lblUserName"].Text;
            SetCellFormat(makerName, userName, true, true, 14,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            makerName.MergeCells = true;

            // Signatories
            var headSignatory1 = worksheet.Range["G12", "H12"];
            SetCellFormat(headSignatory1, "KT. ĐỘI TRƯỞNG", true, true, 13,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            headSignatory1.MergeCells = true;

            var headSignatory2 = worksheet.Range["G13", "H13"];
            SetCellFormat(headSignatory2, "PHÓ ĐỘI TRƯỞNG", true, true, 13,
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter);
            headSignatory2.MergeCells = true;
        }

        private void FormatHeaderRow(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            var rowHead = worksheet.Range["A8", "J9"];
            rowHead.Font.Bold = true;
            rowHead.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;
            rowHead.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        private void ExportDataTableToExcel(System.Data.DataTable dataTable, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            if (dataTable.Rows.Count == 0) return;

            // Convert DataTable to array (more efficient than cell-by-cell assignment)
            object[,] dataArray = ConvertDataTableToArray(dataTable);

            // Calculate data range
            int rowStart = 10;
            int columnStart = 1;
            int rowEnd = rowStart + dataTable.Rows.Count - 1;
            int columnEnd = dataTable.Columns.Count;

            // Get data range and assign values in one operation
            var dataRange = worksheet.Range[
                worksheet.Cells[rowStart, columnStart],
                worksheet.Cells[rowEnd, columnEnd]
            ];
            dataRange.Value2 = dataArray;

            // Apply formatting to the entire data range at once
            FormatDataRange(dataRange);

            // Apply specific column formatting
            ApplyColumnSpecificFormatting(worksheet, rowStart, rowEnd, columnStart, columnEnd);
        }

        private object[,] ConvertDataTableToArray(System.Data.DataTable dataTable)
        {
            object[,] array = new object[dataTable.Rows.Count, dataTable.Columns.Count];

            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                DataRow dataRow = dataTable.Rows[row];
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    array[row, col] = dataRow[col];
                }
            }

            return array;
        }

        private void FormatDataRange(Microsoft.Office.Interop.Excel.Range dataRange)
        {
            dataRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;
            dataRange.WrapText = true;
            dataRange.Font.Name = "Times New Roman";
            dataRange.Font.Size = 11;
            dataRange.NumberFormat = "#,##0";
            dataRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            dataRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
        }

        private void ApplyColumnSpecificFormatting(Microsoft.Office.Interop.Excel.Worksheet worksheet,
            int rowStart, int rowEnd, int columnStart, int columnEnd)
        {
            // Center align STT column (Column A)
            var sttRange = worksheet.Range[
                worksheet.Cells[rowStart, columnStart],
                worksheet.Cells[rowEnd, columnStart]
            ];
            sttRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            // Left align description column (Column B)
            var descriptionRange = worksheet.Range[
                worksheet.Cells[rowStart, columnStart + 1],
                worksheet.Cells[rowEnd, columnStart + 1]
            ];
            descriptionRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            descriptionRange.Rows.AutoFit();
        }

        private void SetCellFormat(Microsoft.Office.Interop.Excel.Range range, string value,
            bool mergeCells, bool bold, int fontSize,
            Microsoft.Office.Interop.Excel.XlHAlign horizontalAlignment)
        {
            if (!string.IsNullOrEmpty(value))
                range.Value2 = value;

            if (mergeCells)
                range.MergeCells = true;

            range.Font.Name = "Times New Roman";
            range.Font.Bold = bold;
            range.Font.Size = fontSize;
            range.HorizontalAlignment = horizontalAlignment;
        }

        private void SaveExcelFile(Microsoft.Office.Interop.Excel.Workbook workbook,
            Microsoft.Office.Interop.Excel.Application excelApp,
            string rfdate, string rtdate, DateTime now)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "Save Excel File";
                saveFileDialog.FileName = $"VAT refund monthly report from {rfdate} to {rtdate} {now:hhmmss tt}.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                }
            }
        }

        private void CleanupExcelObjects(Microsoft.Office.Interop.Excel.Worksheet worksheet,
            Microsoft.Office.Interop.Excel.Workbook workbook,
            Microsoft.Office.Interop.Excel.Application excelApp)
        {
            try
            {
                // Release worksheet
                if (worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    worksheet = null;
                }

                // Close and release workbook
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                // Quit and release Excel application
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                }

                // Force garbage collection to clean up COM objects
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                // Log cleanup error if needed
                System.Diagnostics.Debug.WriteLine($"Error during Excel cleanup: {ex.Message}");
            }
        }       
        private bool TryParseReportDates(out DateTime start, out DateTime end)
        {
            // Local helper to parse individual fields
            bool TryGetDate(System.Windows.Forms.TextBox tb, DateTimePicker dp, out DateTime dt)
            {
                string s = tb.Text?.Trim() ?? string.Empty;

                // Attempt to parse the text box input first
                if (!string.IsNullOrWhiteSpace(s) &&
                    DateTime.TryParseExact(s, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                {
                    return true;
                }

                // Fallback to the DateTimePicker value if text parsing fails
                if (dp != null)
                {
                    dt = dp.Value;
                    return true;
                }

                dt = DateTime.Now;
                return false;
            }

            // Execute parsing for both Start (From) and End (To) dates
            bool startValid = TryGetDate(txtRfdate, dtpRfdate, out start);
            bool endValid = TryGetDate(txtRtdate, dtpRtdate, out end);

            if (!startValid || !endValid)
            {
                MessageBox.Show("Please enter valid dates in the format dd/MM/yyyy.",
                                "Invalid Date Input",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
        public void ExportReportToOpenOffice(System.Data.DataTable dataTable, string sheetName, string title)
    {
        try
        {
            // 1. Date Logic
            if (!TryParseReportDates(out DateTime start, out DateTime end)) return;

            string rfdate = start.ToString("dd-MM-yyyy");
            string rtdate = end.ToString("dd-MM-yyyy");
            DateTime now = DateTime.Now;

            // 2. Initialize Workbook
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Set Global Default Font
                worksheet.Style.Font.FontName = "Times New Roman";

                // 3. Setup Page & Report Structure
                SetupPageFormatting(worksheet);
                CreateHeaders(worksheet, title, now);
                CreateColumnHeaders(worksheet);

                // 4. Export Data
                ExportDataTableToWorksheet(dataTable, worksheet);

                // 5. Footer Signatures (Dynamic positioning)
                int signatureStartRow = 10 + dataTable.Rows.Count + 3;
                CreateFooterSignatures(worksheet, signatureStartRow);

                // 6. Save Dialog
                SaveReport(workbook, rfdate, rtdate, now);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

        #region Helper Methods

        private void SetupPageFormatting(IXLWorksheet worksheet)
        {
            var ps = worksheet.PageSetup;
            ps.PageOrientation = XLPageOrientation.Landscape;
            ps.PaperSize = XLPaperSize.A4Paper;
            ps.Margins.SetTop(0.75).SetBottom(0.75).SetLeft(0.75).SetRight(0.5);
        }

        private void CreateHeaders(IXLWorksheet worksheet, string title, DateTime now)
        {
            // Department Info
            var headerRange = worksheet.Range("A1:C1").Merge();
            headerRange.Value = "CHI CỤC HẢI QUAN KHU VỰC II";
            headerRange.Style.Font.SetBold();
            headerRange.Style.Font.FontSize = 13;
            headerRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            var subHeader = worksheet.Range("A2:C2").Merge();
            subHeader.Value = "HẢI QUAN CỬA KHẨU\nSÂN BAY QUỐC TẾ TÂN SƠN NHẤT";
            subHeader.Style.Font.SetBold();
            subHeader.Style.Font.FontSize = 13;
            subHeader.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            subHeader.Style.Alignment.SetWrapText();
            worksheet.Row(2).Height = 29;

            //// 1. Create a very thin 'spacer' row
            //worksheet.Row(3).Height = 2;

            //// 2. Draw the line by coloring the top border of a specific segment
            //// This creates a short line in the middle of your header area
            //var shortLine = worksheet.Range("A3:C3");
            //shortLine.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            //shortLine.Style.Border.TopBorderColor = XLColor.Black;

            // Appendix
            worksheet.Cell("H1").SetValue("Phụ lục I");
            worksheet.Cell("H1").Style.Font.SetBold();
            worksheet.Cell("H1").Style.Font.Italic = true;
            worksheet.Cell("H1").Style.Font.FontSize = 14;

            // Main Titles
            worksheet.Range("A5:J5").Merge().SetValue(title).Style.Font.SetBold();
            worksheet.Range("A5:J5").Style.Font.FontSize = 14;
            worksheet.Range("A5:J5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            string subTitleStr = $"(Đính kèm công văn số:        /BC-CKTSN ngày     /{now.Month:00}/{now.Year:0000})";
            worksheet.Range("A6:J6").Merge().SetValue(subTitleStr);
            worksheet.Range("A6:J6").Style.Font.FontSize = 14;
            worksheet.Range("A6:J6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private void CreateColumnHeaders(IXLWorksheet worksheet)
        {
            // Set Column Widths
            double[] widths = { 5, 21, 15, 15, 15, 15, 12, 12, 8, 8 };
            for (int i = 0; i < widths.Length; i++) worksheet.Column(i + 1).Width = widths[i];

            // Labels & Merging
            worksheet.Range("A8:A9").Merge().SetValue("STT");
            worksheet.Range("B8:B9").Merge().SetValue("MẶT HÀNG");

            worksheet.Range("C8:D8").Merge().SetValue("TRỊ GIÁ HÀNG HOÀN THUẾ (ĐỒNG)");
            worksheet.Cell("C9").Value = "Trong tháng";
            worksheet.Cell("D9").Value = "Lũy kế đến kỳ báo cáo";

            worksheet.Range("E8:F8").Merge().SetValue("SỐ TIỀN THUẾ GTGT ĐƯỢC HOÀN (ĐỒNG)");
            worksheet.Cell("E9").Value = "Trong tháng";
            worksheet.Cell("F9").Value = "Lũy kế đến kỳ báo cáo";

            worksheet.Range("G8:H8").Merge().SetValue("SỐ TIỀN DỊCH VỤ NGÂN HÀNG ĐƯỢC HƯỞNG (ĐỒNG)");
            worksheet.Cell("G9").Value = "Trong tháng";
            worksheet.Cell("H9").Value = "Lũy kế đến kỳ báo cáo";

            worksheet.Range("I8:J8").Merge().SetValue("SỐ LƯỢNG NGƯỜI NƯỚC NGOÀI (NGƯỜI)");
            worksheet.Cell("I9").Value = "Trong tháng";
            worksheet.Cell("J9").Value = "Lũy kế đến kỳ báo cáo";

            // Header Styling
            var headerBlock = worksheet.Range("A8:J9");
            headerBlock.Style.Font.SetBold();
            headerBlock.Style.Font.FontSize = 11;
            headerBlock.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerBlock.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerBlock.Style.Alignment.WrapText = true;
            headerBlock.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerBlock.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            worksheet.Row(8).Height = 60;
            worksheet.Row(9).Height = 50;
        }
        private void SaveReport(XLWorkbook workbook, string rfdate, string rtdate, DateTime now)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                // Define the file type and default extension
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.DefaultExt = "xlsx";

                // Construct a clean, descriptive filename               
                string fileName = $"VAT_Refund_Monthly_Report_{rfdate}_to_{rtdate}_{now:HHmmss}.xlsx";
                saveFileDialog.FileName = fileName;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Save the file to the chosen path
                        workbook.SaveAs(saveFileDialog.FileName);

                        MessageBox.Show("Report exported successfully!", "Export Success",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Optional: Automatically open the file after saving
                        // System.Diagnostics.Process.Start(saveFileDialog.FileName);
                    }
                    catch (IOException ex)
                    {
                        // Handle cases where the file is already open in Excel/OpenOffice
                        MessageBox.Show("Cannot save file. Please close the file if it is currently open in another program and try again.\n\n" + ex.Message,
                                        "File Access Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void ExportDataTableToWorksheet(System.Data.DataTable dt, IXLWorksheet worksheet)
    {
        if (dt.Rows.Count == 0) return;

            int rowStart = 10;
            worksheet.Cell(rowStart, 1).InsertData(dt);

            var dataRange = worksheet.Range(rowStart, 1, rowStart + dt.Rows.Count - 1, dt.Columns.Count);
            //Force a specific range to be numeric(Columns C to J)            
            foreach (var cell in dataRange.Cells())
            {
                if (decimal.TryParse(cell.Value.ToString(), out decimal val))
                {
                    cell.Value = val; // Re-assigning as a numeric type
                }
            }            
            // Apply Thousand Separator & General Formatting
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Font.SetFontSize(12);
            dataRange.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            dataRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            dataRange.Style.Alignment.WrapText = true;

            // Thousand Separator Format: #,##0
            dataRange.Style.NumberFormat.Format = "#,##0";

        ApplyColumnSpecificFormatting(worksheet, rowStart, rowStart + dt.Rows.Count - 1);
    }

        private void ApplyColumnSpecificFormatting(IXLWorksheet ws, int start, int end)
        {
            ws.Range(start, 1, end, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center); // STT
            ws.Range(start, 2, end, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);   // Name
            ws.Row(10).Height = 70; // Optional: Set a default row height for better appearance
            ws.Column(2).Width = 20;
            ws.Style.Alignment.WrapText = true;    
            ws.Style.Font.FontSize = 11;
        }

    private void CreateFooterSignatures(IXLWorksheet worksheet, int row)
    {
        string userName = ((Form)this.MdiParent).Controls["lblUserName"].Text;

            // Left Side - Maker
            worksheet.Range(row, 2, row, 3).Merge().SetValue("NGƯỜI LẬP").Style.Font.SetBold();
            worksheet.Range(row, 2, row, 3).Style.Font.FontSize = 13;
            worksheet.Range(row, 2, row, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Range(row, 2, row, 3).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);           

            worksheet.Range(row + 5, 2, row + 5, 3).Merge().SetValue(userName).Style.Font.SetBold();
            worksheet.Range(row + 5, 2, row + 5, 3).Style.Font.FontSize = 14;
            worksheet.Range(row + 5, 2, row + 5, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Range(row + 5, 2, row + 5, 3).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            worksheet.Row(row + 5).Height = 20; // Optional: Increase row height for better appearance

            // Right Side - Signatories
            worksheet.Range(row - 1, 7, row - 1, 8).Merge().SetValue("KT. ĐỘI TRƯỞNG").Style.Font.SetBold();
            worksheet.Range(row - 1, 7, row - 1, 8).Style.Font.FontSize = 13;
            worksheet.Range(row - 1, 7, row - 1, 8).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Range(row - 1, 7, row - 1, 8).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

            worksheet.Range(row, 7, row, 8).Merge().SetValue("PHÓ ĐỘI TRƯỞNG").Style.Font.SetBold();
            worksheet.Range(row, 7, row, 8).Style.Font.FontSize = 13;
            worksheet.Range(row, 7, row, 8).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            worksheet.Range(row, 7, row, 8).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

    #endregion
    private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (chbMonthlyReport.Checked == true)
                {
                    System.Data.DataTable dataTable = new System.Data.DataTable();

                    DataColumn col1 = new DataColumn("STT");
                    DataColumn col2 = new DataColumn("Motahanghoa");

                    DataColumn col3 = new DataColumn("TongTrigiaHHchuaVAT (VND)");
                    DataColumn col4 = new DataColumn("TongLKTrigiaHHchuaVAT (VND)");

                    DataColumn col5 = new DataColumn("TongSotienVATDH (VND)");
                    DataColumn col6 = new DataColumn("TongLKSotienVATDH (VND)");

                    DataColumn col7 = new DataColumn("TongSotienDVNHH (VND)");
                    DataColumn col8 = new DataColumn("TongLKSotienDVNHH (VND)");

                    DataColumn col9 = new DataColumn("TongLuotHK (HK)");
                    DataColumn col10 = new DataColumn("TongLKLuotHK (HK)");

                    dataTable.Columns.Add(col1);
                    dataTable.Columns.Add(col2);
                    dataTable.Columns.Add(col3);
                    dataTable.Columns.Add(col4);
                    dataTable.Columns.Add(col5);
                    dataTable.Columns.Add(col6);
                    dataTable.Columns.Add(col7);
                    dataTable.Columns.Add(col8);
                    dataTable.Columns.Add(col9);
                    dataTable.Columns.Add(col10);

                    foreach (DataGridViewRow dtgvRow in dataGridView1.Rows)
                    {
                        DataRow dtrow = dataTable.NewRow();

                        dtrow[0] = dtgvRow.Cells[0].Value;
                        dtrow[1] = dtgvRow.Cells[1].Value;
                        dtrow[2] = dtgvRow.Cells[2].Value;
                        dtrow[3] = dtgvRow.Cells[3].Value;
                        dtrow[4] = dtgvRow.Cells[4].Value;
                        dtrow[5] = dtgvRow.Cells[5].Value;
                        dtrow[6] = dtgvRow.Cells[6].Value;
                        dtrow[7] = dtgvRow.Cells[7].Value;
                        dtrow[8] = dtgvRow.Cells[8].Value;
                        dtrow[9] = dtgvRow.Cells[9].Value;

                        dataTable.Rows.Add(dtrow);
                    }
                    if (rdoOpenOffice.Checked)
                    {
                        ExportReportToOpenOffice(dataTable, "Sheet1", "BÁO CÁO TÌNH HÌNH HOÀN THUẾ GTGT CHO NGƯỜI NƯỚC NGOÀI XUẤT CẢNH");
                    }
                    else
                    {
                        ExportReport(dataTable, "Sheet1", "BÁO CÁO TÌNH HÌNH HOÀN THUẾ GTGT CHO NGƯỜI NƯỚC NGOÀI XUẤT CẢNH");
                    }
                }
                else // chbMonthlyReport.Check == false
                {
                    if (rdoOpenOffice.Checked)
                    {
                        ExportSpreadsheet();
                    }
                    else
                    {
                        ExportExcel();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            try
            {
                Utility ut = new Utility();
                var conn = ut.OpenDB();

                string loginName = ((Form)this.MdiParent).Controls["lblLoginName"].Text;
                //string loginName = "HQ10-0152";

                btnUndo.Enabled = false;
                string soHD = txtSoHD.Text.Trim();
                string soHC = txtSoHC.Text.Trim();
                string rfDate = txtRfdate.Text.Trim();
                string rtDate = txtRtdate.Text.Trim();
                DateTime ngayHientai = DateTime.Now;
                var ngayhientai = ngayHientai.ToString("yyyy-MM-dd hh:mm:ss tt");

                DialogResult dr = MessageBox.Show("Are you sure to delete data?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();

                        using (var trans = conn.BeginTransaction())
                        {
                            using (SqlCommand cmd = new SqlCommand("", conn, trans))
                            {
                                cmd.CommandText = "SET ANSI_WARNINGS OFF";
                                cmd.ExecuteNonQuery();

                                // Clear temporary tables
                                cmd.CommandText = "Delete From tmpBC03";
                                cmd.ExecuteNonQuery();
                                cmd.CommandText = "Delete From tmpBC04";
                                cmd.ExecuteNonQuery();
                                cmd.CommandText = "Delete From tmpBC09";
                                cmd.ExecuteNonQuery();

                                if (txtRfdate.TextLength != 0 && txtRtdate.TextLength != 0 && string.IsNullOrEmpty(soHD) == true)
                                {
                                    DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");
                                    DateTime prtdate = DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    var rtdate = prtdate.ToString("yyyy-MM-dd hh:mm:ss tt");

                                    // Get the earliest import date in the range
                                    cmd.CommandText = "Select Top 1 NgaynhapHT " +
                                         "From BC09 " +
                                         "Where NgayHT Between N'" + rfdate + "' And N'" + rtdate + "' " +
                                         "And LoginName = N'" + loginName + "' " + // ADDED: Restrict to user's data
                                         "Order by NgaynhapHT asc";
                                    var result = cmd.ExecuteScalar();

                                    if (result == null || result == DBNull.Value)
                                    {
                                        MessageBox.Show("No data found for deletion or you don't have permission to delete this data.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        trans.Rollback();
                                        return;
                                    }

                                    DateTime pngaynhapht = Convert.ToDateTime(result);
                                    var ngaynhapht = pngaynhapht.ToString("yyyy-MM-dd");
                                    DateTime pnextngaynhapht = pngaynhapht.AddDays(1);
                                    var nextngaynhapht = pnextngaynhapht.ToString("yyyy-MM-dd");

                                    // Backup data to temporary tables with LoginName check
                                    cmd.CommandText = "Insert Into tmpBC03 (ThoigianGD, MasoGD, HotenHK, SotienVNDHT, Ghichu, NgaynhapHT, LoginName) " +
                                                    "Select ThoigianGD, MasoGD, HotenHK, SotienVNDHT, Ghichu, NgaynhapHT, LoginName " +
                                                    "From BC03 " +
                                                    "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'"; // ADDED: Only user's data
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "Insert Into tmpBC04 " +
                                        "Select KyhieuSoNgay, MasoGD, TenDNBH, SotienVATHD, NgayHT, Ghichu, NgaynhapHT, LoginName " +
                                                    "From BC04 " +
                                                    "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'"; // ADDED: Only user's data
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "Insert Into tmpBC09 " +
                                            "Select KyhieuHD, SoHD, NgayHD, TenDNBH, MasoDN, HoTenHK, SoHC, NgayHC, Quoctich, TrigiaHHchuaVAT, NgayHT, SotienVATDH, " +
                                            "SotienDVNHH, Ghichu, NgaynhapHT, LoginName " +
                                                        "From BC09 " +
                                                        "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                        "And LoginName = N'" + loginName + "'"; // ADDED: Only user's data
                                    cmd.ExecuteNonQuery();

                                    // Update temporary tables with current user info
                                    cmd.CommandText = "Update tmpBC03 " +
                                        "Set NgaynhapHT = N'" + ngaynhapht + "', LoginName = N'" + loginName + "' " +
                                        "Where LoginName = N'" + loginName + "'";
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "Update tmpBC04 " +
                                        "Set NgaynhapHT = N'" + ngaynhapht + "', LoginName = N'" + loginName + "' " +
                                        "Where LoginName = N'" + loginName + "'";
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "Update tmpBC09 " +
                                        "Set NgaynhapHT = N'" + ngaynhapht + "', LoginName = N'" + loginName + "' " +
                                        "Where LoginName = N'" + loginName + "'";
                                    cmd.ExecuteNonQuery();

                                    // Count user's data before deletion
                                    cmd.CommandText = "Select Count(SoHD) From BC09 " +
                                                    "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'";
                                    int countbfdel = Convert.ToInt32(cmd.ExecuteScalar());

                                    // Delete only user's data
                                    cmd.CommandText = "Delete From BC03 " +
                                                    "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'";
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "Delete From BC04 " +
                                                    "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'";
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "Delete From BC09 " +
                                                    "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'";
                                    cmd.ExecuteNonQuery();

                                    // Count user's data after deletion
                                    cmd.CommandText = "Select Count(SoHD) From BC09 " +
                                                    "Where (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'";
                                    int countatdel = Convert.ToInt32(cmd.ExecuteScalar());

                                    int delrows = countbfdel - countatdel;

                                    if (delrows > 0)
                                    {
                                        MessageBox.Show("Deleted " + delrows + " datarow(s) successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                        // Start undo timer
                                        timer.Interval = 60000;
                                        timer.Tick += Timer_Tick;
                                        timer.Start();
                                        btnUndo.Enabled = true;

                                        // Update total rows display
                                        cmd.CommandText = "Select Count(SoHD) From BC09";
                                        txtTotalrows.Text = Convert.ToString(Convert.ToInt32(cmd.ExecuteScalar()).ToString("#,##0"));
                                        this.Refresh();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Delete data failed or no data found for your account.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    }
                                }
                                else if (txtRfdate.TextLength != 0 && txtRtdate.TextLength != 0 && string.IsNullOrEmpty(soHD) == false)
                                {
                                    DateTime prfdate = DateTime.ParseExact(rfDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    var rfdate = prfdate.ToString("yyyy-MM-dd hh:mm:ss tt");
                                    DateTime prtdate = DateTime.ParseExact(rtDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    var rtdate = prtdate.ToString("yyyy-MM-dd hh:mm:ss tt");

                                    // Get the import date with LoginName check
                                    cmd.CommandText = "Select Top 1 NgaynhapHT " +
                                         "From BC09 " +
                                         "Where NgayHT Between N'" + rfdate + "' And N'" + rtdate + "' " +
                                         "And SoHD = N'" + soHD + "' And SoHC = N'" + soHC + "' " +
                                         "And LoginName = N'" + loginName + "' " + // ADDED: Restrict to user's data
                                         "Order by NgaynhapHT asc";
                                    var result = cmd.ExecuteScalar();

                                    if (result == null || result == DBNull.Value)
                                    {
                                        MessageBox.Show("No data found for deletion or you don't have permission to delete this data.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                        trans.Rollback();
                                        return;
                                    }

                                    DateTime pngaynhapht = Convert.ToDateTime(result);
                                    var ngaynhapht = pngaynhapht.ToString("yyyy-MM-dd");
                                    DateTime pnextngaynhapht = pngaynhapht.AddDays(1);
                                    var nextngaynhapht = pnextngaynhapht.ToString("yyyy-MM-dd");

                                    // Backup specific record to temporary table with LoginName check
                                    cmd.CommandText = "Insert Into tmpBC09 " +
                                            "Select KyhieuHD, SoHD, NgayHD, TenDNBH, MasoDN, HoTenHK, SoHC, NgayHC, Quoctich, TrigiaHHchuaVAT, NgayHT, SotienVATDH, " +
                                            "SotienDVNHH, Ghichu, NgaynhapHT, LoginName " +
                                                        "From BC09 " +
                                                        "Where SoHD = N'" + soHD + "' And SoHC = N'" + soHC + "' " +
                                                        "And (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                        "And LoginName = N'" + loginName + "'"; // ADDED: Only user's data
                                    cmd.ExecuteNonQuery();

                                    // Update temporary table
                                    cmd.CommandText = "Update tmpBC09 " +
                                        "Set NgaynhapHT = N'" + ngaynhapht + "', LoginName = N'" + loginName + "' " +
                                        "Where LoginName = N'" + loginName + "'";
                                    cmd.ExecuteNonQuery();

                                    // Count specific record before deletion
                                    cmd.CommandText = "Select Count(SoHD) From BC09 " +
                                                    "Where SoHD = N'" + soHD + "' And SoHC = N'" + soHC + "' " +
                                                    "And (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'";
                                    int countbfdel = Convert.ToInt32(cmd.ExecuteScalar());

                                    // Delete specific record with LoginName check
                                    cmd.CommandText = "Delete From BC09 " +
                                                    "Where SoHD = N'" + soHD + "' And SoHC = N'" + soHC + "' " +
                                                    "And (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'"; // ADDED: Only user's data
                                    cmd.ExecuteNonQuery();

                                    // Count after deletion
                                    cmd.CommandText = "Select Count(SoHD) From BC09 " +
                                                    "Where SoHD = N'" + soHD + "' And SoHC = N'" + soHC + "' " +
                                                    "And (NgaynhapHT between N'" + ngaynhapht + "' and N'" + nextngaynhapht + "') " +
                                                    "And LoginName = N'" + loginName + "'";
                                    int countatdel = Convert.ToInt32(cmd.ExecuteScalar());

                                    int delrows = countbfdel - countatdel;

                                    if (delrows > 0)
                                    {
                                        MessageBox.Show("Deleted " + delrows + " datarow(s) successfully.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                        // Start undo timer
                                        timer.Interval = 60000;
                                        timer.Tick += Timer_Tick;
                                        timer.Start();
                                        btnUndo.Enabled = true;

                                        // Update total rows display
                                        cmd.CommandText = "Select Count(SoHD) From BC09";
                                        txtTotalrows.Text = Convert.ToString(Convert.ToInt32(cmd.ExecuteScalar()).ToString("#,##0"));
                                        this.Refresh();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Delete data failed or no data found for your account.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Tax refund date cannot be null. Please define the tax refund date.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }

                                cmd.CommandText = "SET ANSI_WARNINGS ON";
                                cmd.ExecuteNonQuery();
                                trans.Commit();
                            }
                            conn.Close();
                            conn.Dispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Datarows deleted unsuccessfully: " + ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void btnUndo_Click(object sender, EventArgs e)
        {
            try
            {
                Utility ut = new Utility();
                var conn = ut.OpenDB();
                string loginName = ((Form)this.MdiParent).Controls["lblLoginName"].Text;
                //string loginName = "HQ10-0152";

                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();

                    //Count suspect customs papers before recovery
                    cmd = new SqlCommand("Select Count(SoHD) From BC09 Where LoginName = N'" + loginName + "'", conn);
                    Int32 count_bs = Convert.ToInt32(cmd.ExecuteScalar());

                    //Import data into temporary table
                    cmd = new SqlCommand("Insert Into BC03 " +
                        "Select ThoigianGD, MasoGD, HotenHK, SotienVNDHT, Ghichu, NgaynhapHT, LoginName " +
                                    "From tmpBC03 " +
                                    "Where LoginName = N'" + loginName + "'", conn);
                    cmd.ExecuteNonQuery();

                    cmd = new SqlCommand("Insert Into BC04 " +
                        "Select KyhieuSoNgay, MasoGD, TenDNBH, SotienVATHD, NgayHT, Ghichu, NgaynhapHT, LoginName " +
                                    "From tmpBC04 " +
                                    "Where LoginName = N'" + loginName + "'", conn);
                    cmd.ExecuteNonQuery();

                    cmd = new SqlCommand("Insert Into BC09 " +
                        "Select KyhieuHD, SoHD, NgayHD, TenDNBH, MasoDN, HoTenHK, SoHC, NgayHC, Quoctich, TrigiaHHchuaVAT, NgayHT, SotienVATDH, " +
                        "SotienDVNHH, Ghichu, NgaynhapHT, LoginName " +
                                    "From tmpBC09 " +
                                    "Where LoginName = N'" + loginName + "'", conn);
                    cmd.ExecuteNonQuery();

                    cmd = new SqlCommand("Select Count(SoHD) From BC09 Where LoginName = N'" + loginName + "'", conn);
                    int count_as = Convert.ToInt32(cmd.ExecuteScalar());
                    int count = count_as - count_bs;

                    if (count > 0)
                    {
                        MessageBox.Show("Recovery successfully" + " " + count + " " + "datarow(s).", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData();
                        btnUndo.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("Data recovery failed.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                //Dong ket noi
                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //dataGridView1.CellFormatting += (sender, e) => ApplyConditionalFormatting(e);
        }

        private void txtSoHC_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = char.ToUpper(e.KeyChar);
        }

        /// <summary>
        /// Returns the cursor position within the unformatted (separator-free) text based on the current selection in the TextBox.
        /// </summary>
        /// <param name="t">The TextBox to inspect.</param>
        /// <returns>The zero-based cursor index into the unformatted text (characters excluding '.' and ',').</returns>
        private int GetUnformattedCursorPosition(System.Windows.Forms.TextBox t)
        {
            if (t == null) return 0;
            int sel = Math.Max(0, Math.Min(t.SelectionStart, t.Text.Length));
            string before = t.Text.Substring(0, sel);
            string unformattedBefore = before.Replace(",", "").Replace(".", "");
            return unformattedBefore.Length;
        }

        /// <summary>
        /// Handles key presses inside the minimum refund amount TextBox. Formats numeric input with thousand separators.
        /// Accepts digits and the minus sign. Blocks other characters except Backspace/Delete (which are handled in KeyDown).
        /// </summary>
        /// <param name="sender">The TextBox (txtminsotienVNDHT).</param>
        /// <param name="e">KeyPress event arguments.</param>
        private void txtminsotienVNDHT_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Allow control keys (carriage return, etc.) to pass through by default.
            if (char.IsControl(e.KeyChar) && e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
            {
                return;
            }

            // Accept digits and '-' sign only for this handler (Backspace/Delete handled in KeyDown)
            if (char.IsDigit(e.KeyChar) || e.KeyChar == '-')
            {
                System.Windows.Forms.TextBox t = (System.Windows.Forms.TextBox)sender;

                // Unformatted current text (remove thousand separators)
                string currentUnformatted = t.Text.Replace(",", "").Replace(".", "");

                // Determine if current value is negative
                bool isNegative = currentUnformatted.StartsWith("-");
                if (isNegative)
                {
                    currentUnformatted = currentUnformatted.Substring(1);
                }

                // Compute unformatted cursor index
                int unformattedCursor = GetUnformattedCursorPosition(t);

                // Build left and right parts based on unformatted cursor
                unformattedCursor = Math.Max(0, Math.Min(unformattedCursor, currentUnformatted.Length));
                string leftPart = currentUnformatted.Substring(0, unformattedCursor);
                string rightPart = currentUnformatted.Substring(unformattedCursor);

                string newUnformatted;

                // Handle toggling negative sign if user typed '-'
                if (e.KeyChar == '-')
                {
                    isNegative = !isNegative;
                    newUnformatted = leftPart + rightPart; // negative sign handled via isNegative flag
                }
                else
                {
                    // Insert digit at the unformatted cursor
                    newUnformatted = leftPart + e.KeyChar + rightPart;
                }

                // Limit raw length to a safe value (prevent overflow/excessive input)
                if (newUnformatted.Length > 20)
                {
                    e.Handled = true;
                    return;
                }

                // Parse numeric value safely
                if (decimal.TryParse(newUnformatted.Length == 0 ? "0" : newUnformatted, out decimal value))
                {
                    if (isNegative && value != 0) value = -value;

                    // Format using Vietnamese culture to ensure '.' grouping separator
                    var viVN = System.Globalization.CultureInfo.GetCultureInfo("vi-VN");
                    t.Text = value.ToString("N0", viVN);

                    // Place caret at end - robust and avoids mapping issues between formatted/unformatted text
                    t.SelectionStart = t.Text.Length;
                }
                else
                {
                    // If parsing failed, block the key
                    e.Handled = true;
                }

                // We've handled this keystroke
                e.Handled = true;
            }
            else
            {
                // Block non-digit/non-control characters
                if (e.KeyChar != (char)Keys.Back && e.KeyChar != (char)Keys.Delete)
                {
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// Handles key down events for the minimum refund amount TextBox to support Backspace/Delete behavior on the unformatted value.
        /// This avoids inconsistencies caused by formatting characters and keeps the value valid and formatted after deletion.
        /// </summary>
        /// <param name="sender">The TextBox (txtminsotienVNDHT).</param>
        /// <param name="e">KeyEvent arguments.</param>
        private void txtminsotienVNDHT_KeyDown(object sender, KeyEventArgs e)
        {
            // Only special-case Backspace and Delete here; other keys are handled in KeyPress or by default behavior.
            if (e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete)
            {
                System.Windows.Forms.TextBox t = (System.Windows.Forms.TextBox)sender;

                // Current unformatted text
                string unformatted = t.Text.Replace(",", "").Replace(".", "");

                // Track if negative
                bool isNegative = unformatted.StartsWith("-");
                if (isNegative) unformatted = unformatted.Substring(1);

                // Compute unformatted cursor index
                int unformattedCursor = GetUnformattedCursorPosition(t);
                unformattedCursor = Math.Max(0, Math.Min(unformattedCursor, unformatted.Length));

                string leftPart = unformatted.Substring(0, unformattedCursor);
                string rightPart = unformatted.Substring(unformattedCursor);

                string newUnformatted = unformatted;

                if (e.KeyCode == Keys.Back)
                {
                    if (leftPart.Length > 0)
                    {
                        leftPart = leftPart.Remove(leftPart.Length - 1);
                        newUnformatted = leftPart + rightPart;
                    }
                    else
                    {
                        // Nothing to delete on backspace
                        newUnformatted = rightPart;
                    }
                }
                else // Delete
                {
                    if (rightPart.Length > 0)
                    {
                        rightPart = rightPart.Remove(0, 1);
                        newUnformatted = leftPart + rightPart;
                    }
                    else
                    {
                        // Nothing to delete
                        newUnformatted = leftPart;
                    }
                }

                // Parse and reformat
                if (decimal.TryParse(string.IsNullOrEmpty(newUnformatted) ? "0" : newUnformatted, out decimal value))
                {
                    if (isNegative && value != 0) value = -value;

                    var viVN = System.Globalization.CultureInfo.GetCultureInfo("vi-VN");
                    t.Text = value.ToString("N0", viVN);

                    // Place caret at end - robust
                    t.SelectionStart = t.Text.Length;
                }

                // Suppress further key processing to avoid double-delete
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }

        /// <summary>
        /// Called when the minimum refund amount TextBox receives focus.
        /// Removes formatting (thousand separators) to allow user to type an unformatted numeric string, and selects all text.
        /// </summary>
        /// <param name="sender">The TextBox control.</param>
        /// <param name="e">Event arguments.</param>
        private void txtminsotienVNDHT_Enter(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox t = (System.Windows.Forms.TextBox)sender;

            // Remove thousand separators so user is editing the raw number
            t.Text = t.Text.Replace(".", "").Replace(",", "");

            // Select all for quick replacement
            t.SelectAll();
        }

        /// <summary>
        /// Called when the minimum refund amount TextBox loses focus.
        /// Ensures the content is a valid number and formats with thousand separators using Vietnamese culture.
        /// If the value is empty or invalid, it is set to "0".
        /// </summary>
        /// <param name="sender">The TextBox control.</param>
        /// <param name="e">Event arguments.</param>
        private void txtminsotienVNDHT_Leave(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox t = (System.Windows.Forms.TextBox)sender;

            // Get unformatted text
            string unformattedText = t.Text.Replace(".", "").Replace(",", "");

            if (decimal.TryParse(string.IsNullOrWhiteSpace(unformattedText) ? "0" : unformattedText, out decimal value))
            {
                System.Globalization.CultureInfo viVN = System.Globalization.CultureInfo.GetCultureInfo("vi-VN");
                t.Text = value.ToString("N0", viVN);
            }
            else if (string.IsNullOrEmpty(t.Text))
            {
                t.Text = "0";
            }
            else
            {
                // Fallback: set to zero for invalid input
                t.Text = "0";
            }
        }
        private void dtpRfdate_CloseUp(object sender, EventArgs e)
        {
            // This event fires when the drop-down calendar is closed.
            // It is a reliable place to update a related control's text.
            try
            {
                // Update the textbox with the selected date in the desired format
                txtRfdate.Text = dtpRfdate.Value.ToString("dd/MM/yyyy");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Notice",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            // You would need to link this method to the control's CloseUp event.
            // You might also keep the ValueChanged event if you need instant update upon *change*.
        }

        private void dtpRtdate_CloseUp(object sender, EventArgs e)
        {
            try
            {
                // Update the textbox with the selected date in the desired format
                txtRtdate.Text = dtpRtdate.Value.ToString("dd/MM/yyyy");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Notice",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
        }
    }
}
    
    
