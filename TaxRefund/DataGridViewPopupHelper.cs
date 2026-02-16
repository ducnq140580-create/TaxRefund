using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

#region
//namespace TaxRefund
//{
//    public class DataGridViewPopupHelper
//    {
//        private readonly Utility _utility;

//        public DataGridViewPopupHelper()
//        {
//            _utility = new Utility();
//        }

//        public void ShowDetailsPopup(string columnName, string value, IWin32Window owner = null)
//        {
//            try
//            {
//                using (var conn = _utility.OpenDB())
//                {
//                    if (conn == null)
//                    {
//                        MessageBox.Show(owner, "Database connection failed.", "Error",
//                            MessageBoxButtons.OK, MessageBoxIcon.Error);
//                        return;
//                    }
//                    if (conn.State == ConnectionState.Closed)
//                    {
//                        conn.Open();
//                    }

//                    string query = "";
//                    string title = "";

//                    if (columnName.Equals("SoHC", StringComparison.OrdinalIgnoreCase))
//                    {
//                        title = $"Passenger Details - SoHC: {value}";
//                        query = @"SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
//                            bc09.SoHD, bc09.NgayHD, bc09.KyhieuHD, bc09.HoTenHK, 
//                            bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
//                            bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH,
//                            bc09.MasoDN, bc09.TenDNBH
//                            FROM BC09 as bc09
//                            WHERE bc09.SoHC = @Value
//                            ORDER BY bc09.NgayHT DESC";
//                    }
//                    else if (columnName.Equals("SoHD", StringComparison.OrdinalIgnoreCase))
//                    {
//                        title = $"Invoice Details - SoHD: {value}";
//                        query = @"SELECT DISTINCT
//                                ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
//                                bc09.SoHD, 
//                                bc09.NgayHD, 
//                                bc09.KyhieuHD,
//                                bc03.ThoigianGD,
//                                bc09.HoTenHK, 
//                                bc09.SoHC, 
//                                bc09.NgayHC, 
//                                bc09.Quoctich, 
//                                bc09.NgayHT,
//                                bc09.TrigiaHHchuaVAT, 
//                                bc09.SotienVATDH, 
//                                bc09.SotienDVNHH,
//                                bc09.MasoDN, 
//                                bc09.TenDNBH
//                            FROM BC09 as bc09
//                            LEFT JOIN BC04 as bc04 ON 
//                                bc09.SoHD = SUBSTRING(
//                                    bc04.KyhieuSoNgay, 
//                                    CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
//                                    CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1)
//                            LEFT JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD  
//                            WHERE bc09.SoHD = @Value AND bc04.KyhieuSoNgay LIKE '%/%/%'";

//                        #region
//                        //title = $"Invoice Details - SoHD: {value}";
//                        //query = @"SELECT DISTINCT
//                        //        ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
//                        //        bc09.SoHD, 
//                        //        bc09.NgayHD, 
//                        //        bc09.KyhieuHD,
//                        //        bc03.ThoigianGD,
//                        //        bc09.HoTenHK, 
//                        //        bc09.SoHC, 
//                        //        bc09.NgayHC, 
//                        //        bc09.Quoctich, 
//                        //        bc09.NgayHT,
//                        //        bc09.TrigiaHHchuaVAT, 
//                        //        bc09.SotienVATDH, 
//                        //        bc09.SotienDVNHH,
//                        //        bc09.MasoDN, 
//                        //        bc09.TenDNBH
//                        //    FROM BC09 as bc09
//                        //    LEFT JOIN BC04 as bc04 ON 
//                        //        bc09.SoHD = SUBSTRING(
//                        //            bc04.KyhieuSoNgay, 
//                        //            CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
//                        //            CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1)                                    
//                        //    RIGHT JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD
//                        //    WHERE bc09.SoHD = @Value AND bc04.KyhieuSoNgay LIKE '%/%/%'"; // GROUP BY bc09.NgayHT";
//                        //    //ORDER BY bc09.NgayHT DESC";
//                        #endregion
//                    }

//                    using (SqlCommand cmd = new SqlCommand(query, conn))
//                    {
//                        cmd.Parameters.AddWithValue("@Value", value);

//                        DataTable dt = new DataTable();
//                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
//                        {
//                            da.Fill(dt);
//                        }

//                        if (dt.Rows.Count > 0)
//                        {
//                            ShowPopupForm(title, dt, owner);
//                        }
//                        else
//                        {
//                            MessageBox.Show(owner, "No additional details found.", "Information",
//                                MessageBoxButtons.OK, MessageBoxIcon.Information);
//                        }
//                    }
//                }
//            }
//            catch (SqlException sqlEx)
//            {
//                MessageBox.Show(owner, $"Database error: {sqlEx.Message}", "Error",
//                    MessageBoxButtons.OK, MessageBoxIcon.Error);
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(owner, $"Error loading details: {ex.Message}", "Error",
//                    MessageBoxButtons.OK, MessageBoxIcon.Error);
//            }            
//        }
//        private static void AutoResizePopupColumns(DataGridView dgv)
//        {
//            if (dgv.Columns.Count == 0) return;

//            // Dictionary of column names and their resize modes
//            var columnResizeRules = new Dictionary<string, DataGridViewAutoSizeColumnMode>
//            {
//                { "STT", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "HoTenHK", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "SoHC", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "Quoctich", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "ThoigianGD", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                //{ "TenDNBH", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            };

//            foreach (DataGridViewColumn column in dgv.Columns)
//            {
//                if (columnResizeRules.TryGetValue(column.Name, out var resizeMode))
//                {
//                    column.AutoSizeMode = resizeMode;

//                    // For fill columns, set minimum width
//                    if (resizeMode == DataGridViewAutoSizeColumnMode.Fill)
//                    {
//                        column.MinimumWidth = 150;
//                    }
//                }
//                else
//                {
//                    // Default behavior for unspecified columns
//                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
//                    column.Width = 120;
//                }
//            }

//            // Set the last column to fill remaining space
//            if (dgv.Columns.Count > 0)
//            {
//                dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
//            }

//            #region
//            //if (dgv.Columns.Count == 0) return;

//            //    // Dictionary of column names and their resize modes
//            //            var columnResizeRules = new Dictionary<string, DataGridViewAutoSizeColumnMode>
//            //    {
//            //        { "STT", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            //        { "HoTenHK", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            //        { "SoHC", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            //        { "Quoctich", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            //         { "ThoigianGD", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            //        { "TenDNBH", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            //    };

//            //foreach (DataGridViewColumn column in dgv.Columns)
//            //{
//            //    if (columnResizeRules.TryGetValue(column.Name, out var resizeMode))
//            //    {
//            //        column.AutoSizeMode = resizeMode;

//            //        // For fill columns, set minimum width
//            //        if (resizeMode == DataGridViewAutoSizeColumnMode.Fill)
//            //        {
//            //            column.MinimumWidth = 150; // Prevent becoming too narrow
//            //        }
//            //    }
//            //    else
//            //    {
//            //        // Default behavior for unspecified columns
//            //        column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
//            //        column.Width = 200; // Fixed width
//            //    }
//            //}

//            //// Optional: Set padding for the last column
//            //dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
//            #endregion

//        }
//        private static void ConfigureGridViewHeaders(DataGridView dgv)
//        {
//            // Font settings
//            var headerFont = new Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

//            // Main header style
//            dgv.EnableHeadersVisualStyles = false;
//            dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
//            {
//                Font = headerFont,
//                ForeColor = Color.DeepSkyBlue,
//                BackColor = Color.WhiteSmoke,
//                Alignment = DataGridViewContentAlignment.MiddleCenter,
//                Padding = new Padding(3),
//            };

//            dgv.ColumnHeadersHeight = 35;

//            #region
//            //// Font settings
//            //var headerFont = new Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

//            //    // Main header style
//            //    dgv.EnableHeadersVisualStyles = false; // Disable system styles
//            //    dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
//            //    {
//            //        Font = headerFont,
//            //        ForeColor = Color.DeepSkyBlue,
//            //        BackColor = Color.WhiteSmoke, // FromArgb(51, 122, 183), // Bootstrap primary blue
//            //        Alignment = DataGridViewContentAlignment.MiddleCenter,
//            //        Padding = new Padding(3),
//            //    };
//            #endregion

//            // Optional: Add gradient background
//            #region
//            //dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
//            //dgv.ColumnHeadersHeight = 50;
//            //dgv.Paint += (sender, e) =>
//            //{
//            //    var headerBounds = new Rectangle(
//            //        0, 0,
//            //        dgv.Width,
//            //        dgv.ColumnHeadersHeight);

//            //    using (var brush = new System.Drawing.Drawing2D.LinearGradientBrush(
//            //        headerBounds,
//            //        Color.FromArgb(70, 130, 180),
//            //        Color.FromArgb(40, 90, 150),
//            //        90f))
//            //    {
//            //        e.Graphics.FillRectangle(brush, headerBounds);
//            //    }
//            //};
//            #endregion
//        }
//        private void ShowPopupForm(string title, DataTable data, IWin32Window owner = null)
//        {
//            var popupForm = new Form
//            {
//                Text = title,
//                StartPosition = FormStartPosition.CenterParent,
//                Width = 1600,
//                Height = 600, // Increased height for better visibility
//                FormBorderStyle = FormBorderStyle.Sizable, // Allow resizing
//                MaximizeBox = true // Allow maximizing
//            };

//            // Create a panel to hold the DataGridView with padding
//            var panel = new Panel
//            {
//                Dock = DockStyle.Fill,
//                Padding = new Padding(10)
//            };

//            var dgv = new DataGridView
//            {
//                Dock = DockStyle.Fill,
//                DataSource = data,
//                ReadOnly = true,
//                AllowUserToAddRows = false,
//                AllowUserToDeleteRows = false,
//                Font = new Font("Microsoft Sans Serif", 9f, FontStyle.Regular),
//                BackgroundColor = SystemColors.Window,
//                ForeColor = SystemColors.ControlText,
//                ColumnHeadersHeight = 35,
//                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None // We'll handle sizing manually
//            };

//            // Apply formatting
//            ConfigureGridViewHeaders(dgv);
//            AutoResizePopupColumns(dgv);
//            ApplyNumberFormatting(dgv);

//            panel.Controls.Add(dgv);
//            popupForm.Controls.Add(panel);

//            popupForm.ShowDialog(owner);

//            #region

//            //var popupForm = new Form
//            //{
//            //    Text = title,
//            //    StartPosition = FormStartPosition.CenterParent,
//            //    Width = 1600,
//            //    Height = 500
//            //};

//            //var dgv = new DataGridView
//            //{
//            //    Dock = DockStyle.Fill,
//            //    DataSource = data,
//            //    ReadOnly = true,
//            //    AllowUserToAddRows = false,
//            //    AllowUserToDeleteRows = false,
//            //    Font = new Font("Microsoft Sans Serif", 10f, FontStyle.Regular),
//            //    BackgroundColor = Color.AntiqueWhite,
//            //    ForeColor = Color.Blue,
//            //    ColumnHeadersHeight = 30
//            //};

//            //// Apply formatting BEFORE showing the dialog
//            //ConfigureGridViewHeaders(dgv);
//            //AutoResizePopupColumns(dgv);

//            //// Apply number formatting if columns exist
//            //ApplyNumberFormatting(dgv);

//            //popupForm.Controls.Add(dgv);
//            //popupForm.ShowDialog(owner);
//            #endregion
//        }

//        private void ApplyNumberFormatting(DataGridView dgv)
//        {
//            if (dgv.Rows.Count == 0 || dgv.Columns.Count == 0) return;

//            DataGridViewCellStyle style = new DataGridViewCellStyle
//            {
//                Format = "N0",
//                Alignment = DataGridViewContentAlignment.MiddleRight
//            };

//            // Check if columns exist before applying formatting
//            string[] numericColumns = { "TrigiaHHchuaVAT", "SotienVATDH", "SotienDVNHH" };

//            foreach (string columnName in numericColumns)
//            {
//                if (dgv.Columns.Contains(columnName))
//                {
//                    dgv.Columns[columnName].DefaultCellStyle = style;
//                }
//            }
//        }
//    }
//}

//namespace TaxRefund
//{
//    public class DataGridViewPopupHelper
//    {
//        private readonly Utility _utility;

//        public DataGridViewPopupHelper()
//        {
//            _utility = new Utility();
//        }

//        public void ShowDetailsPopup(string columnName, string value, IWin32Window owner = null)
//        {
//            try
//            {
//                using (var conn = _utility.OpenDB())
//                {
//                    if (conn == null)
//                    {
//                        MessageBox.Show(owner, "Database connection failed.", "Error",
//                            MessageBoxButtons.OK, MessageBoxIcon.Error);
//                        return;
//                    }
//                    if (conn.State == ConnectionState.Closed)
//                    {
//                        conn.Open();
//                    }

//                    string query = "";
//                    string title = "";

//                    if (columnName.Equals("SoHC", StringComparison.OrdinalIgnoreCase))
//                    {
//                        title = $"Passenger Details - SoHC: {value}";
//                        query = @"SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
//                            bc09.SoHD, bc09.NgayHD, bc09.KyhieuHD, bc09.HoTenHK, 
//                            bc09.SoHC, bc09.NgayHC, bc09.Quoctich, bc09.NgayHT,
//                            bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH,
//                            bc09.MasoDN, bc09.TenDNBH
//                            FROM BC09 as bc09
//                            WHERE bc09.SoHC = @Value
//                            ORDER BY bc09.NgayHT DESC";
//                    }
//                    else if (columnName.Equals("SoHD", StringComparison.OrdinalIgnoreCase))
//                    {
//                        title = $"Invoice Details - SoHD: {value}";
//                        query = @"SELECT DISTINCT
//                                ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
//                                bc09.SoHD, 
//                                bc09.NgayHD, 
//                                bc09.KyhieuHD,
//                                bc03.ThoigianGD,
//                                bc09.HoTenHK, 
//                                bc09.SoHC, 
//                                bc09.NgayHC, 
//                                bc09.Quoctich, 
//                                bc09.NgayHT,
//                                bc09.TrigiaHHchuaVAT, 
//                                bc09.SotienVATDH, 
//                                bc09.SotienDVNHH,
//                                bc09.MasoDN, 
//                                bc09.TenDNBH
//                            FROM BC09 as bc09
//                            LEFT JOIN BC04 as bc04 ON 
//                                bc09.SoHD = SUBSTRING(
//                                    bc04.KyhieuSoNgay, 
//                                    CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
//                                    CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) - 1)
//                            LEFT JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD  
//                            WHERE bc09.SoHD = @Value AND bc04.KyhieuSoNgay LIKE '%/%/%'";
//                    }

//                    using (SqlCommand cmd = new SqlCommand(query, conn))
//                    {
//                        cmd.Parameters.AddWithValue("@Value", value);

//                        System.Data.DataTable dt = new System.Data.DataTable();
//                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
//                        {
//                            da.Fill(dt);
//                        }

//                        if (dt.Rows.Count > 0)
//                        {
//                            ShowPopupForm(title, dt, owner);
//                        }
//                        else
//                        {
//                            MessageBox.Show(owner, "No additional details found.", "Information",
//                                MessageBoxButtons.OK, MessageBoxIcon.Information);
//                        }
//                    }
//                }
//            }
//            catch (SqlException sqlEx)
//            {
//                MessageBox.Show(owner, $"Database error: {sqlEx.Message}", "Error",
//                    MessageBoxButtons.OK, MessageBoxIcon.Error);
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(owner, $"Error loading details: {ex.Message}", "Error",
//                    MessageBoxButtons.OK, MessageBoxIcon.Error);
//            }
//        }

//        private static void ConfigureGridViewHeaders(DataGridView dgv)
//        {
//            var headerFont = new System.Drawing.Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

//            dgv.EnableHeadersVisualStyles = false;
//            dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
//            {
//                Font = headerFont,
//                ForeColor = Color.DeepSkyBlue,
//                BackColor = Color.WhiteSmoke,
//                Alignment = DataGridViewContentAlignment.MiddleCenter,
//                Padding = new Padding(3),
//            };

//            dgv.ColumnHeadersHeight = 35;
//        }

//        private static void AutoResizePopupColumns(DataGridView dgv)
//        {
//            //if (dgv.Columns.Count == 0) return;

//            var columnResizeRules = new Dictionary<string, DataGridViewAutoSizeColumnMode>
//            {
//                { "STT", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "HoTenHK", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "SoHC", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "Quoctich", DataGridViewAutoSizeColumnMode.DisplayedCells },
//                { "ThoigianGD", DataGridViewAutoSizeColumnMode.DisplayedCells },
//            };

//            foreach (DataGridViewColumn column in dgv.Columns)
//            {
//                if (columnResizeRules.TryGetValue(column.Name, out var resizeMode))
//                {
//                    column.AutoSizeMode = resizeMode;

//                    if (resizeMode == DataGridViewAutoSizeColumnMode.Fill)
//                    {
//                        column.MinimumWidth = 150;
//                    }
//                }
//                else
//                {
//                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
//                    column.Width = 120;
//                }
//            }

//            if (dgv.Columns.Count > 0)
//            {
//                dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
//            }
//        }

//        private void ApplyNumberFormatting(DataGridView dgv)
//        {
//            if (dgv.Rows.Count == 0 || dgv.Columns.Count == 0) return;

//            //DataGridViewCellStyle style = new DataGridViewCellStyle
//            //{
//            //    Format = "N0",
//            //    Alignment = DataGridViewContentAlignment.MiddleRight
//            //};

//            //string[] numericColumns = { "TrigiaHHchuaVAT", "SotienVATDH", "SotienDVNHH" };

//            //foreach (string columnName in numericColumns)
//            //{
//            //    if (dgv.Columns.Contains(columnName))
//            //    {
//            //        dgv.Columns[columnName].DefaultCellStyle = style;
//            //    }
//            //}

//            DataGridViewCellStyle numberStyle = new DataGridViewCellStyle
//            {
//                Format = "N0",
//                Alignment = DataGridViewContentAlignment.MiddleRight,
//                BackColor = Color.LightYellow // Optional: highlight numeric columns
//            };
//            // Define numeric columns with their display names
//            var numericColumns = new Dictionary<string, string>
//            {
//                { "TrigiaHHchuaVAT", "Goods Value (ex VAT)" },
//                { "SotienVATDH", "VAT Amount" },
//                { "SotienDVNHH", "Service Fee" }
//            };

//            foreach (var numericCol in numericColumns)
//            {
//                if (dgv.Columns.Contains(numericCol.Key))
//                {
//                    dgv.Columns[numericCol.Key].DefaultCellStyle = numberStyle;
//                    dgv.Columns[numericCol.Key].HeaderText = numericCol.Value; // Update header text
//                    dgv.Columns[numericCol.Key].Width = 150; // Set consistent width for numeric columns
//                }
//            }
//        }

//        // Alternative: Apply formatting based on DataTable column data types
//        private void ApplyNumberFormattingFromDataTable(DataGridView dgv, System.Data.DataTable dt)
//        {
//            //if (dgv.Rows.Count == 0 || dt.Columns.Count == 0) return;

//            DataGridViewCellStyle numberStyle = new DataGridViewCellStyle
//            {
//                Format = "N0",
//                Alignment = DataGridViewContentAlignment.MiddleRight
//            };

//            DataGridViewCellStyle currencyStyle = new DataGridViewCellStyle
//            {
//                Format = "C0",
//                Alignment = DataGridViewContentAlignment.MiddleRight
//            };

//            // Apply formatting based on column names
//            foreach (DataColumn column in dt.Columns)
//            {
//                if (dgv.Columns.Contains(column.ColumnName))
//                {
//                    var dgvColumn = dgv.Columns[column.ColumnName];

//                    // Format specific numeric columns
//                    if (column.ColumnName.Equals("TrigiaHHchuaVAT") ||
//                        column.ColumnName.Equals("SotienVATDH") ||
//                        column.ColumnName.Equals("SotienDVNHH"))
//                    {
//                        dgvColumn.DefaultCellStyle = numberStyle;
//                    }
//                    // You can add more conditions based on data types
//                    else if (column.DataType == typeof(decimal) || column.DataType == typeof(double))
//                    {
//                        dgvColumn.DefaultCellStyle = numberStyle;
//                    }
//                }
//            }
//        }

//        private void ShowPopupForm(string title, System.Data.DataTable data, IWin32Window owner = null)
//        {
//            var popupForm = new Form
//            {
//                Text = title,
//                StartPosition = FormStartPosition.CenterParent,
//                Width = 1600,
//                Height = 600,
//                FormBorderStyle = FormBorderStyle.Sizable,
//                MaximizeBox = true
//            };

//            var panel = new Panel
//            {
//                Dock = DockStyle.Fill,
//                Padding = new Padding(10)
//            };

//            var dgv = new DataGridView
//            {
//                Dock = DockStyle.Fill,
//                DataSource = data,
//                ReadOnly = true,
//                AllowUserToAddRows = false,
//                AllowUserToDeleteRows = false,
//                Font = new System.Drawing.Font("Microsoft Sans Serif", 9f, FontStyle.Regular),
//                BackgroundColor = SystemColors.Window,
//                ForeColor = SystemColors.ControlText,
//                ColumnHeadersHeight = 35,
//            };

//            // Apply formatting methods - NOW INCLUDING ApplyNumberFormatting
//            ConfigureGridViewHeaders(dgv);
//            AutoResizePopupColumns(dgv);
//            ApplyNumberFormatting(dgv); // This method is now properly called
//            //ApplyNumberFormattingFromDataTable(dgv, data);

//            panel.Controls.Add(dgv);
//            popupForm.Controls.Add(panel);
//            popupForm.ShowDialog(owner);
//        }
//    }
//}
#endregion

namespace TaxRefund
{
    // Define a class to hold the query and title data
    internal class QueryInfo
    {
        public string Query { get; set; }
        public string Title { get; set; }
    }

    public class DataGridViewPopupHelper
    {
        private readonly Utility _utility; // Assuming Utility is a class in the project
        private const string DatabaseError = "Database connection failed.";
        private const string NoDetailsFound = "No additional details found.";
      
        public DataGridViewPopupHelper()
        {
            _utility = new Utility() ?? throw new ArgumentNullException(nameof(_utility));
        }

        // --- Public Interface ---
        public void ShowDetailsPopup(string columnName, string value, IWin32Window owner = null)
        {
            try
            {
                // 1. Determine Query and Title
                QueryInfo queryInfo = GetQueryInfo(columnName, value);

                if (string.IsNullOrEmpty(queryInfo.Query))
                {
                    // No query defined for the given columnName, exit silently or show message
                    MessageBox.Show(owner, $"No configuration for column: {columnName}", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 2. Execute Query and Get Data
                System.Data.DataTable data = ExecuteDetailsQuery(queryInfo.Query, value, owner);

                // 3. Display Popup
                if (data != null && data.Rows.Count > 0)
                {
                    ShowPopupForm(queryInfo.Title, data, owner);
                }
                else
                {
                    MessageBox.Show(owner, NoDetailsFound, "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show(owner, $"Database error: {sqlEx.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, $"Error loading details: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- Data Access and Logic Separation ---

        /// <summary>
        /// Selects the appropriate SQL query and title based on the column name.
        /// </summary>

        //private static QueryInfo GetQueryInfo(string columnName, string value)
        //{
        //    string query = "";
        //    string title = "";

        //    switch (columnName.ToUpperInvariant())
        //    {
        //        case "SOHC":
        //            title = $"Passenger Details - SoHC: {value}";
        //            query = @"SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
        //                      bc09.SoHD, 
        //                      bc09.KyhieuHD, bc09.NgayHD, bc09.NgayHT,
        //                      bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, 
        //                      bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH,
        //                      bc09.MasoDN, bc09.TenDNBH
        //                      FROM BC09 as bc09
        //                      WHERE bc09.SoHC = @Value
        //                      ORDER BY bc09.NgayHT DESC";
        //            break;

        //        case "SOHD":
        //            title = $"Invoice Details - SoHD: {value}";
        //            query = @"SELECT DISTINCT
        //                          ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
        //                          bc09.SoHD, bc09.NgayHD, bc09.KyhieuHD,
        //                          bc04.KyhieuSoNgay,                                   
        //                          bc09.NgayHT,
        //                          bc03.ThoigianGD,
        //                          bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, 
        //                          bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH,
        //                          bc09.MasoDN, bc09.TenDNBH
        //                      FROM BC09 as bc09
        //                      INNER JOIN BC04 as bc04 ON bc09.SoHD = SUBSTRING(
        //                          bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1, 
        //                          CHARINDEX('/', bc04.KyhieuSoNgay, CHARINDEX('/', bc04.KyhieuSoNgay) + 1) - CHARINDEX('/', bc04.KyhieuSoNgay) + 1)
        //                      INNER JOIN BC03 as bc03 ON bc04.MasoGD = bc03.MasoGD  
        //                      WHERE bc09.SoHD = @Value AND bc04.KyhieuSoNgay LIKE '%/%/%'";

        //            break;

        //        default:
        //            // query and title remain empty
        //            break;
        //    }

        //    return new QueryInfo { Query = query, Title = title };
        //}

        private static QueryInfo GetQueryInfo(string columnName, string value)
        {
            string query = "";
            string title = "";

            switch (columnName.ToUpperInvariant())
            {
                case "NGAYHT":
                    title = $"Transaction Summary - Date: {value}";
                    // Using the refactored CTE logic for better performance and readability
                    query = @"WITH ParsedBC04 AS (
                        SELECT MasoGD, KyhieuSoNgay,
                            SUBSTRING(KyhieuSoNgay, 
                                CHARINDEX('/', KyhieuSoNgay) + 1, 
                                CHARINDEX('/', KyhieuSoNgay, CHARINDEX('/', KyhieuSoNgay) + 1) - CHARINDEX('/', KyhieuSoNgay) - 1
                            ) AS ExtractedSoHD
                        FROM BC04
                        WHERE KyhieuSoNgay LIKE '%/%/%'
                    )
                    SELECT 
                        ROW_NUMBER() OVER (ORDER BY b9.NgayHT DESC) AS [STT], 
                        b9.SoHD, b9.NgayHD, b9.KyhieuHD, b9.NgayHT,
                        b3.ThoigianGD, b9.HoTenHK, b9.SoHC, b9.Quoctich, 
                        b9.TrigiaHHchuaVAT, b9.SotienVATDH, b9.SotienDVNHH,
                        b9.MasoDN, b9.TenDNBH
                    FROM BC09 AS b9
                    LEFT JOIN ParsedBC04 AS b4 ON b9.SoHD = b4.ExtractedSoHD
                    LEFT JOIN BC03 AS b3 ON b4.MasoGD = b3.MasoGD
                    WHERE CAST(b9.NgayHT AS DATE) = CAST(@Value AS DATE)";
                    break;

                case "SOHC":
                    title = $"Passenger Details - SoHC: {value}";
                    query = @"SELECT ROW_NUMBER() OVER (ORDER BY bc09.NgayHT DESC) AS [STT], 
                      bc09.SoHD, bc09.KyhieuHD, bc09.NgayHD, bc09.NgayHT,
                      bc09.HoTenHK, bc09.SoHC, bc09.NgayHC, bc09.Quoctich, 
                      bc09.TrigiaHHchuaVAT, bc09.SotienVATDH, bc09.SotienDVNHH,
                      bc09.MasoDN, bc09.TenDNBH
                      FROM BC09 as bc09
                      WHERE bc09.SoHC = @Value
                      ORDER BY bc09.NgayHT DESC";
                    break;

                case "SOHD":
                    title = $"Invoice Details - SoHD: {value}";
                    query = @"WITH ParsedBC04 AS (
                        SELECT MasoGD, KyhieuSoNgay,
                            SUBSTRING(KyhieuSoNgay, 
                                CHARINDEX('/', KyhieuSoNgay) + 1, 
                                CHARINDEX('/', KyhieuSoNgay, CHARINDEX('/', KyhieuSoNgay) + 1) - CHARINDEX('/', KyhieuSoNgay) - 1
                            ) AS ExtractedSoHD
                        FROM BC04
                        WHERE KyhieuSoNgay LIKE '%/%/%'
                    )
                    SELECT 
                        ROW_NUMBER() OVER (ORDER BY b9.NgayHT DESC) AS [STT], 
                        b9.SoHD, b9.NgayHD, b9.KyhieuHD, b9.NgayHT,
                        b3.ThoigianGD, b9.HoTenHK, b9.SoHC, b9.Quoctich, 
                        b9.TrigiaHHchuaVAT, b9.SotienVATDH, b9.SotienDVNHH,
                        b9.MasoDN, b9.TenDNBH
                    FROM BC09 AS b9
                    INNER JOIN ParsedBC04 AS b4 ON b9.SoHD = b4.ExtractedSoHD
                    INNER JOIN BC03 AS b3 ON b4.MasoGD = b3.MasoGD
                    WHERE b9.SoHD = @Value";
                    break;

                default:
                    break;
            }

            return new QueryInfo { Query = query, Title = title };
        }

        /// <summary>
        /// Executes the SQL query and returns the resulting DataTable.
        /// </summary>
        private System.Data.DataTable ExecuteDetailsQuery(string query, string parameterValue, IWin32Window owner)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            // The 'using' block will ensure the connection is disposed of and closed
            using (var conn = _utility.OpenDB() as SqlConnection) // Explicit cast to SqlConnection needed if OpenDB returns IDbConnection
            {
                if (conn == null)
                {
                    MessageBox.Show(owner, DatabaseError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }

                // Ensure connection is open
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    // Use SqlParameter for better type handling, but AddWithValue is acceptable here
                    cmd.Parameters.AddWithValue("@Value", parameterValue);

                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(dt);
                    }
                }
            } // Connection is closed and disposed here

            return dt;
        }


        // --- UI Configuration Methods ---
        private static void ConfigureGridViewHeaders(DataGridView dgv)
        {
            var headerFont = new System.Drawing.Font("Microsoft Sans Serif", 10f, FontStyle.Bold);

            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                Font = headerFont,
                ForeColor = Color.DeepSkyBlue,
                BackColor = Color.WhiteSmoke,
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Padding = new Padding(3),
            };

            dgv.ColumnHeadersHeight = 35;
        }

        private static void AutoResizePopupColumns(DataGridView dgv)
        {
            var columnResizeRules = new Dictionary<string, DataGridViewAutoSizeColumnMode>(StringComparer.OrdinalIgnoreCase)
            {
                { "STT", DataGridViewAutoSizeColumnMode.DisplayedCells },
                { "HoTenHK", DataGridViewAutoSizeColumnMode.DisplayedCells },
                { "SoHC", DataGridViewAutoSizeColumnMode.DisplayedCells },
                { "Quoctich", DataGridViewAutoSizeColumnMode.DisplayedCells },
                { "ThoigianGD", DataGridViewAutoSizeColumnMode.DisplayedCells },
                { "TenDNBH", DataGridViewAutoSizeColumnMode.DisplayedCells },
            };

            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (columnResizeRules.TryGetValue(column.Name, out var resizeMode))
                {
                    column.AutoSizeMode = resizeMode;
                }
                else
                {
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    column.Width = 120;
                }
            }

            //// Ensure the last column fills the remaining space
            //if (dgv.Columns.Count > 0)
            //{
            //    dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //}
        }

        private void ApplyNumberFormatting(DataGridView dgv)
        {
            if (dgv.Rows.Count == 0 || dgv.Columns.Count == 0) return;

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
                if (dgv.Columns.Contains(numericCol.Key))
                {
                    var col = dgv.Columns[numericCol.Key];
                    col.DefaultCellStyle = numberStyle;
                    col.HeaderText = numericCol.Value;
                    col.Width = 150; // Set consistent width for numeric columns
                }
            }
        }
        private static void ApplyCustomColumnHeaders(DataGridView dgv)
        {
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
                    { "ThoigianGD", "Transaction Time" } // Specific to the SoHD query                   
                };

            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (columnHeaderMap.TryGetValue(column.Name, out var displayHeader))
                {
                    column.HeaderText = displayHeader;
                }
            }
        }
        private void ShowPopupForm(string title, System.Data.DataTable data, IWin32Window owner = null)
        {
            var popupForm = new Form
            {
                Text = title,
                StartPosition = FormStartPosition.CenterParent,
                Width = 1600,
                Height = 600,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true
            };

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                DataSource = data, // 1. Data Source is set first
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9f, FontStyle.Regular),
                BackgroundColor = SystemColors.Window,
                ForeColor = SystemColors.ControlText,
                ColumnHeadersHeight = 35,
            };

            // 2. Apply general headers/styles *before* custom formatting
            ConfigureGridViewHeaders(dgv);

            // 3. Attach an event handler that runs *after* the DGV has finished its layout cycle.
            // This is the most reliable way to ensure columns are ready for formatting.
            dgv.DataBindingComplete += (sender, e) =>
            {
                // These calls are now guaranteed to run when the column collection is stable
                AutoResizePopupColumns(dgv);
                ApplyCustomColumnHeaders(dgv);
                ApplyNumberFormatting(dgv);
            };

            panel.Controls.Add(dgv);
            popupForm.Controls.Add(panel);
            popupForm.ShowDialog(owner);
        }
    }
}