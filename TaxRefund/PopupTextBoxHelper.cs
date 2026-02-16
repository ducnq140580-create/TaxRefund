using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace TaxRefund
{ 
    public class PopupTextBoxHelper
    {
        // Dictionary to store query configurations for each TextBox
        private static readonly Dictionary<string, (string query, string title)> TextBoxQueryMap =
            new Dictionary<string, (string, string)>
        {
        { "txtSoHC", (@"SELECT ROW_NUMBER() OVER (ORDER BY NgayHT DESC) as STT, KyhieuHD, SoHD, NgayHD, TrigiaHHchuaVAT, 
                      SotienVATDH, SotienDVNHH, NgayHT, MasoDN, TenDNBH FROM BC09 WHERE SoHC = @SearchValue", "Passport Details Lookup") },
        { "txtSoHD", (@"SELECT ROW_NUMBER() OVER (ORDER BY NgayHT DESC) as STT, KyhieuHD, SoHD, NgayHD, TrigiaHHchuaVAT, 
                      SotienVATDH, SotienDVNHH, NgayHT, MasoDN, TenDNBH FROM BC09 WHERE SoHD = @SearchValue ORDER BY NgayHT DESC", "Invoice Details Lookup") },
        };

        public static void EnableTextBoxPopup(TextBox textBox)
        {
            if (!TextBoxQueryMap.ContainsKey(textBox.Name))
            {
                throw new ArgumentException($"No query configuration found for TextBox {textBox.Name}");
            }

            // Add click event handler
            textBox.Click += (sender, e) =>
            {
                ShowQueryPopup(textBox);
            };

            // Optional: Add visual cue
            textBox.Cursor = Cursors.Hand;
            textBox.BackColor = Color.AliceBlue;
            textBox.ReadOnly = true; // Prevent direct editing
        }

        private static void AutoResizePopupColumns(DataGridView dgv)
        {
            if (dgv.Columns.Count == 0) return;

            // Dictionary of column names and their resize modes
            var columnResizeRules = new Dictionary<string, DataGridViewAutoSizeColumnMode>
        {
            { "TenDNBH", DataGridViewAutoSizeColumnMode.DisplayedCells },
        };

            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (columnResizeRules.TryGetValue(column.Name, out var resizeMode))
                {
                    column.AutoSizeMode = resizeMode;

                    // For fill columns, set minimum width
                    if (resizeMode == DataGridViewAutoSizeColumnMode.Fill)
                    {
                        column.MinimumWidth = 100; // Prevent becoming too narrow
                    }
                }
                else
                {
                    // Default behavior for unspecified columns
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    column.Width = 120; // Fixed width
                }
            }

            // Optional: Set padding for the last column
            dgv.Columns[dgv.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private static void ConfigureGridViewHeaders(DataGridView dgv)
        {
            // Font settings
            var headerFont = new Font("Microsoft Sans Serif", 11f, FontStyle.Bold);

            // Main header style
            dgv.EnableHeadersVisualStyles = false; // Disable system styles
            dgv.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                Font = headerFont,
                ForeColor = Color.DeepSkyBlue,
                BackColor = Color.WhiteSmoke,
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Padding = new Padding(3)
            };
        }

        private static void ApplyNumberFormatting(DataGridView dgv)
        {
            // Create number formatting style
            DataGridViewCellStyle numberStyle = new DataGridViewCellStyle
            {
                Format = "N0", // Format with thousand separators, no decimals
                Alignment = DataGridViewContentAlignment.MiddleRight,
                BackColor = Color.White, // Ensure background is consistent
                ForeColor = Color.Blue, // Match your existing color scheme
                Font = new Font("Microsoft Sans Serif", 11f, FontStyle.Regular)
            };

            // List of columns that should have number formatting
            string[] numericColumns = { "TrigiaHHchuaVAT", "SotienVATDH", "SotienDVNHH" };

            foreach (string columnName in numericColumns)
            {
                if (dgv.Columns.Contains(columnName))
                {
                    dgv.Columns[columnName].DefaultCellStyle = numberStyle;

                    // Also ensure the column header is properly aligned
                    dgv.Columns[columnName].HeaderCell.Style.Alignment =
                        DataGridViewContentAlignment.MiddleCenter;
                }
            }
        }
        private static void ShowQueryPopup(TextBox textBox)
        {
            var config = TextBoxQueryMap[textBox.Name];

            using (var popup = new Form())
            {
                popup.Text = config.title;
                popup.StartPosition = FormStartPosition.CenterParent;
                popup.Width = 1200;
                popup.Height = 600;
                popup.FormBorderStyle = FormBorderStyle.FixedDialog;

                var scrollPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    Padding = new Padding(0, 0, 20, 0)
                };

                // Add DataGridView for results
                var dgv = new DataGridView
                {
                    ScrollBars = ScrollBars.Both,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    AllowUserToAddRows = false,
                    AllowUserToDeleteRows = false,
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    Margin = new Padding(5, 20, 5, 5),
                    Font = new Font("Microsoft Sans Serif", 11f, FontStyle.Regular),
                    BackgroundColor = Color.AntiqueWhite,
                    ForeColor = Color.Blue,
                    ColumnHeadersHeight = 30,
                };

                scrollPanel.Controls.Add(dgv);
                popup.Controls.Add(scrollPanel);

                // Handle the DataBindingComplete event to apply formatting AFTER data is loaded
                dgv.DataBindingComplete += (s, e) =>
                {
                    ConfigureGridViewHeaders(dgv);
                    AutoResizePopupColumns(dgv);

                    // Apply number formatting here (only if columns exist)
                    if (dgv.Columns.Contains("TrigiaHHchuaVAT"))
                    {
                        //DataGridViewCellStyle style = new DataGridViewCellStyle();
                        //style.Format = "N0";
                        //style.Alignment = DataGridViewContentAlignment.MiddleRight;

                        //dgv.Columns["TrigiaHHchuaVAT"].DefaultCellStyle = style;
                        //dgv.Columns["SotienVATDH"].DefaultCellStyle = style;
                        //dgv.Columns["SotienDVNHH"].DefaultCellStyle = style;

                        ApplyNumberFormatting(dgv);
                    }
                };

                // Handle row selection
                dgv.CellDoubleClick += (s, e) =>
                {
                    if (e.RowIndex >= 0 && dgv.Rows[e.RowIndex].Cells[0].Value != null)
                    {
                        textBox.Text = dgv.Rows[e.RowIndex].Cells[0].Value.ToString();
                        popup.DialogResult = DialogResult.OK;
                    }
                };

                // Load data - this will trigger the DataBindingComplete event
                LoadData(dgv, config.query, textBox.Text);

                // Show the dialog AFTER setting up all event handlers
                if (popup.ShowDialog() == DialogResult.OK)
                {
                    textBox.Focus();
                }
            }
        }

        //private static void ShowQueryPopup(TextBox textBox)
        //{
        //    var config = TextBoxQueryMap[textBox.Name];

        //    using (var popup = new Form())
        //    {
        //        popup.Text = config.title;
        //        popup.StartPosition = FormStartPosition.CenterParent;
        //        popup.Width = 1200;
        //        popup.Height = 600;
        //        popup.FormBorderStyle = FormBorderStyle.FixedDialog;

        //        var scrollPanel = new Panel
        //        {
        //            Dock = DockStyle.Fill,
        //            AutoScroll = true,
        //            Padding = new Padding(0, 0, 20, 0)
        //        };

        //        // Add DataGridView for results
        //        var dgv = new DataGridView
        //        {
        //            ScrollBars = ScrollBars.Both,
        //            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        //            AllowUserToAddRows = false,
        //            AllowUserToDeleteRows = false,
        //            Dock = DockStyle.Fill,
        //            ReadOnly = true,
        //            Margin = new Padding(5, 20, 5, 5),
        //            Font = new Font("Microsoft Sans Serif", 11f, FontStyle.Regular),
        //            BackgroundColor = Color.AntiqueWhite,
        //            ForeColor = Color.Blue,
        //            ColumnHeadersHeight = 30,
        //        };

        //        scrollPanel.Controls.Add(dgv);
        //        popup.Controls.Add(scrollPanel);

        //        LoadData(dgv, config.query, textBox.Text);

        //        //// Handle the DataBindingComplete event to apply formatting AFTER data is loaded
        //        //dgv.DataBindingComplete += (s, e) =>
        //        //{
        //            ConfigureGridViewHeaders(dgv);
        //            AutoResizePopupColumns(dgv);
        //            //ApplyNumberFormatting(dgv); // Apply number formatting here

        //            // ADD THE NUMBER FORMATTING CODE HERE - RIGHT AFTER AutoResizePopupColumns
        //            DataGridViewCellStyle style = new DataGridViewCellStyle();
        //            style.Format = "N0";
        //            dgv.Columns["TrigiaHHchuaVAT"].DefaultCellStyle = style;
        //            dgv.Columns["SotienVATDH"].DefaultCellStyle = style;
        //            dgv.Columns["SotienDVNHH"].DefaultCellStyle = style;

        //            dgv.Columns["TrigiaHHchuaVAT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //            dgv.Columns["SotienVATDH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //            dgv.Columns["SotienDVNHH"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        //            // Handle row selection (this should come AFTER the formatting)
        //            dgv.CellDoubleClick += (s, e) =>
        //            {
        //                if (dgv.CurrentRow != null)
        //                {
        //                    textBox.Text = dgv.CurrentRow.Cells[0].Value?.ToString();
        //                    popup.DialogResult = DialogResult.OK;
        //                }
        //            };

        //            if (popup.ShowDialog() == DialogResult.OK)
        //            {
        //                textBox.Focus();
        //            }
        //        //};

        //        //// Handle row selection
        //        //dgv.CellDoubleClick += (s, e) =>
        //        //{
        //        //    if (e.RowIndex >= 0 && dgv.Rows[e.RowIndex].Cells[0].Value != null)
        //        //    {
        //        //        textBox.Text = dgv.Rows[e.RowIndex].Cells[0].Value.ToString();
        //        //        popup.DialogResult = DialogResult.OK;
        //        //    }
        //        //};

        //        //if (popup.ShowDialog() == DialogResult.OK)
        //        //{
        //        //    textBox.Focus();
        //        //}
        //    }
        //}

        private static void LoadData(DataGridView dgv, string query, string searchValue)
        {
            try
            {
                Utility ut = new Utility();
                var conn = ut.OpenDB();
                using (var cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@SearchValue", searchValue);
                    var adapter = new SqlDataAdapter(cmd);
                    var dt = new System.Data.DataTable();
                    adapter.Fill(dt);

                    dgv.DataSource = dt;

                    if (dt.Rows.Count == 0)
                    {
                        MessageBox.Show("No records found.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}");
            }
        }
    }
}
