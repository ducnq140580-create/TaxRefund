using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TaxRefund
{ // UI Helper
    public static class DataGridViewHelper
    {
        private static readonly Dictionary<string, string> ColumnHeaderMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "STT", "No." },
                { "Motahanghoa", "Description Of Gooods" },
                { "TongTrigiaHHchuaVAT (VND)", "Total Goods Value (ex VAT)" },
                { "TongLKTrigiaHHchuaVAT (VND)", "Total Accumulated Goods Value (ex VAT)" },
                { "TongSotienVATDH (VND)", "Total VAT Refundable Amount" },
                { "TongLKSotienVATDH (VND)", "Total Accumulated VAT Refund" },
                { "TongSotienDVNHH (VND)", "Total Service Fee" },
                { "TongLKSotienDVNHH (VND)", "Total Accumulated Service Fee" },
                { "TongLuotHK (HK)", "Total Passenger Turns" },
                { "TongLKLuotHK (HK)", "Total Accumulated Passenger Turns" },
                { "SosanhTongTrigiaHHchuaVAT (%)", "Goods Value Comparison (%)" },
                { "SosanhTongSotienVATDH (%)", "VAT Refund Comparison (%)" },
                { "SosanhTongLuotHK (%)", "Passenger Turn Comparison (%)" }
            };

        public static void ConfigureReportGridView(DataGridView dataGridView)
        {
            // Apply header text
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                if (ColumnHeaderMap.TryGetValue(column.Name, out var displayHeader))
                {
                    column.HeaderText = displayHeader;
                }
            }

            // Apply number formatting
            var numberStyle = new DataGridViewCellStyle { Format = "N0" };

            string[] numberColumns = {
                "TongTrigiaHHchuaVAT (VND)",
                "TongLKTrigiaHHchuaVAT (VND)",
                "TongSotienVATDH (VND)",
                "TongLKSotienVATDH (VND)",
                "TongSotienDVNHH (VND)",
                "TongLKSotienDVNHH (VND)",
                "TongLuotHK (HK)",
                "TongLKLuotHK (HK)"
            };

            foreach (var colName in numberColumns)
            {
                if (dataGridView.Columns.Contains(colName))
                {
                    dataGridView.Columns[colName].DefaultCellStyle = numberStyle;
                    dataGridView.Columns[colName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }

            // Center alignment for specific columns
            dataGridView.Columns["STT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView.Columns["Motahanghoa"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView.Columns["SosanhTongTrigiaHHchuaVAT (%)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView.Columns["SosanhTongSotienVATDH (%)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView.Columns["SosanhTongLuotHK (%)"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Styling
            var headerFont = new System.Drawing.Font("Microsoft Sans Serif", 11f, FontStyle.Bold);
            dataGridView.EnableHeadersVisualStyles = false;
            dataGridView.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                Font = headerFont,
                ForeColor = Color.DeepSkyBlue,
                BackColor = Color.AntiqueWhite,
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Padding = new Padding(3)
            };

            dataGridView.DefaultCellStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 11f, FontStyle.Regular);
            dataGridView.DefaultCellStyle.BackColor = Color.FloralWhite;
            dataGridView.DefaultCellStyle.ForeColor = Color.DarkBlue;
            dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView.GridColor = Color.Silver;
            dataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridView.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView.MultiSelect = true;
        }
    }
}
