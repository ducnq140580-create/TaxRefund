using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace TaxRefund
{
    /*
 NOTE: This is an improved version compared to what is shown in the video (https://www.youtube.com/watch?v=kpUEv3euhe4) 
  it has an additional feature of showing the
 DataGrid's cells as disabled when the loading cursor is showing.
 */
    public class ProgressDataGridView : DataGridView
    {   
        public event RunWorkerCompletedEventHandler SearchCompleted;

        private const int AnimationAngleIncrement = 45;
        private const int AnimationDelayMs = 75;
        private const int CursorSize = 70;
        private const int BrushWidth = 6;

        private readonly BackgroundWorker _animationWorker = new BackgroundWorker();
        private readonly BackgroundWorker _dataWorker = new BackgroundWorker();

        private int _currentAngle;
        private bool _showLoadingCursor;
        private Bitmap _gridCellsImageCopy;
        private ControlsRectangle _gridRectangle;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out ControlsRectangle lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct ControlsRectangle
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public int Width => Right - Left;
            public int Height => Bottom - Top;
        }

        public ProgressDataGridView()
        {
            InitializeWorkers();
        }

        private void InitializeWorkers()
        {
            _animationWorker.DoWork += AnimationWorker_DoWork;
            _dataWorker.DoWork += DataWorker_DoWork;
            _dataWorker.RunWorkerCompleted += DataWorker_RunWorkerCompleted;
        }

        public void SimilarSearch(string commandText, Dictionary<string, object> parameters, bool autoGenerateColumns)
        {
            ExecuteSearch(commandText, parameters, autoGenerateColumns);          
        }

        public void NormalSearch(string commandText, Dictionary<string, object> parameters, bool autoGenerateColumns)
        {
            ExecuteSearch(commandText, parameters, autoGenerateColumns);
        }

        public void HighValueSearch(string commandText, Dictionary<string, object> parameters, bool autoGenerateColumns)
        {
            ExecuteSearch(commandText, parameters, autoGenerateColumns);
        }

        public void RefundManyTimesSearch(string commandText, Dictionary<string, object> parameters, bool autoGenerateColumns)
        {
            ExecuteSearch(commandText, parameters, autoGenerateColumns);
        }

        // Added a public DuplicateSearch wrapper to expose the existing ExecuteSearch pipeline
        public void DuplicateSearch(string commandText, Dictionary<string, object> parameters, bool autoGenerateColumns)
        {
            ExecuteSearch(commandText, parameters, autoGenerateColumns);
        }

        private void ExecuteSearch(string commandText, Dictionary<string, object> parameters, bool autoGenerateColumns)
        {
            if (_animationWorker.IsBusy || _dataWorker.IsBusy)
            {
                return;
            }

            PrepareForLoading();
            StartLoadingAnimation();
            StartDataLoading(commandText, autoGenerateColumns, parameters);           
        }

        private void PrepareForLoading()
        {
            _showLoadingCursor = true;
            GetGridBodyAndSaveToImage();
        }

        private void StartLoadingAnimation()
        {
            _animationWorker.RunWorkerAsync();
        }

        private void StartDataLoading(string commandText, bool autoGenerateColumns, Dictionary<string, object> parameters)
        {
            object[] arguments = { commandText, autoGenerateColumns, parameters };
            _dataWorker.RunWorkerAsync(arguments);
        }       
        private void DataWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            CompleteLoading();
            // Raise the completion event
            SearchCompleted?.Invoke(this, e);
        }

        private void CompleteLoading()
        {
            _showLoadingCursor = false;
            Invalidate();
            _gridCellsImageCopy = null;
        }

        private void DataWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var arguments = (object[])e.Argument;
            var commandText = (string)arguments[0];
            var autoGenerateColumns = (bool)arguments[1];
            var parameters = arguments.Length > 2 ? (Dictionary<string, object>)arguments[2] : null;

            var utility = new Utility();
            using (var conn = utility.OpenDB())
            {
                try
                {
                    if (conn.State == ConnectionState.Closed)
                    {
                        conn.Open();
                        LoadDataToGrid(conn, commandText, autoGenerateColumns, parameters);
                    }
                }
                finally
                {
                    if (conn.State != ConnectionState.Closed)
                    {
                        conn.Close();
                    }
                }
            }
        }

        private void LoadDataToGrid(SqlConnection conn, string commandText, bool autoGenerateColumns, Dictionary<string, object> parameters)
        {
            using (var cmd = new SqlCommand(commandText, conn))
            {
                AddParametersToCommand(cmd, parameters);

                using (var dataAdapter = new SqlDataAdapter(cmd))
                using (var dataSet = new DataSet())
                {
                    dataAdapter.Fill(dataSet);
                    var dt = dataSet.Tables[0];

                    Invoke(new Action(() =>
                    {
                        DataSource = dt;
                        AutoGenerateColumns = autoGenerateColumns;
                    }));
                }
            }
        }

        private static void AddParametersToCommand(SqlCommand cmd, Dictionary<string, object> parameters)
        {
            if (parameters == null) return;

            foreach (var param in parameters)
            {
                cmd.Parameters.AddWithValue(param.Key, param.Value ?? DBNull.Value);
            }
        }

        private void AnimationWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            while (_showLoadingCursor)
            {
                UpdateAnimationAngle();
                PaintLoadingCursor();
                Thread.Sleep(AnimationDelayMs);
            }
        }

        private void UpdateAnimationAngle()
        {
            _currentAngle += AnimationAngleIncrement;
            if (_currentAngle > 360)
            {
                _currentAngle = 0;
            }
        }

        private void PaintLoadingCursor()
        {
            using (var graphics = CreateGraphics())
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;

                var cursorPosition = CalculateCursorPosition();
                DrawGridBackground(graphics);
                DrawBaseArc(graphics, cursorPosition);
                DrawAnimatedArc(graphics, cursorPosition);
            }
        }

        private Rectangle CalculateCursorPosition()
        {
            int cursorX = (Width / 2) - (CursorSize / 2);
            int cursorY = (Height / 2) - (CursorSize / 2);
            return new Rectangle(cursorX, cursorY, CursorSize, CursorSize);
        }

        private void DrawGridBackground(Graphics graphics)
        {
            int x = RowHeadersVisible ? RowHeadersWidth : 0;
            int y = ColumnHeadersVisible ? ColumnHeadersHeight : 0;
            graphics.DrawImage(_gridCellsImageCopy, x, y);
        }

        private void DrawBaseArc(Graphics graphics, Rectangle position)
        {
            using (var brush = new LinearGradientBrush(ClientRectangle,
                   Color.FromArgb(93, 93, 93), Color.FromArgb(50, 205, 50),
                   LinearGradientMode.Vertical))
            using (var pen = new Pen(brush, BrushWidth))
            {
                pen.DashStyle = DashStyle.Dot;
                graphics.DrawArc(pen, position, 0, 360);
            }
        }

        private void DrawAnimatedArc(Graphics graphics, Rectangle position)
        {
            using (var brush = new SolidBrush(Color.White))
            using (var pen = new Pen(brush, BrushWidth))
            {
                pen.DashStyle = DashStyle.Dot;
                graphics.DrawArc(pen, position, _currentAngle, 90);
            }
        }

        private void GetGridBodyAndSaveToImage()
        {
            GetWindowRect(Handle, out _gridRectangle);

            var dimensions = CalculateGridBodyDimensions();
            _gridCellsImageCopy = new Bitmap(dimensions.Width, dimensions.Height);

            using (var bitmapGraphics = Graphics.FromImage(_gridCellsImageCopy))
            {
                var copyPosition = new Point(
                    _gridRectangle.Left + (RowHeadersVisible ? RowHeadersWidth : 0),
                    _gridRectangle.Top + (ColumnHeadersVisible ? ColumnHeadersHeight : 0)
                );

                bitmapGraphics.CopyFromScreen(
                    copyPosition.X, copyPosition.Y,
                    0, 0,
                    new Size(_gridRectangle.Width, _gridRectangle.Height),
                    CopyPixelOperation.SourceCopy);
            }

            if (Rows.Count > 0)
            {
                _gridCellsImageCopy = (Bitmap)ToolStripRenderer.CreateDisabledImage(_gridCellsImageCopy);
            }
        }

        private Size CalculateGridBodyDimensions()
        {
            int rowHeadsWidth = RowHeadersVisible ? RowHeadersWidth : 0;
            int columnHeadsHeight = ColumnHeadersVisible ? ColumnHeadersHeight : 0;
            return new Size(Width - rowHeadsWidth - 1, Height - columnHeadsHeight - 1);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            if (!_showLoadingCursor)
            {
                base.OnPaint(e);
            }
        }        
    }
}
