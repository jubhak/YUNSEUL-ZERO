using System.Drawing.Drawing2D;

namespace DataParser
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        private TextBox txtInput = null!;
        private Button btnProcess = null!;
        private Button btnSelectExcel = null!;
        private Button btnInsert = null!;
        private Label lblInput = null!;
        private Label lblResult = null!;
        private Label lblExcelPath = null!;
        private Label lblStatus = null!;
        private Panel panelTop = null!;
        private Panel panelMid = null!;
        private Panel panelBottom = null!;
        private Panel panelTable = null!;
        private Panel panelHeader = null!;
        private Label hdrTitle = null!;
        private Label hdrValue = null!;
        private Label hdrUnique = null!;
        private Label hdrCol = null!;
        private ComboBox cboSheet = null!;
        private Label lblSheetSelect = null!;

        // 결과 테이블 행 데이터
        private TextBox[] _txtKeywords = null!;
        private TextBox[] _txtValues = null!;
        private RadioButton[] _rdoUnique = null!;
        private TextBox[] _txtColNums = null!;
        private Panel[] _hLines = null!;
        private Panel _vLine1 = null!;
        private Panel _vLine2 = null!;
        private Panel _vLine3 = null!;

        // 다건 추출 관련
        private static readonly int MAX_RECORDS = 20;
        private List<string[]> _allRecords = new List<string[]>();
        private int _currentRecordIndex = 0;
        private Label[] _recordBtns = null!;

        // UPDATE/INSERT 상태 표시 라벨
        private Label lblUpsertStatus = null!;

        // 구글 시트 관련
        private RadioButton rdoLocal = null!;
        private RadioButton rdoGoogle = null!;
        private Panel panelLocalSettings = null!;
        private Panel panelGoogleSettings = null!;
        private Button btnSelectJson = null!;
        private Label lblJsonPath = null!;
        private TextBox txtSheetUrl = null!;
        private Button btnGoogleConnect = null!;
        private ComboBox cboGoogleSheet = null!;
        private Label lblGoogleStatus = null!;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        // 라운드 버튼
        private class RoundButton : Button
        {
            public int Radius { get; set; } = 2;
            public RoundButton()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                e.Graphics.Clear(Parent?.BackColor ?? Color.Black);
                var rect = new Rectangle(1, 1, Width - 3, Height - 3);
                using var path = MakeRound(rect, Radius);
                using var bg = new SolidBrush(BackColor);
                e.Graphics.FillPath(bg, path);
                TextRenderer.DrawText(e.Graphics, Text, Font, new Rectangle(0, 0, Width, Height), ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            private static GraphicsPath MakeRound(Rectangle r, int rad)
            {
                var p = new GraphicsPath();
                int d = rad * 2;
                p.AddArc(r.X, r.Y, d, d, 180, 90);
                p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                p.CloseFigure();
                return p;
            }
        }

        [System.Runtime.InteropServices.DllImport("Gdi32.dll")]
        private static extern IntPtr CreateRoundRectRgn(int x1, int y1, int x2, int y2, int cx, int cy);

        [System.Runtime.InteropServices.DllImport("User32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [System.Runtime.InteropServices.DllImport("User32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, ref RECT lParam);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private static void SetWindowMargins(IntPtr handle, int left, int right)
        {
            const int EM_SETMARGINS = 0xD3;
            SendMessage(handle, EM_SETMARGINS, 0x1 | 0x2, (right << 16) | left);
        }

        private static void SetTextBoxRect(TextBox tb, int left, int top, int right, int bottom)
        {
            const int EM_SETRECT = 0xB3;
            var rect = new RECT { Left = left, Top = top, Right = right, Bottom = bottom };
            SendMessage(tb.Handle, EM_SETRECT, 0, ref rect);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            this.SuspendLayout();

            // === 색상 ===
            Color BG     = Color.FromArgb(13, 17, 23);
            Color CARD   = Color.FromArgb(22, 27, 34);
            Color BORDER = Color.FromArgb(48, 54, 61);
            Color FG     = Color.FromArgb(230, 237, 243);
            Color FG2    = Color.FromArgb(139, 148, 158);
            Color BLUE   = Color.FromArgb(88, 166, 255);
            Color GREEN  = Color.FromArgb(63, 185, 80);

            // ============================================================
            // 상단: 타이틀 + 입력 영역
            // ============================================================
            panelTop = new Panel { Dock = DockStyle.Top, Height = 296, BackColor = BG, Padding = new Padding(10, 0, 10, 0) };

            // 타이틀 바 (panelTop 내부 맨 위)
            var panelTitleBar = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = BG };
            var lblTitle = new Label
            {
                Text = "  [▓]  YUNSEUL-ZERO   VER 1.9",
                ForeColor = BLUE,
                BackColor = Color.Transparent,
                Font = new Font("Consolas", 9F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(0, 7)
            };
            panelTitleBar.Controls.Add(lblTitle);

            // 타이틀 바 드래그
            panelTitleBar.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { _dragging = true; _dragStart = e.Location; } };
            panelTitleBar.MouseMove += (s, e) => { if (_dragging) { var p = PointToScreen(e.Location); SetDesktopLocation(p.X - _dragStart.X, p.Y - _dragStart.Y); } };
            panelTitleBar.MouseUp += (s, e) => _dragging = false;
            lblTitle.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { _dragging = true; _dragStart = e.Location; } };
            lblTitle.MouseMove += (s, e) => { if (_dragging) { var p = this.PointToScreen(e.Location); SetDesktopLocation(p.X - _dragStart.X, p.Y - _dragStart.Y); } };
            lblTitle.MouseUp += (s, e) => _dragging = false;

            // DATA INPUT 라벨 위 10px 여백용 패널
            var panelInputGap = new Panel { Dock = DockStyle.Top, Height = 10, BackColor = BG };

            lblInput = new Label
            {
                Text = "  ▌ DATA INPUT",
                ForeColor = Color.FromArgb(255, 140, 50),
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7.5F, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 22
            };

            txtInput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F),
                BackColor = CARD,
                ForeColor = FG,
                BorderStyle = BorderStyle.FixedSingle,
                AcceptsReturn = true,
                AcceptsTab = true,
                WordWrap = true
            };

            var panelBtn1 = new Panel { Dock = DockStyle.Bottom, Height = 38, BackColor = BG };
            btnProcess = new RoundButton
            {
                Text = "▶  데이터 추출",
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                BackColor = BLUE,
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false
            };
            btnProcess.FlatAppearance.BorderSize = 0;
            btnProcess.Click += BtnProcess_Click;
            panelBtn1.Controls.Add(btnProcess);
            panelBtn1.Resize += (s, e) =>
            {
                var sz = TextRenderer.MeasureText(btnProcess.Text, btnProcess.Font);
                btnProcess.Size = new Size(sz.Width + 20, 32);
                btnProcess.Location = new Point((panelBtn1.Width - btnProcess.Width) / 2, (panelBtn1.Height - btnProcess.Height) / 2);
            };

            panelTop.Controls.Add(txtInput);       // Fill
            panelTop.Controls.Add(panelBtn1);       // Bottom
            panelTop.Controls.Add(lblInput);        // Top (DATA INPUT 라벨)
            panelTop.Controls.Add(panelInputGap);   // Top (10px 여백)
            panelTop.Controls.Add(panelTitleBar);   // Top (타이틀 — 가장 마지막 Add → 가장 위)

            // ============================================================
            // 중간: 결과 테이블 (Label + TextBox 기반)
            // ============================================================
            panelMid = new Panel { Dock = DockStyle.Fill, BackColor = BG, Padding = new Padding(10, 4, 10, 4) };

            // PARSED RESULT 라벨 + 행 추가 버튼을 담는 패널
            var panelResultHeader = new Panel { Dock = DockStyle.Top, Height = 22, BackColor = BG };
            lblResult = new Label
            {
                Text = "  ▌ PARSED RESULT",
                ForeColor = GREEN,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7.5F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(0, 3)
            };
            var btnAddRow = new Label
            {
                Text = "+",
                ForeColor = GREEN,
                BackColor = Color.FromArgb(30, 36, 44),
                Font = new Font("Consolas", 8F, FontStyle.Bold),
                Size = new Size(18, 14),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand
            };
            btnAddRow.Click += BtnAddRow_Click;
            btnAddRow.MouseEnter += (s, e) => btnAddRow.ForeColor = Color.White;
            btnAddRow.MouseLeave += (s, e) => btnAddRow.ForeColor = GREEN;
            panelResultHeader.Controls.Add(lblResult);
            panelResultHeader.Controls.Add(btnAddRow);

            // UPDATE/INSERT 상태 라벨
            lblUpsertStatus = new Label
            {
                Text = "",
                ForeColor = Color.FromArgb(255, 200, 50),
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7F, FontStyle.Bold),
                AutoSize = true,
                Visible = false
            };
            panelResultHeader.Controls.Add(lblUpsertStatus);

            // 다건 추출 레코드 번호 버튼 (동그라미 1~20)
            _recordBtns = new Label[MAX_RECORDS];
            for (int ri = 0; ri < MAX_RECORDS; ri++)
            {
                int idx = ri;
                _recordBtns[ri] = new Label
                {
                    Text = (ri + 1).ToString(),
                    ForeColor = FG2,
                    BackColor = Color.FromArgb(30, 36, 44),
                    Font = new Font("Segoe UI", 6.5F, FontStyle.Bold),
                    Size = new Size(18, 18),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    Visible = false
                };
                _recordBtns[ri].Paint += (s, ev) =>
                {
                    var lbl = (Label)s!;
                    ev.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    ev.Graphics.Clear(lbl.Parent?.BackColor ?? BG);
                    using var bgBrush = new SolidBrush(lbl.BackColor);
                    ev.Graphics.FillEllipse(bgBrush, 0, 0, lbl.Width - 1, lbl.Height - 1);
                    TextRenderer.DrawText(ev.Graphics, lbl.Text, lbl.Font,
                        new Rectangle(0, 0, lbl.Width, lbl.Height), lbl.ForeColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                };
                _recordBtns[ri].Click += (s, ev) => SelectRecord(idx);
                _recordBtns[ri].MouseEnter += (s, ev) => { var l = (Label)s!; if (l.BackColor != Color.FromArgb(88, 166, 255)) l.ForeColor = Color.White; };
                _recordBtns[ri].MouseLeave += (s, ev) => { var l = (Label)s!; if (l.BackColor != Color.FromArgb(88, 166, 255)) l.ForeColor = Color.FromArgb(139, 148, 158); };
                panelResultHeader.Controls.Add(_recordBtns[ri]);
            }

            panelResultHeader.Resize += (s, e) =>
            {
                btnAddRow.Location = new Point(lblResult.Right + 6, 4);
                LayoutRecordButtons();
            };

            int hdrH = 26;
            int count = 13;
            var hdrBg = Color.FromArgb(1, 4, 9);

            // 헤더 패널 (고정, 스크롤 안 됨)
            panelHeader = new Panel { Dock = DockStyle.Top, Height = hdrH, BackColor = hdrBg };
            hdrTitle = new Label { Text = "KEYWORDS", BackColor = hdrBg, ForeColor = BLUE, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };
            hdrValue = new Label { Text = "VALUE", BackColor = hdrBg, ForeColor = BLUE, Font = new Font("Segoe UI", 9F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };
            hdrUnique = new Label { Text = "UNIQUE", BackColor = hdrBg, ForeColor = Color.FromArgb(255, 140, 50), Font = new Font("Segoe UI", 7.5F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(2, 0, 0, 0) };
            hdrCol = new Label { Text = "COL", BackColor = hdrBg, ForeColor = BLUE, Font = new Font("Segoe UI", 8F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(2, 0, 0, 0) };
            panelHeader.Controls.AddRange(new Control[] { hdrTitle, hdrValue, hdrUnique, hdrCol });

            // 데이터 패널 (스크롤 가능)
            panelTable = new Panel { Dock = DockStyle.Fill, BackColor = CARD, AutoScroll = true };

            _txtKeywords = new TextBox[count];
            _txtValues = new TextBox[count];
            _rdoUnique = new RadioButton[count];
            _txtColNums = new TextBox[count];

            _hLines = new Panel[count + 1]; // 가로선
            _vLine1 = new Panel { BackColor = BORDER, Width = 1 }; // title|value
            _vLine2 = new Panel { BackColor = BORDER, Width = 1 }; // value|unique
            _vLine3 = new Panel { BackColor = BORDER, Width = 1 }; // unique|col

            for (int i = 0; i <= count; i++)
                _hLines[i] = new Panel { BackColor = BORDER, Height = 1 };

            for (int i = 0; i < count; i++)
            {
                _txtKeywords[i] = new TextBox
                {
                    BackColor = CARD, ForeColor = FG2,
                    Font = new Font("Consolas", 8F),
                    BorderStyle = BorderStyle.None,
                    TextAlign = HorizontalAlignment.Left
                };
                _txtKeywords[i].Leave += (s, ev) => SaveSettings();
                _txtValues[i] = new TextBox
                {
                    BackColor = CARD, ForeColor = FG,
                    Font = new Font("Segoe UI", 9F),
                    BorderStyle = BorderStyle.None,
                    TextAlign = HorizontalAlignment.Left
                };
                _rdoUnique[i] = new RadioButton
                {
                    BackColor = CARD, ForeColor = FG,
                    AutoCheck = false,
                    Appearance = Appearance.Normal,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Text = "",
                    Cursor = Cursors.Hand
                };
                int capturedIdx = i;
                _rdoUnique[i].Click += (s, ev) => UniqueRadio_Click(capturedIdx);
                _txtColNums[i] = new TextBox
                {
                    BackColor = Color.FromArgb(30, 36, 44), ForeColor = FG,
                    Font = new Font("Consolas", 9F),
                    BorderStyle = BorderStyle.FixedSingle,
                    TextAlign = HorizontalAlignment.Right,
                    MaxLength = 3
                };
                _txtColNums[i].KeyPress += (s, ev) =>
                { if (!char.IsLetter(ev.KeyChar) && !char.IsControl(ev.KeyChar)) ev.Handled = true; };
                _txtColNums[i].Leave += (s, ev) =>
                { var tb = (TextBox)s!; tb.Text = tb.Text.Trim().ToUpper(); SaveSettings(); };

                panelTable.Controls.Add(_txtKeywords[i]);
                panelTable.Controls.Add(_txtValues[i]);
                panelTable.Controls.Add(_rdoUnique[i]);
                panelTable.Controls.Add(_txtColNums[i]);
            }

            for (int i = 0; i <= count; i++) panelTable.Controls.Add(_hLines[i]);
            panelTable.Controls.Add(_vLine1);
            panelTable.Controls.Add(_vLine2);
            panelTable.Controls.Add(_vLine3);

            // 데이터 레이아웃 (동적 행 수 지원)
            panelTable.Resize += (s, e) => LayoutTable();

            panelMid.Controls.Add(panelTable);   // Fill (스크롤)
            panelMid.Controls.Add(panelHeader);  // Top (고정)
            panelMid.Controls.Add(panelResultHeader);

            // ============================================================
            // 하단: 모드 선택 + 로컬/구글 설정 + 입력처리
            // ============================================================
            panelBottom = new Panel { Dock = DockStyle.Bottom, Height = 180, BackColor = BG, Padding = new Padding(10, 6, 10, 4) };

            // --- 모드 선택 라디오 버튼 ---
            rdoLocal = new RadioButton
            {
                Text = "로컬 엑셀",
                ForeColor = FG,
                BackColor = BG,
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                AutoSize = true,
                Checked = true,
                Cursor = Cursors.Hand
            };
            rdoGoogle = new RadioButton
            {
                Text = "구글 시트",
                ForeColor = FG,
                BackColor = BG,
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                AutoSize = true,
                Cursor = Cursors.Hand
            };
            rdoLocal.CheckedChanged += (s, ev) => ToggleMode();
            rdoGoogle.CheckedChanged += (s, ev) => ToggleMode();

            // --- 로컬 엑셀 설정 패널 ---
            panelLocalSettings = new Panel { BackColor = BG, Visible = true };

            btnSelectExcel = new RoundButton
            {
                Text = "📂  엑셀파일 선택",
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                BackColor = BLUE,
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false
            };
            btnSelectExcel.FlatAppearance.BorderSize = 0;
            btnSelectExcel.Click += BtnSelectExcel_Click;

            lblExcelPath = new Label
            {
                Text = "선택된 파일 없음",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7F),
                AutoSize = false,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                Cursor = Cursors.Default
            };
            lblExcelPath.Click += (s, ev) =>
            {
                if (!string.IsNullOrEmpty(_excelPath) && File.Exists(_excelPath))
                {
                    try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(_excelPath) { UseShellExecute = true }); }
                    catch { }
                }
            };
            lblExcelPath.MouseEnter += (s, ev) => { if (!string.IsNullOrEmpty(_excelPath)) lblExcelPath.Font = new Font(lblExcelPath.Font, FontStyle.Underline); };
            lblExcelPath.MouseLeave += (s, ev) => { lblExcelPath.Font = new Font(lblExcelPath.Font, FontStyle.Regular); };

            lblSheetSelect = new Label
            {
                Text = "엑셀시트 선택",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7.5F, FontStyle.Bold),
                AutoSize = true
            };

            cboSheet = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(30, 36, 44),
                ForeColor = FG,
                Font = new Font("Segoe UI", 8F),
                FlatStyle = FlatStyle.Flat,
                Width = 160
            };

            // 콤보박스를 감싸는 보더 패널
            Color cboBorder = Color.FromArgb(204, 204, 204);
            var panelCboWrap = new Panel
            {
                BackColor = cboBorder,
                Padding = new Padding(1)
            };
            cboSheet.Dock = DockStyle.Fill;
            panelCboWrap.Controls.Add(cboSheet);
            panelCboWrap.Paint += (s, ev) =>
            {
                using var pen = new Pen(cboBorder, 1);
                var rc = new Rectangle(0, 0, panelCboWrap.Width - 1, panelCboWrap.Height - 1);
                using var path = new GraphicsPath();
                int d = 2;
                path.AddArc(rc.X, rc.Y, d, d, 180, 90);
                path.AddArc(rc.Right - d, rc.Y, d, d, 270, 90);
                path.AddArc(rc.Right - d, rc.Bottom - d, d, d, 0, 90);
                path.AddArc(rc.X, rc.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                ev.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                ev.Graphics.DrawPath(pen, path);
            };

            panelLocalSettings.Controls.Add(btnSelectExcel);
            panelLocalSettings.Controls.Add(lblExcelPath);
            panelLocalSettings.Controls.Add(lblSheetSelect);
            panelLocalSettings.Controls.Add(panelCboWrap);

            panelLocalSettings.Resize += (s, e) =>
            {
                int cx = panelLocalSettings.Width;
                var szExcel = TextRenderer.MeasureText(btnSelectExcel.Text, btnSelectExcel.Font);
                btnSelectExcel.Size = new Size(szExcel.Width + 20, 30);
                btnSelectExcel.Location = new Point(0, 0);

                lblExcelPath.Location = new Point(btnSelectExcel.Width + 8, 4);
                lblExcelPath.Width = cx - btnSelectExcel.Width - 12;

                int sheetTop = btnSelectExcel.Bottom + 6;
                int cboH = cboSheet.Height + 2;
                panelCboWrap.Height = cboH;
                int lblH = lblSheetSelect.Height;
                int rowCenter = sheetTop + Math.Max(cboH, lblH) / 2;
                lblSheetSelect.Location = new Point(0, rowCenter - lblH / 2);
                panelCboWrap.Location = new Point(lblSheetSelect.Width + 6, rowCenter - cboH / 2);
                panelCboWrap.Width = cx - lblSheetSelect.Width - 10;
            };

            // --- 구글 시트 설정 패널 ---
            panelGoogleSettings = new Panel { BackColor = BG, Visible = false };

            btnSelectJson = new RoundButton
            {
                Text = "⚙  OAuth 설정",
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                BackColor = Color.FromArgb(255, 140, 50),
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false,
                Visible = false
            };
            btnSelectJson.FlatAppearance.BorderSize = 0;
            btnSelectJson.Click += BtnSelectJson_Click;

            lblJsonPath = new Label
            {
                Text = "",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7F),
                AutoSize = false,
                Height = 18,
                Visible = false,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var lblUrlLabel = new Label
            {
                Text = "시트 URL",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7.5F, FontStyle.Bold),
                AutoSize = true
            };

            txtSheetUrl = new TextBox
            {
                BackColor = Color.FromArgb(30, 36, 44),
                ForeColor = FG,
                Font = new Font("Segoe UI", 9F),
                BorderStyle = BorderStyle.None
            };
            txtSheetUrl.GotFocus += (s, ev) => { if (txtSheetUrl.Text == "구글 스프레드시트 URL 붙여넣기") { txtSheetUrl.Text = ""; txtSheetUrl.ForeColor = FG; } };
            txtSheetUrl.LostFocus += (s, ev) => { if (string.IsNullOrWhiteSpace(txtSheetUrl.Text)) { txtSheetUrl.Text = "구글 스프레드시트 URL 붙여넣기"; txtSheetUrl.ForeColor = FG2; } };
            txtSheetUrl.Text = "구글 스프레드시트 URL 붙여넣기";
            txtSheetUrl.ForeColor = FG2;

            // TextBox를 Panel로 감싸서 높이+패딩+보더 효과
            var panelUrlWrap = new Panel
            {
                BackColor = Color.FromArgb(30, 36, 44),
                BorderStyle = BorderStyle.FixedSingle
            };
            panelUrlWrap.Controls.Add(txtSheetUrl);
            // 수직 중앙 배치: Panel 리사이즈 시 TextBox 위치 계산
            panelUrlWrap.Resize += (s, ev) =>
            {
                int tbH = txtSheetUrl.Height;
                int pH = panelUrlWrap.ClientSize.Height;
                txtSheetUrl.Location = new Point(5, (pH - tbH) / 2);
                txtSheetUrl.Width = panelUrlWrap.ClientSize.Width - 10;
            };
            panelUrlWrap.Click += (s, ev) => txtSheetUrl.Focus();

            btnGoogleConnect = new RoundButton
            {
                Text = "로그인",
                Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                BackColor = GREEN,
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false,
                Size = new Size(60, 24)
            };
            btnGoogleConnect.FlatAppearance.BorderSize = 0;
            btnGoogleConnect.Click += BtnGoogleConnect_Click;

            var lblGSheetSelect = new Label
            {
                Text = "시트 선택",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7.5F, FontStyle.Bold),
                AutoSize = true
            };

            cboGoogleSheet = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(30, 36, 44),
                ForeColor = FG,
                Font = new Font("Segoe UI", 8F),
                FlatStyle = FlatStyle.Flat,
                Width = 160
            };

            lblGoogleStatus = new Label
            {
                Text = "",
                ForeColor = GREEN,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 7F, FontStyle.Bold),
                AutoSize = true
            };

            panelGoogleSettings.Controls.Add(btnSelectJson);
            panelGoogleSettings.Controls.Add(lblJsonPath);
            panelGoogleSettings.Controls.Add(lblUrlLabel);
            panelGoogleSettings.Controls.Add(panelUrlWrap);
            panelGoogleSettings.Controls.Add(btnGoogleConnect);
            panelGoogleSettings.Controls.Add(lblGSheetSelect);
            panelGoogleSettings.Controls.Add(cboGoogleSheet);
            panelGoogleSettings.Controls.Add(lblGoogleStatus);

            panelGoogleSettings.Resize += (s, e) =>
            {
                int cx = panelGoogleSettings.Width;

                int row1Top = 0;
                int rowH = 24; // panelUrlWrap 높이 기준
                lblUrlLabel.Location = new Point(0, row1Top + (rowH - lblUrlLabel.Height) / 2);
                int urlLeft = lblUrlLabel.Right + 6;
                int connectW = btnGoogleConnect.Width + 4;
                panelUrlWrap.Location = new Point(urlLeft, row1Top);
                panelUrlWrap.Width = cx - urlLeft - connectW - 4;
                panelUrlWrap.Height = rowH;
                btnGoogleConnect.Location = new Point(panelUrlWrap.Right + 4, row1Top + (rowH - btnGoogleConnect.Height) / 2);

                int row2Top = panelUrlWrap.Bottom + 6;
                lblGSheetSelect.Location = new Point(0, row2Top + 3);
                cboGoogleSheet.Location = new Point(lblGSheetSelect.Right + 6, row2Top);
                cboGoogleSheet.Width = Math.Min(200, cx - lblGSheetSelect.Right - 10);
                lblGoogleStatus.Location = new Point(cboGoogleSheet.Right + 10, row2Top + 3);
            };

            // --- 구분선 + 입력 버튼 ---
            var separator = new Panel { BackColor = Color.FromArgb(48, 54, 61), Height = 1 };

            btnInsert = new RoundButton
            {
                Text = "⬇  입력처리",
                Font = new Font("Segoe UI", 8.25F, FontStyle.Bold),
                BackColor = GREEN,
                ForeColor = Color.FromArgb(13, 17, 23),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                AutoSize = false
            };
            btnInsert.FlatAppearance.BorderSize = 0;
            btnInsert.Click += BtnInsert_Click;

            lblStatus = new Label
            {
                Text = "",
                ForeColor = FG2,
                BackColor = Color.Transparent,
                Font = new Font("Segoe UI", 8F),
                AutoSize = false,
                Height = 18,
                TextAlign = ContentAlignment.MiddleCenter
            };

            panelBottom.Controls.Add(rdoLocal);
            panelBottom.Controls.Add(rdoGoogle);
            panelBottom.Controls.Add(panelLocalSettings);
            panelBottom.Controls.Add(panelGoogleSettings);
            panelBottom.Controls.Add(separator);
            panelBottom.Controls.Add(lblStatus);
            panelBottom.Controls.Add(btnInsert);

            panelBottom.Resize += (s, e) =>
            {
                int cx = panelBottom.ClientSize.Width;
                int pad = panelBottom.Padding.Left;

                // 모드 선택 라디오
                rdoLocal.Location = new Point(pad, 6);
                rdoGoogle.Location = new Point(rdoLocal.Right + 16, 6);

                // 설정 패널 위치
                int settingsTop = rdoLocal.Bottom + 4;
                int settingsHeight = 76;
                panelLocalSettings.SetBounds(pad, settingsTop, cx - pad * 2, settingsHeight);
                panelGoogleSettings.SetBounds(pad, settingsTop, cx - pad * 2, settingsHeight);

                separator.Width = cx - pad * 2;
                separator.Location = new Point(pad, settingsTop + settingsHeight + 4);

                var szIns = TextRenderer.MeasureText(btnInsert.Text, btnInsert.Font);
                btnInsert.Size = new Size(szIns.Width + 20, 30);
                btnInsert.Location = new Point((cx - btnInsert.Width) / 2, separator.Bottom + 5);
            };

            // ============================================================
            // Form — 원래 순서 그대로 (panelMid=Fill, panelBottom=Bottom, panelTop=Top)
            // ============================================================
            this.Controls.Add(panelMid);
            this.Controls.Add(panelBottom);
            this.Controls.Add(panelTop);

            this.Text = "YUNSEUL-ZERO";
            this.BackColor = BG;
            this.Size = new Size(640, 700);
            this.MinimumSize = new Size(560, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.AutoScaleMode = AutoScaleMode.Font;
            this.DoubleBuffered = true;

            this.Resize += (s, e) => this.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 10, 10));
            this.HandleCreated += (s, e) => this.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 10, 10));

            this.ResumeLayout(false);
        }
    }
}
