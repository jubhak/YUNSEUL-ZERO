using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;

namespace DataParser
{
    public partial class Form1 : Form
    {
        private string? _excelPath;
        private Point _dragStart;
        private bool _dragging;
        private static readonly string SettingsFile = Path.Combine(AppContext.BaseDirectory, "settings.json");
        private static readonly int ROW_COUNT = 13;

        // 구글 시트 관련
        private GoogleSheetService? _googleService;
        private string? _sheetUrl;
        // OAuth Client 정보 (oauth.json에 저장, 저장소에 올리지 않음)
        private static readonly string OAuthSettingsFile = Path.Combine(AppContext.BaseDirectory, "oauth.json");
        private string _oauthClientId = "";
        private string _oauthClientSecret = "";
        private static readonly string[] DefaultKeywords = new[]
        {
            "ID",
            "Networker Name, networker",
            "1., Full Name, name",
            "2., Gender",
            "3., Birthday, date of birth, dob",
            "4., Contact, phone, mobile",
            "5., Occupation, job",
            "6., City of Residence, city, residence",
            "7., Church Name, church",
            "8., Denomination",
            "9., Church Position, position",
            "10., Married/Single, marital status",
            "11., Facebook address, facebook, fb"
        };

        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            this.Icon = CreateAutoInputIcon();
            EnableFormDrag();
            AddCloseButton();
            InitTable();
            LoadOAuthSettings();
            LoadSettings();
            this.Shown += (s, e) => { txtInput.Focus(); TryAutoGoogleLogin(); };
        }

        private static Icon CreateAutoInputIcon()
        {
            int sz = 64;
            using var bmp = new Bitmap(sz, sz);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.Transparent);
            using var bgBrush = new SolidBrush(Color.FromArgb(22, 27, 34));
            using var bgPath = new System.Drawing.Drawing2D.GraphicsPath();
            int r = 12;
            bgPath.AddArc(0, 0, r, r, 180, 90); bgPath.AddArc(sz - r, 0, r, r, 270, 90);
            bgPath.AddArc(sz - r, sz - r, r, r, 0, 90); bgPath.AddArc(0, sz - r, r, r, 90, 90);
            bgPath.CloseFigure(); g.FillPath(bgBrush, bgPath);
            using var docPen = new Pen(Color.FromArgb(88, 166, 255), 2.2f);
            g.DrawRectangle(docPen, 12, 8, 22, 28);
            using var linePen = new Pen(Color.FromArgb(139, 148, 158), 1.5f);
            g.DrawLine(linePen, 16, 15, 30, 15); g.DrawLine(linePen, 16, 20, 28, 20); g.DrawLine(linePen, 16, 25, 26, 25);
            using var arrowPen = new Pen(Color.FromArgb(63, 185, 80), 3f);
            arrowPen.EndCap = System.Drawing.Drawing2D.LineCap.Round; arrowPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
            g.DrawLine(arrowPen, 44, 14, 44, 40); g.DrawLine(arrowPen, 37, 33, 44, 42); g.DrawLine(arrowPen, 51, 33, 44, 42);
            using var basePen = new Pen(Color.FromArgb(63, 185, 80), 2f);
            g.DrawLine(basePen, 36, 46, 52, 46);
            using var numFont = new Font("Consolas", 10f, FontStyle.Bold);
            using var numBrush = new SolidBrush(Color.FromArgb(88, 166, 255));
            g.DrawString("0", numFont, numBrush, 22, 40);
            return Icon.FromHandle(bmp.GetHicon());
        }

        private void InitTable()
        {
            for (int i = 0; i < DefaultKeywords.Length && i < _txtKeywords.Length; i++)
                _txtKeywords[i].Text = DefaultKeywords[i];

            // 최초 실행 기본값: "ID" (인덱스 0) UNIQUE 선택
            if (_rdoUnique.Length > 0)
                _rdoUnique[0].Checked = true;
        }

        private void EnableFormDrag()
        {
            foreach (Control c in new Control[] { panelTop, lblInput, panelMid, lblResult, panelBottom })
            {
                c.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { _dragging = true; _dragStart = e.Location; } };
                c.MouseMove += (s, e) => { if (_dragging) { var p = PointToScreen(e.Location); SetDesktopLocation(p.X - _dragStart.X, p.Y - _dragStart.Y); } };
                c.MouseUp += (s, e) => _dragging = false;
            }
        }

        private void LayoutTable()
        {
            int w = panelTable.ClientSize.Width;
            int colW = 50, uniW = 50, rowH = 24;
            int titleW = (int)(w * 0.42);
            int valW = w - titleW - uniW - colW;
            int rowCount = _txtKeywords.Length;
            while (_hLines.Length <= rowCount)
            {
                var nl = new Panel { BackColor = Color.FromArgb(48, 54, 61), Height = 1 };
                panelTable.Controls.Add(nl);
                Array.Resize(ref _hLines, _hLines.Length + 1);
                _hLines[_hLines.Length - 1] = nl;
            }
            int y = 0;
            for (int i = 0; i < rowCount; i++)
            {
                _hLines[i].SetBounds(0, y, w, 1); _hLines[i].Visible = true; y += 1;
                _txtKeywords[i].SetBounds(4, y + 5, titleW - 8, rowH - 4);
                _txtValues[i].SetBounds(titleW + 4, y + 5, valW - 8, rowH - 4);
                _rdoUnique[i].SetBounds(titleW + valW + (uniW - 16) / 2, y + (rowH - 16) / 2, 16, 16);
                _txtColNums[i].SetBounds(titleW + valW + uniW + 4, y + 2, colW - 8, rowH - 4);
                y += rowH;
            }
            _hLines[rowCount].SetBounds(0, y, w, 1); _hLines[rowCount].Visible = true;
            for (int i = rowCount + 1; i < _hLines.Length; i++) _hLines[i].Visible = false;
            _vLine1.SetBounds(titleW, 0, 1, y + 1);
            _vLine2.SetBounds(titleW + valW, 0, 1, y + 1);
            _vLine3.SetBounds(titleW + valW + uniW, 0, 1, y + 1);
            _vLine1.BringToFront(); _vLine2.BringToFront(); _vLine3.BringToFront();
            panelTable.AutoScrollPosition = new Point(0, 0);
            panelTable.AutoScrollMinSize = new Size(0, y + 2);

            // 헤더도 동일한 w 기준으로 레이아웃 (스크롤바 포함된 동일 너비)
            int hdrH = panelHeader.Height;
            hdrTitle.SetBounds(0, 0, titleW, hdrH);
            hdrValue.SetBounds(titleW, 0, valW, hdrH);
            hdrUnique.SetBounds(titleW + valW, 0, uniW, hdrH);
            hdrCol.SetBounds(titleW + valW + uniW, 0, colW, hdrH);
        }

        private void AddCloseButton()
        {
            var b = new Label { Text = "\u2715", ForeColor = Color.FromArgb(139,148,158), BackColor = Color.Transparent,
                Font = new Font("Segoe UI",9F,FontStyle.Bold), Size = new Size(28,22),
                TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand };
            b.Click += (s, e) => Close();
            b.MouseEnter += (s, e) => b.ForeColor = Color.FromArgb(248,81,73);
            b.MouseLeave += (s, e) => b.ForeColor = Color.FromArgb(139,148,158);
            this.Controls.Add(b); b.BringToFront();
            this.Resize += (s, e) => b.Location = new Point(this.ClientSize.Width - 32, 4);
            b.Location = new Point(this.ClientSize.Width - 32, 4);
        }

        private void SaveSettings()
        {
            try {
                var lines = new List<string> { "EXCEL=" + (_excelPath ?? ""), "ROWS=" + _txtKeywords.Length };
                // UNIQUE 선택 인덱스 저장
                int uniqueIdx = -1;
                for (int i = 0; i < _rdoUnique.Length; i++)
                    if (_rdoUnique[i].Checked) { uniqueIdx = i; break; }
                lines.Add("UNIQUE=" + uniqueIdx);
                // 구글 시트 설정 저장
                lines.Add("MODE=" + (rdoGoogle.Checked ? "GOOGLE" : "LOCAL"));
                lines.Add("SHEETURL=" + (_sheetUrl ?? ""));
                for (int i = 0; i < _txtKeywords.Length; i++)
                { lines.Add($"KW{i}=" + _txtKeywords[i].Text.Trim()); lines.Add($"COL{i}=" + _txtColNums[i].Text.Trim()); }
                File.WriteAllLines(SettingsFile, lines);
            } catch { }
        }

        private void LoadSettings()
        {
            try {
                if (!File.Exists(SettingsFile)) return;
                var saved = new Dictionary<string, string>(); string? savedExcel = null; int savedRows = ROW_COUNT; int savedUnique = -1;
                string savedMode = "LOCAL"; string? savedSheetUrl = null;
                foreach (string line in File.ReadAllLines(SettingsFile))
                { int eq = line.IndexOf('='); if (eq < 0) continue; string key = line[..eq], val = line[(eq+1)..];
                  if (key == "EXCEL") savedExcel = val;
                  else if (key == "ROWS" && int.TryParse(val, out int rc)) savedRows = rc;
                  else if (key == "UNIQUE" && int.TryParse(val, out int ui)) savedUnique = ui;
                  else if (key == "MODE") savedMode = val;
                  else if (key == "SHEETURL") savedSheetUrl = val;
                  else saved[key] = val; }
                if (!string.IsNullOrEmpty(savedExcel) && File.Exists(savedExcel))
                { _excelPath = savedExcel; lblExcelPath.Text = Path.GetFileName(_excelPath); lblExcelPath.ForeColor = Color.FromArgb(88,166,255); lblExcelPath.Cursor = Cursors.Hand; LoadSheetNames(); }
                // 구글 시트 설정 복원
                if (!string.IsNullOrEmpty(savedSheetUrl))
                { _sheetUrl = savedSheetUrl; txtSheetUrl.Text = savedSheetUrl; txtSheetUrl.ForeColor = Color.FromArgb(230, 237, 243); }
                if (savedMode == "GOOGLE") { rdoGoogle.Checked = true; } else { rdoLocal.Checked = true; }
                while (_txtKeywords.Length < savedRows) AddTableRow();
                for (int i = 0; i < _txtKeywords.Length; i++)
                { if (saved.TryGetValue($"KW{i}", out string? kw) && !string.IsNullOrEmpty(kw)) _txtKeywords[i].Text = kw;
                  if (saved.TryGetValue($"COL{i}", out string? col))
                  {
                      // 기존 숫자 형식 호환: 숫자면 알파벳으로 변환
                      if (!string.IsNullOrEmpty(col) && int.TryParse(col.Trim(), out int numVal) && numVal > 0)
                          _txtColNums[i].Text = ColumnNumberToLetter(numVal);
                      else
                          _txtColNums[i].Text = col?.Trim().ToUpper() ?? "";
                  } }
                // UNIQUE 복원
                if (savedUnique >= 0 && savedUnique < _rdoUnique.Length)
                {
                    for (int i = 0; i < _rdoUnique.Length; i++) _rdoUnique[i].Checked = false;
                    _rdoUnique[savedUnique].Checked = true;
                }
            } catch { }
        }

        private void BtnProcess_Click(object? sender, EventArgs e)
        {
            string raw = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(raw)) { MessageBox.Show("데이터를 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            ParseAllRecords(raw);
        }

        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Title = "엑셀 파일 선택", Filter = "Excel Files|*.xlsx;*.xls", RestoreDirectory = true };
            if (dlg.ShowDialog() == DialogResult.OK)
            { _excelPath = dlg.FileName; lblExcelPath.Text = Path.GetFileName(_excelPath); lblExcelPath.ForeColor = Color.FromArgb(88,166,255); lblExcelPath.Cursor = Cursors.Hand; LoadSheetNames(); SaveSettings(); }
        }

        private void LoadSheetNames()
        {
            cboSheet.Items.Clear();
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath)) return;
            try {
                using var pkg = new ExcelPackage(new FileInfo(_excelPath));
                int autoIdx = -1;
                for (int i = 0; i < pkg.Workbook.Worksheets.Count; i++)
                { string n = pkg.Workbook.Worksheets[i].Name; cboSheet.Items.Add(n);
                  if (n.Equals("Lahore2", StringComparison.OrdinalIgnoreCase)) autoIdx = i; }
                if (autoIdx >= 0) cboSheet.SelectedIndex = autoIdx; else if (cboSheet.Items.Count > 0) cboSheet.SelectedIndex = 0;
            } catch { }
        }

        private void BtnInsert_Click(object? sender, EventArgs e)
        {
            // 구글 시트 모드
            if (rdoGoogle.Checked)
            {
                BtnInsertGoogle_Click();
                return;
            }

            // 로컬 엑셀 모드
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            { MessageBox.Show("먼저 엑셀 파일을 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (cboSheet.SelectedIndex < 0)
            { MessageBox.Show("엑셀 시트를 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            // 키워드가 비어있는 행 확인
            var emptyRows = new List<int>();
            for (int i = 0; i < _txtKeywords.Length; i++)
                if (string.IsNullOrWhiteSpace(_txtKeywords[i].Text)) emptyRows.Add(i);

            if (emptyRows.Count > 0)
            {
                var result = MessageBox.Show("키워드 값이 없는 항목이 있습니다.\n해당 행을 삭제하시겠습니까?",
                    "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    RemoveRows(emptyRows);
                    SaveSettings();
                }
                ResetScroll();
            }

            // 현재 표시 중인 레코드의 수정사항 저장
            SaveCurrentRecordValues();

            // 다건 레코드가 있으면 다건 입력, 없으면 현재 테이블 값으로 단건 입력
            int recordCount = _allRecords.Count;
            if (recordCount <= 1)
            {
                // 단건: 현재 테이블 값 사용 (기존 로직)
                var singleValues = new string[_txtKeywords.Length];
                for (int i = 0; i < _txtKeywords.Length; i++)
                    singleValues[i] = _txtValues[i].Text.Trim();
                InsertRecords(new List<string[]> { singleValues });
            }
            else
            {
                // 다건
                var dlgResult = MessageBox.Show($"{recordCount}건을 입력합니다.",
                    "알림", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dlgResult != DialogResult.OK) return;
                InsertRecords(_allRecords);
            }
        }

        /// <summary>
        /// 레코드 목록을 엑셀에 입력합니다. 단건/다건 공통 로직.
        /// </summary>
        private void InsertRecords(List<string[]> records)
        {
            string selectedSheet = cboSheet.SelectedItem?.ToString() ?? "";

            // COL 매핑 확인
            var colMappings = new List<(int rowIdx, int col)>();
            for (int i = 0; i < _txtKeywords.Length; i++)
            {
                string colStr = _txtColNums[i].Text.Trim();
                int colIdx = ColumnLetterToNumber(colStr);
                if (colIdx > 0) colMappings.Add((i, colIdx));
            }
            if (colMappings.Count == 0) { MessageBox.Show("컬럼(COL)이 지정된 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            // 데이터가 있는 레코드만 필터
            var validRecords = new List<string[]>();
            foreach (var rec in records)
            {
                bool hasData = colMappings.Any(cm => cm.rowIdx < rec.Length && !string.IsNullOrEmpty(rec[cm.rowIdx]));
                if (hasData) validRecords.Add(rec);
            }
            if (validRecords.Count == 0) { MessageBox.Show("입력할 데이터가 없습니다.\n먼저 데이터를 입력하고 1차 가공을 해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            // UNIQUE 선택 확인
            int uniqueIdx = -1;
            for (int i = 0; i < _rdoUnique.Length; i++)
                if (_rdoUnique[i].Checked) { uniqueIdx = i; break; }

            try {
                byte[] fileBytes = File.ReadAllBytes(_excelPath!);
                using var stream = new MemoryStream(fileBytes);
                using var pkg = new ExcelPackage(stream);
                var ws = pkg.Workbook.Worksheets[selectedSheet];
                if (ws == null) { MessageBox.Show($"시트 '{selectedSheet}'를 찾을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                int totalRows = ws.Dimension?.End.Row ?? 0;
                int totalCols = ws.Dimension?.End.Column ?? 0;

                // 데이터 유효성 검사 목록 로드 (드롭다운 선택항목 대소문자 매칭용)
                var validationLists = GetExcelValidationLists(ws);

                int insertedCount = 0;
                int updatedCount = 0;
                int firstNewRow = -1;
                var updatedRows = new List<int>();

                foreach (var rec in validRecords)
                {
                    if (uniqueIdx >= 0 && uniqueIdx < rec.Length)
                    {
                        string uniqueValue = CleanSpecialChars(rec[uniqueIdx].Trim());
                        int uniqueCol = colMappings.FirstOrDefault(cm => cm.rowIdx == uniqueIdx).col;
                        if (uniqueCol <= 0 || string.IsNullOrEmpty(uniqueValue))
                        {
                            // UNIQUE 값이 없으면 새 행에 추가
                            int newRow = FindLastDataRow(ws, totalRows, totalCols) + 1;
                            WriteRecordToRow(ws, newRow, rec, colMappings, validationLists);
                            if (firstNewRow < 0) firstNewRow = newRow;
                            insertedCount++;
                            totalRows = Math.Max(totalRows, newRow);
                            continue;
                        }

                        // 엑셀에서 해당 컬럼에서 동일한 값 검색
                        int foundRow = -1;
                        for (int row = 1; row <= totalRows; row++)
                        {
                            var cellVal = ws.Cells[row, uniqueCol].Value;
                            if (cellVal != null && cellVal.ToString()!.Trim().Equals(uniqueValue, StringComparison.OrdinalIgnoreCase))
                            { foundRow = row; break; }
                        }

                        if (foundRow > 0)
                        {
                            // 기존 행 업데이트
                            foreach (var (rowIdx, col) in colMappings)
                            {
                                if (rowIdx == uniqueIdx) continue;
                                string val = (rowIdx < rec.Length) ? rec[rowIdx].Trim() : "";
                                if (string.IsNullOrEmpty(val)) continue;
                                val = CleanSpecialChars(val);
                                // 데이터 유효성 검사 목록이 있으면 대소문자 무시 매칭
                                if (validationLists.TryGetValue(col, out var validList))
                                    val = MatchValidationValue(val, validList);
                                ws.Cells[foundRow, col].Value = val;
                            }
                            updatedRows.Add(foundRow);
                            updatedCount++;
                        }
                        else
                        {
                            int newRow = FindLastDataRow(ws, totalRows, totalCols) + 1;
                            WriteRecordToRow(ws, newRow, rec, colMappings, validationLists);
                            if (firstNewRow < 0) firstNewRow = newRow;
                            insertedCount++;
                            totalRows = Math.Max(totalRows, newRow);
                        }
                    }
                    else
                    {
                        int newRow = FindLastDataRow(ws, totalRows, totalCols) + 1;
                        WriteRecordToRow(ws, newRow, rec, colMappings, validationLists);
                        if (firstNewRow < 0) firstNewRow = newRow;
                        insertedCount++;
                        totalRows = Math.Max(totalRows, newRow);
                    }
                }

                File.WriteAllBytes(_excelPath!, pkg.GetAsByteArray());

                // 결과 메시지
                var msgParts = new List<string>();
                if (insertedCount > 0)
                    msgParts.Add($"{firstNewRow}행부터 {insertedCount}건이 입력되었습니다.");
                if (updatedCount > 0)
                {
                    string rowList = string.Join(", ", updatedRows.Select(r => $"{r}행"));
                    msgParts.Add($"{rowList} 수정되었습니다.");
                }
                string msg = string.Join("\n", msgParts);
                MessageBox.Show(msg, "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetAfterInsert();

            } catch (Exception ex) {
                string msg = ex.Message ?? "";
                string innerMsg = ex.InnerException?.Message ?? "";
                string allMsg = msg + " " + innerMsg;
                if (ex is IOException
                    || allMsg.Contains("being used by another process", StringComparison.OrdinalIgnoreCase)
                    || allMsg.Contains("locked", StringComparison.OrdinalIgnoreCase)
                    || allMsg.Contains("Error saving file", StringComparison.OrdinalIgnoreCase)
                    || allMsg.Contains("denied", StringComparison.OrdinalIgnoreCase))
                    MessageBox.Show("엑셀 파일이 열려 있어 수정할 수 없습니다.\n파일을 닫고 다시 시도해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else MessageBox.Show($"엑셀 저장 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 워크시트에서 데이터가 있는 마지막 행 번호를 반환합니다.
        /// </summary>
        private static int FindLastDataRow(ExcelWorksheet ws, int totalRows, int totalCols)
        {
            int lastDataRow = 0;
            for (int row = 1; row <= totalRows; row++)
                for (int c = 1; c <= totalCols; c++)
                { var v = ws.Cells[row, c].Value; if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) { lastDataRow = row; break; } }
            return lastDataRow;
        }

        /// <summary>
        /// 레코드 값을 지정된 행에 기록합니다.
        /// </summary>
        private void WriteRecordToRow(ExcelWorksheet ws, int row, string[] rec, List<(int rowIdx, int col)> colMappings, Dictionary<int, List<string>>? validationLists = null)
        {
            foreach (var (rowIdx, col) in colMappings)
            {
                string val = (rowIdx < rec.Length) ? rec[rowIdx].Trim() : "";
                if (string.IsNullOrEmpty(val)) continue;
                val = CleanSpecialChars(val);
                // 데이터 유효성 검사 목록이 있으면 대소문자 무시 매칭
                if (validationLists != null && validationLists.TryGetValue(col, out var validList))
                    val = MatchValidationValue(val, validList);
                ws.Cells[row, col].Value = val;
            }
        }

        private void ResetAfterInsert()
        {
            txtInput.Text = "";
            _allRecords.Clear();
            _currentRecordIndex = 0;
            for (int i = 0; i < _txtKeywords.Length; i++) _txtValues[i].Text = "";
            lblUpsertStatus.Visible = false;
            lblUpsertStatus.Text = "";
            UpdateRecordButtons();
            ResetScroll();
            txtInput.Focus();
        }

        /// <summary>
        /// 입력 텍스트에서 데이터를 추출합니다. 빈 줄은 무시하고 모든 줄을 하나의 풀로 취급합니다.
        /// 한번 매칭된 줄은 재사용하지 않으며, 남은 줄에서 추가 레코드를 반복 추출합니다.
        /// 최대 MAX_RECORDS(20)건까지 추출합니다.
        /// </summary>
        private void ParseAllRecords(string rawText)
        {
            rawText = CleanMessageHeader(rawText);
            rawText = CleanFacebookAdd(rawText);
            _allRecords.Clear();
            _currentRecordIndex = 0;

            // 빈 줄 무시, 모든 줄을 하나의 풀로 취급
            string[] allLines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(l => l.Trim()).Where(l => l.Length > 0).ToArray();

            if (allLines.Length == 0)
            {
                UpdateRecordButtons();
                return;
            }

            // 첫 번째 레코드 추출
            var usedGlobal = new HashSet<int>();
            string[] firstValues = ParseSingleRecordTracked(allLines, usedGlobal);
            if (firstValues.Any(v => !string.IsNullOrEmpty(v)))
                _allRecords.Add(firstValues);

            // 남은 줄에서 추가 레코드 반복 추출
            while (_allRecords.Count < MAX_RECORDS)
            {
                // 아직 사용되지 않은 줄만 모음
                var remainingLines = new List<string>();
                var remainingIndices = new List<int>();
                for (int i = 0; i < allLines.Length; i++)
                {
                    if (!usedGlobal.Contains(i))
                    {
                        remainingLines.Add(allLines[i]);
                        remainingIndices.Add(i);
                    }
                }
                if (remainingLines.Count == 0) break;

                var localUsed = new HashSet<int>();
                string[] values = ParseSingleRecordTracked(remainingLines.ToArray(), localUsed);

                if (values.Count(v => !string.IsNullOrEmpty(v)) < 2) break; // 유효 데이터 부족하면 중단

                _allRecords.Add(values);
                // localUsed의 인덱스를 원래 allLines 인덱스로 변환하여 usedGlobal에 추가
                foreach (int li in localUsed)
                    usedGlobal.Add(remainingIndices[li]);
            }

            // 첫 번째 레코드 표시
            if (_allRecords.Count > 0)
                DisplayRecord(0);
            else
                for (int i = 0; i < _txtKeywords.Length; i++) _txtValues[i].Text = "";

            UpdateRecordButtons();
        }

        /// <summary>
        /// 줄 배열에서 키워드 기반으로 값을 추출하고, 사용된 줄 인덱스를 usedLines에 기록합니다.
        /// </summary>
        private string[] ParseSingleRecordTracked(string[] inputLines, HashSet<int> usedLines)
        {
            var values = new string[_txtKeywords.Length];
            for (int i = 0; i < _txtKeywords.Length; i++)
            {
                string kwText = _txtKeywords[i].Text.Trim();
                if (string.IsNullOrEmpty(kwText)) { values[i] = ""; continue; }
                string[] keywords = kwText.Split(',').Select(k => k.Trim()).Where(k => k.Length > 0).ToArray();
                var textKws = new List<string>(); var numPats = new List<int>();
                foreach (string kw in keywords)
                { var m = Regex.Match(kw, @"^(\d+)\.$"); if (m.Success) numPats.Add(int.Parse(m.Groups[1].Value)); else textKws.Add(kw); }
                string? found = null;
                int foundLineIdx = -1;
                foreach (string kw in textKws) { (found, foundLineIdx) = SearchByTextKeywordEx(inputLines, kw, usedLines); if (found != null) break; }
                if (found == null) foreach (int num in numPats) { (found, foundLineIdx) = SearchByNumberPatternEx(inputLines, num, usedLines); if (found != null) break; }
                if (found != null && foundLineIdx >= 0) usedLines.Add(foundLineIdx);
                values[i] = found ?? "";
            }
            return values;
        }

        private static (string? value, int lineIndex) SearchByTextKeywordEx(string[] lines, string keyword, HashSet<int> usedLines)
        {
            string kwLow = keyword.ToLower();
            for (int li = 0; li < lines.Length; li++)
            {
                if (usedLines.Contains(li)) continue;
                string line = lines[li].Trim(); if (string.IsNullOrEmpty(line)) continue;
                string stripped = Regex.Replace(line, @"^\d+\.\s*", "");
                int ci = stripped.IndexOf(':'); if (ci < 0) continue;
                string keyPart = Regex.Replace(stripped[..ci].Trim(), @"\(.*?\)", "").Trim();
                string valuePart = stripped[(ci+1)..].Trim();
                if (keyPart.StartsWith("Is he", StringComparison.OrdinalIgnoreCase) || keyPart.StartsWith("does he", StringComparison.OrdinalIgnoreCase)) continue;
                if (keyPart.Equals("Personal Information", StringComparison.OrdinalIgnoreCase)) continue;
                if (keyPart.StartsWith("L2", StringComparison.OrdinalIgnoreCase) && keyPart.Contains("Form", StringComparison.OrdinalIgnoreCase)) continue;
                string keyLow = keyPart.ToLower();
                bool matched;
                if (kwLow.Length <= 2)
                    matched = keyLow == kwLow;
                else
                    matched = keyLow == kwLow || keyLow.Contains(kwLow) || kwLow.Contains(keyLow);
                if (matched)
                {
                    if (string.IsNullOrEmpty(valuePart) || valuePart.StartsWith("("))
                    { int lc = stripped.LastIndexOf(':'); if (lc > ci) valuePart = stripped[(lc+1)..].Trim(); }
                    if (!string.IsNullOrEmpty(valuePart)) return (valuePart, li);
                }
            }
            if (kwLow == "id")
            {
                for (int li = 0; li < lines.Length; li++)
                {
                    if (usedLines.Contains(li)) continue;
                    string l = Regex.Replace(lines[li].Trim(), @"[()]", "").Trim();
                    if (Regex.IsMatch(l, @"^[A-Za-z]\d+[-\s]*\d+$")) return (l.Replace(" ", ""), li);
                }
                // P34-078, P34- 043 같은 패턴에서 "-" 뒤 숫자를 추출하여 앞의 0 제거 (078 → 78)
                for (int li = 0; li < lines.Length; li++)
                {
                    if (usedLines.Contains(li)) continue;
                    var pm = Regex.Match(lines[li].Trim(), @"[A-Za-z]\d+-\s*(\d+)");
                    if (pm.Success)
                    {
                        string numPart = pm.Groups[1].Value.TrimStart('0');
                        if (string.IsNullOrEmpty(numPart)) numPart = "0";
                        return (numPart, li);
                    }
                }
            }
            return (null, -1);
        }

        private static (string? value, int lineIndex) SearchByNumberPatternEx(string[] lines, int number, HashSet<int> usedLines)
        {
            for (int li = 0; li < lines.Length; li++)
            {
                if (usedLines.Contains(li)) continue;
                string line = lines[li].Trim(); if (string.IsNullOrEmpty(line)) continue;
                var m = Regex.Match(line, @"^(\d+)\.\s*(.*)$"); if (!m.Success || int.Parse(m.Groups[1].Value) != number) continue;
                string rest = m.Groups[2].Value.Trim(); int ci = rest.IndexOf(':');
                if (ci >= 0) { string v = rest[(ci+1)..].Trim();
                    if (string.IsNullOrEmpty(v) || v.StartsWith("(")) { int lc = rest.LastIndexOf(':'); if (lc > ci) v = rest[(lc+1)..].Trim(); } return (v, li); }
                else return (rest, li);
            }
            return (null, -1);
        }

        /// <summary>
        /// 지정된 인덱스의 레코드 값을 테이블에 표시합니다.
        /// </summary>
        private void DisplayRecord(int index)
        {
            if (index < 0 || index >= _allRecords.Count) return;
            _currentRecordIndex = index;
            string[] values = _allRecords[index];
            for (int i = 0; i < _txtKeywords.Length; i++)
                _txtValues[i].Text = (i < values.Length) ? values[i] : "";
            UpdateRecordButtonStyles();
            CheckUpsertStatus();
        }

        /// <summary>
        /// 레코드 번호 버튼 클릭 시 호출됩니다.
        /// </summary>
        private void SelectRecord(int index)
        {
            if (index < 0 || index >= _allRecords.Count) return;
            // 현재 표시 중인 레코드의 수정사항을 저장
            SaveCurrentRecordValues();
            DisplayRecord(index);
        }

        /// <summary>
        /// 현재 테이블에 표시된 값을 _allRecords에 반영합니다.
        /// </summary>
        private void SaveCurrentRecordValues()
        {
            if (_currentRecordIndex < 0 || _currentRecordIndex >= _allRecords.Count) return;
            string[] values = _allRecords[_currentRecordIndex];
            // 키워드 행이 추가되어 배열 크기가 달라졌을 수 있으므로 맞춤
            if (values.Length < _txtKeywords.Length)
            {
                var newValues = new string[_txtKeywords.Length];
                Array.Copy(values, newValues, values.Length);
                _allRecords[_currentRecordIndex] = newValues;
                values = newValues;
            }
            for (int i = 0; i < _txtKeywords.Length; i++)
                values[i] = _txtValues[i].Text;
        }

        /// <summary>
        /// 레코드 버튼의 표시/숨김 및 위치를 업데이트합니다.
        /// </summary>
        private void UpdateRecordButtons()
        {
            for (int i = 0; i < MAX_RECORDS; i++)
                _recordBtns[i].Visible = (i < _allRecords.Count);
            UpdateRecordButtonStyles();
            LayoutRecordButtons();
        }

        /// <summary>
        /// 현재 선택된 레코드 버튼의 스타일을 업데이트합니다.
        /// </summary>
        private void UpdateRecordButtonStyles()
        {
            Color BLUE = Color.FromArgb(88, 166, 255);
            Color INACTIVE_BG = Color.FromArgb(30, 36, 44);
            Color FG2 = Color.FromArgb(139, 148, 158);
            for (int i = 0; i < MAX_RECORDS; i++)
            {
                if (i == _currentRecordIndex && _allRecords.Count >= 1)
                {
                    _recordBtns[i].BackColor = BLUE;
                    _recordBtns[i].ForeColor = Color.FromArgb(13, 17, 23);
                }
                else
                {
                    _recordBtns[i].BackColor = INACTIVE_BG;
                    _recordBtns[i].ForeColor = FG2;
                }
                _recordBtns[i].Invalidate();
            }
        }

        /// <summary>
        /// 레코드 버튼의 위치를 레이아웃합니다. + 버튼 우측 30px부터 배치.
        /// </summary>
        private void LayoutRecordButtons()
        {
            // btnAddRow는 panelResultHeader 안에 있음 — lblResult.Right + 6 위치
            // 레코드 버튼은 btnAddRow 우측 30px부터
            Label? btnAdd = null;
            foreach (Control c in _recordBtns[0].Parent?.Controls ?? new Control.ControlCollection(this))
                if (c is Label lbl && lbl.Text == "+") { btnAdd = lbl; break; }

            int startX = (btnAdd != null) ? btnAdd.Right + 30 : 120;
            int y = 2;
            int gap = 2;
            int lastRight = startX;
            for (int i = 0; i < MAX_RECORDS; i++)
            {
                if (!_recordBtns[i].Visible) continue;
                _recordBtns[i].Location = new Point(startX + i * (_recordBtns[i].Width + gap), y);
                lastRight = _recordBtns[i].Right;
            }
            // UPDATE/INSERT 라벨 위치: 마지막 레코드 버튼 우측 10px
            if (lblUpsertStatus.Visible)
                lblUpsertStatus.Location = new Point(lastRight + 10, 4);
        }

        /// <summary>
        /// 현재 표시 중인 레코드의 UNIQUE 값을 기준으로 엑셀에 이미 존재하는지 확인하여
        /// (UPDATE) 또는 (INSERT) 상태를 표시합니다.
        /// </summary>
        private void CheckUpsertStatus()
        {
            lblUpsertStatus.Visible = false;
            lblUpsertStatus.Text = "";

            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath)) return;
            if (cboSheet.SelectedIndex < 0) return;

            // UNIQUE 선택 확인
            int uniqueIdx = -1;
            for (int i = 0; i < _rdoUnique.Length; i++)
                if (_rdoUnique[i].Checked) { uniqueIdx = i; break; }
            if (uniqueIdx < 0) return;

            string uniqueValue = _txtValues[uniqueIdx].Text.Trim();
            string uniqueColStr = _txtColNums[uniqueIdx].Text.Trim();
            int uniqueCol = ColumnLetterToNumber(uniqueColStr);
            if (uniqueCol <= 0 || string.IsNullOrEmpty(uniqueValue)) return;

            string selectedSheet = cboSheet.SelectedItem?.ToString() ?? "";
            try
            {
                byte[] fileBytes = File.ReadAllBytes(_excelPath);
                using var stream = new MemoryStream(fileBytes);
                using var pkg = new ExcelPackage(stream);
                var ws = pkg.Workbook.Worksheets[selectedSheet];
                if (ws == null) return;
                int totalRows = ws.Dimension?.End.Row ?? 0;
                int totalCols = ws.Dimension?.End.Column ?? 0;

                string cleanValue = CleanSpecialChars(uniqueValue);
                int foundRow = -1;
                for (int row = 1; row <= totalRows; row++)
                {
                    var cellVal = ws.Cells[row, uniqueCol].Value;
                    if (cellVal != null && cellVal.ToString()!.Trim().Equals(cleanValue, StringComparison.OrdinalIgnoreCase))
                    { foundRow = row; break; }
                }

                if (foundRow > 0)
                    lblUpsertStatus.Text = $"UPDATE Row {foundRow}";
                else
                {
                    int lastDataRow = FindLastDataRow(ws, totalRows, totalCols);
                    lblUpsertStatus.Text = $"INSERT Row {lastDataRow + 1}";
                }
                lblUpsertStatus.Visible = true;

                // panelResultHeader 기준 우측 정렬 (레이아웃 완료 후 위치 재계산)
                var parent = lblUpsertStatus.Parent;
                if (parent != null)
                {
                    var textSize = TextRenderer.MeasureText(lblUpsertStatus.Text, lblUpsertStatus.Font);
                    lblUpsertStatus.Location = new Point(parent.ClientSize.Width - textSize.Width - 4, 4);
                    // 레이아웃이 아직 안 끝났을 수 있으므로 한 번 더 보정
                    parent.BeginInvoke(() =>
                    {
                        var sz = TextRenderer.MeasureText(lblUpsertStatus.Text, lblUpsertStatus.Font);
                        lblUpsertStatus.Location = new Point(parent.ClientSize.Width - sz.Width - 4, 4);
                    });
                }
            }
            catch { }
        }

        private void BtnAddRow_Click(object? sender, EventArgs e) => AddTableRow();

        private void ResetScroll()
        {
            panelTable.AutoScrollPosition = new Point(0, 0);
            LayoutTable();
        }

        private void RemoveRows(List<int> indices)
        {
            foreach (int idx in indices.OrderByDescending(x => x))
            {
                panelTable.Controls.Remove(_txtKeywords[idx]);
                panelTable.Controls.Remove(_txtValues[idx]);
                panelTable.Controls.Remove(_rdoUnique[idx]);
                panelTable.Controls.Remove(_txtColNums[idx]);
                _txtKeywords[idx].Dispose();
                _txtValues[idx].Dispose();
                _rdoUnique[idx].Dispose();
                _txtColNums[idx].Dispose();

                var kwList = _txtKeywords.ToList(); kwList.RemoveAt(idx); _txtKeywords = kwList.ToArray();
                var valList = _txtValues.ToList(); valList.RemoveAt(idx); _txtValues = valList.ToArray();
                var rdoList = _rdoUnique.ToList(); rdoList.RemoveAt(idx); _rdoUnique = rdoList.ToArray();
                var colList = _txtColNums.ToList(); colList.RemoveAt(idx); _txtColNums = colList.ToArray();
            }
            panelTable.AutoScrollPosition = new Point(0, 0);
            LayoutTable();
            panelTable.Invalidate();
        }

        private void UniqueRadio_Click(int clickedIdx)
        {
            if (_rdoUnique[clickedIdx].Checked)
            {
                _rdoUnique[clickedIdx].Checked = false;
            }
            else
            {
                for (int i = 0; i < _rdoUnique.Length; i++)
                    _rdoUnique[i].Checked = false;
                _rdoUnique[clickedIdx].Checked = true;
            }
            SaveSettings();
        }

        private void AddTableRow()
        {
            int idx = _txtKeywords.Length;
            Array.Resize(ref _txtKeywords, idx + 1); Array.Resize(ref _txtValues, idx + 1); Array.Resize(ref _rdoUnique, idx + 1); Array.Resize(ref _txtColNums, idx + 1);
            Color CARD = Color.FromArgb(22,27,34); Color FG = Color.FromArgb(230,237,243); Color FG2 = Color.FromArgb(139,148,158);
            _txtKeywords[idx] = new TextBox { BackColor = CARD, ForeColor = FG2, Font = new Font("Consolas",8F), BorderStyle = BorderStyle.None, Text = "" };
            _txtKeywords[idx].Leave += (s, ev) => SaveSettings();
            _txtValues[idx] = new TextBox { BackColor = CARD, ForeColor = FG, Font = new Font("Segoe UI",9F), BorderStyle = BorderStyle.None, Text = "" };
            _rdoUnique[idx] = new RadioButton { BackColor = CARD, ForeColor = FG, AutoCheck = false, Appearance = Appearance.Normal, Text = "", Cursor = Cursors.Hand };
            int capturedIdx = idx;
            _rdoUnique[idx].Click += (s, ev) => UniqueRadio_Click(capturedIdx);
            _txtColNums[idx] = new TextBox { BackColor = Color.FromArgb(30,36,44), ForeColor = FG, Font = new Font("Consolas",9F), BorderStyle = BorderStyle.FixedSingle, TextAlign = HorizontalAlignment.Right, MaxLength = 3, Text = "" };
            _txtColNums[idx].KeyPress += (s, ev) => { if (!char.IsLetter(ev.KeyChar) && !char.IsControl(ev.KeyChar)) ev.Handled = true; };
            _txtColNums[idx].Leave += (s, ev) => { var tb = (TextBox)s!; tb.Text = tb.Text.Trim().ToUpper(); SaveSettings(); };
            panelTable.Controls.Add(_txtKeywords[idx]); panelTable.Controls.Add(_txtValues[idx]); panelTable.Controls.Add(_rdoUnique[idx]); panelTable.Controls.Add(_txtColNums[idx]);
            LayoutTable();
            panelTable.ScrollControlIntoView(_txtKeywords[idx]); _txtKeywords[idx].Focus();
        }

        private static string CleanFacebookAdd(string text)
        { text = Regex.Replace(text, @"\(\s*Facebook\s*Add[^)]*\)", "", RegexOptions.IgnoreCase);
          text = Regex.Replace(text, @"Facebook\s*Add\b[^,\r\n]*", "", RegexOptions.IgnoreCase); return text.Trim(); }

        /// <summary>
        /// 카카오톡 등 메신저 메시지 헤더([날짜] 이름:)를 제거합니다.
        /// 예: "[2026-05-04 오전 1:17] 이 지혜: ✅ P34- 043" → "✅ P34- 043"
        /// 헤더 뒤의 내용은 보존합니다.
        /// </summary>
        private static string CleanMessageHeader(string text)
        {
            // [날짜/시간] 이름: 패턴만 제거
            // ] 뒤의 이름 부분: ":"가 아닌 문자들 + 첫 번째 ":" (lazy)
            text = Regex.Replace(text, @"^\[\d{4}-\d{2}-\d{2}[^\]]*\]\s*[^:：\r\n]*?[：:]\s*", "", RegexOptions.Multiline);
            // 제거 후 생긴 빈 줄 정리 (연속 빈 줄을 하나로)
            text = Regex.Replace(text, @"(\r?\n){3,}", "\n\n");
            return text.Trim();
        }

        private static string CleanSpecialChars(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new System.Text.StringBuilder(text.Length);
            foreach (char c in text)
            { if ((c>='A'&&c<='Z')||(c>='a'&&c<='z')||(c>='0'&&c<='9')) { sb.Append(c); continue; }
              if (c>=0xAC00&&c<=0xD7A3) { sb.Append(c); continue; } if (c>=0x3131&&c<=0x318E) { sb.Append(c); continue; }
              if (" -+/.@,:;'\"()_!?#&%$=~`<>[]{}\\|^".Contains(c)) { sb.Append(c); continue; } }
            return sb.ToString().Trim();
        }

        /// <summary>
        /// 값을 데이터 유효성 검사 목록과 대소문자 무시로 매칭합니다.
        /// 매칭되는 항목이 있으면 해당 항목의 원래 대소문자를 반환하고, 없으면 원래 값을 그대로 반환합니다.
        /// </summary>
        private static string MatchValidationValue(string value, List<string>? validationList)
        {
            if (validationList == null || validationList.Count == 0 || string.IsNullOrEmpty(value))
                return value;

            // 정확히 일치하면 그대로 반환
            var exact = validationList.FirstOrDefault(v => v == value);
            if (exact != null) return exact;

            // 대소문자 무시 매칭
            var matched = validationList.FirstOrDefault(v =>
                v.Equals(value, StringComparison.OrdinalIgnoreCase));
            if (matched != null) return matched;

            // 앞뒤 공백 제거 후 대소문자 무시 매칭
            var trimMatched = validationList.FirstOrDefault(v =>
                v.Trim().Equals(value.Trim(), StringComparison.OrdinalIgnoreCase));
            if (trimMatched != null) return trimMatched;

            return value;
        }

        /// <summary>
        /// EPPlus 워크시트에서 컬럼별 데이터 유효성 검사 목록을 가져옵니다.
        /// 반환: Dictionary&lt;col(1-based), List&lt;string&gt; validValues&gt;
        /// </summary>
        private static Dictionary<int, List<string>> GetExcelValidationLists(ExcelWorksheet ws)
        {
            var result = new Dictionary<int, List<string>>();
            try
            {
                foreach (var dv in ws.DataValidations)
                {
                    if (dv is not ExcelDataValidationList listDv) continue;

                    // 유효성 검사가 적용된 열 번호 추출
                    var address = dv.Address?.Address;
                    if (string.IsNullOrEmpty(address)) continue;

                    // 주소에서 열 번호 추출 (예: "D:D", "D2:D1048576", "$D$2:$D$1048576")
                    var cols = ExtractColumnsFromAddress(address);

                    // 선택항목 목록 가져오기
                    var values = new List<string>();
                    if (listDv.Formula?.Values != null)
                    {
                        foreach (var item in listDv.Formula.Values)
                        {
                            if (!string.IsNullOrEmpty(item)) values.Add(item);
                        }
                    }

                    if (values.Count > 0)
                    {
                        foreach (int col in cols)
                        {
                            if (!result.ContainsKey(col))
                                result[col] = values;
                        }
                    }
                }
            }
            catch { /* 유효성 규칙을 가져올 수 없으면 무시 */ }
            return result;
        }

        /// <summary>
        /// 엑셀 주소 문자열에서 열 번호(1-based)를 추출합니다.
        /// </summary>
        private static List<int> ExtractColumnsFromAddress(string address)
        {
            var cols = new HashSet<int>();
            // 여러 범위가 쉼표로 구분될 수 있음
            var parts = address.Split(',');
            foreach (var part in parts)
            {
                string clean = part.Trim().Replace("$", "");
                // "D:D", "D2:D1048576", "D2" 등에서 열 문자 추출
                var match = Regex.Match(clean, @"^([A-Za-z]+)");
                if (match.Success)
                {
                    int col = ColumnLetterToNumber(match.Groups[1].Value);
                    if (col > 0) cols.Add(col);
                }
            }
            return cols.ToList();
        }

        /// <summary>
        /// 알파벳 컬럼 문자열(A, B, ..., Z, AA, AB, ...)을 엑셀 컬럼 번호(1-based)로 변환합니다.
        /// </summary>
        private static int ColumnLetterToNumber(string colLetter)
        {
            if (string.IsNullOrWhiteSpace(colLetter)) return -1;
            colLetter = colLetter.Trim().ToUpper();
            int result = 0;
            foreach (char c in colLetter)
            {
                if (c < 'A' || c > 'Z') return -1;
                result = result * 26 + (c - 'A' + 1);
            }
            return result;
        }

        /// <summary>
        /// 엑셀 컬럼 번호(1-based)를 알파벳 문자열(A, B, ..., Z, AA, AB, ...)로 변환합니다.
        /// </summary>
        private static string ColumnNumberToLetter(int colNumber)
        {
            if (colNumber <= 0) return "";
            string result = "";
            while (colNumber > 0)
            {
                colNumber--;
                result = (char)('A' + colNumber % 26) + result;
                colNumber /= 26;
            }
            return result;
        }

        // ============================================================
        // 구글 시트 관련 메서드
        // ============================================================

        private void LoadOAuthSettings()
        {
            try
            {
                if (!File.Exists(OAuthSettingsFile)) return;
                foreach (string line in File.ReadAllLines(OAuthSettingsFile))
                {
                    int eq = line.IndexOf('=');
                    if (eq < 0) continue;
                    string key = line[..eq], val = line[(eq + 1)..];
                    if (key == "CLIENT_ID") _oauthClientId = val;
                    else if (key == "CLIENT_SECRET") _oauthClientSecret = val;
                }
            }
            catch { }
        }

        private void SaveOAuthSettings()
        {
            try
            {
                File.WriteAllLines(OAuthSettingsFile, new[]
                {
                    "CLIENT_ID=" + _oauthClientId,
                    "CLIENT_SECRET=" + _oauthClientSecret
                });
            }
            catch { }
        }

        private async void TryAutoGoogleLogin()
        {
            if (string.IsNullOrEmpty(_oauthClientId) || string.IsNullOrEmpty(_oauthClientSecret)) return;
            try
            {
                _googleService ??= new GoogleSheetService();
                bool ok = await _googleService.TryAutoLoginAsync(_oauthClientId, _oauthClientSecret);
                if (ok)
                {
                    lblGoogleStatus.Text = "✓ 로그인됨";
                    lblGoogleStatus.ForeColor = Color.FromArgb(63, 185, 80);
                    btnGoogleConnect.Text = "로그아웃";
                    // 시트 URL이 있으면 자동 연결
                    if (!string.IsNullOrEmpty(_sheetUrl))
                    {
                        try
                        {
                            _googleService.SetSpreadsheet(_sheetUrl);
                            LoadGoogleSheetNames();
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private void ToggleMode()
        {
            panelLocalSettings.Visible = rdoLocal.Checked;
            panelGoogleSettings.Visible = rdoGoogle.Checked;
            SaveSettings();
        }

        private void BtnSelectJson_Click(object? sender, EventArgs e)
        {
            // OAuth 방식에서는 이 버튼이 "설정" 역할 — Client ID/Secret 입력 다이얼로그
            ShowOAuthSetupDialog();
        }

        private void ShowOAuthSetupDialog()
        {
            var dlg = new Form
            {
                Text = "OAuth 설정",
                Size = new Size(450, 200),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(22, 27, 34)
            };

            var lblId = new Label { Text = "Client ID:", ForeColor = Color.FromArgb(230, 237, 243), Location = new Point(12, 20), AutoSize = true };
            var txtId = new TextBox { Text = _oauthClientId, Location = new Point(12, 40), Width = 410, BackColor = Color.FromArgb(30, 36, 44), ForeColor = Color.FromArgb(230, 237, 243) };
            var lblSec = new Label { Text = "Client Secret:", ForeColor = Color.FromArgb(230, 237, 243), Location = new Point(12, 72), AutoSize = true };
            var txtSec = new TextBox { Text = _oauthClientSecret, Location = new Point(12, 92), Width = 410, BackColor = Color.FromArgb(30, 36, 44), ForeColor = Color.FromArgb(230, 237, 243) };
            var btnOk = new Button { Text = "저장", Location = new Point(170, 130), Size = new Size(100, 30), BackColor = Color.FromArgb(63, 185, 80), ForeColor = Color.FromArgb(13, 17, 23), FlatStyle = FlatStyle.Flat };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.Click += (s, ev) =>
            {
                _oauthClientId = txtId.Text.Trim();
                _oauthClientSecret = txtSec.Text.Trim();
                SaveOAuthSettings();
                dlg.DialogResult = DialogResult.OK;
                dlg.Close();
            };

            dlg.Controls.AddRange(new Control[] { lblId, txtId, lblSec, txtSec, btnOk });
            dlg.ShowDialog(this);
        }

        private async void BtnGoogleConnect_Click(object? sender, EventArgs e)
        {
            // 이미 로그인된 상태면 로그아웃
            if (_googleService != null && _googleService.IsConnected)
            {
                _googleService.Logout();
                _googleService = null;
                lblGoogleStatus.Text = "";
                btnGoogleConnect.Text = "로그인";
                cboGoogleSheet.Items.Clear();
                return;
            }

            if (string.IsNullOrEmpty(_oauthClientId) || string.IsNullOrEmpty(_oauthClientSecret))
            {
                MessageBox.Show("먼저 OAuth 설정에서 Client ID와 Secret을 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ShowOAuthSetupDialog();
                return;
            }

            string url = txtSheetUrl.Text.Trim();
            if (string.IsNullOrEmpty(url) || url == "구글 스프레드시트 URL 붙여넣기")
            {
                MessageBox.Show("구글 스프레드시트 URL을 입력해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                lblGoogleStatus.Text = "로그인 중...";
                lblGoogleStatus.ForeColor = Color.FromArgb(255, 200, 50);

                _googleService?.Dispose();
                _googleService = new GoogleSheetService();
                bool ok = await _googleService.LoginAsync(_oauthClientId, _oauthClientSecret);

                if (ok)
                {
                    _sheetUrl = url;
                    _googleService.SetSpreadsheet(url);
                    LoadGoogleSheetNames();

                    lblGoogleStatus.Text = "✓ 로그인됨";
                    lblGoogleStatus.ForeColor = Color.FromArgb(63, 185, 80);
                    btnGoogleConnect.Text = "로그아웃";
                    SaveSettings();
                }
            }
            catch (Exception ex)
            {
                lblGoogleStatus.Text = "✗ 실패";
                lblGoogleStatus.ForeColor = Color.FromArgb(248, 81, 73);
                MessageBox.Show($"구글 로그인 실패:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadGoogleSheetNames()
        {
            cboGoogleSheet.Items.Clear();
            if (_googleService == null || !_googleService.IsConnected) return;
            try
            {
                var sheets = _googleService.GetSheetNames();
                int autoIdx = -1;
                for (int i = 0; i < sheets.Count; i++)
                {
                    cboGoogleSheet.Items.Add(sheets[i]);
                    if (sheets[i].Equals("Lahore2", StringComparison.OrdinalIgnoreCase)) autoIdx = i;
                }
                if (autoIdx >= 0) cboGoogleSheet.SelectedIndex = autoIdx;
                else if (cboGoogleSheet.Items.Count > 0) cboGoogleSheet.SelectedIndex = 0;
            }
            catch { }
        }

        /// <summary>
        /// 구글 시트 모드에서 입력처리를 수행합니다.
        /// </summary>
        private void BtnInsertGoogle_Click()
        {
            if (_googleService == null || !_googleService.IsConnected)
            {
                MessageBox.Show("먼저 구글 계정에 로그인해주세요.\n(시트 URL 입력 → 로그인 버튼 클릭)", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (cboGoogleSheet.SelectedIndex < 0)
            {
                MessageBox.Show("구글 시트를 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 키워드가 비어있는 행 확인
            var emptyRows = new List<int>();
            for (int i = 0; i < _txtKeywords.Length; i++)
                if (string.IsNullOrWhiteSpace(_txtKeywords[i].Text)) emptyRows.Add(i);

            if (emptyRows.Count > 0)
            {
                var result = MessageBox.Show("키워드 값이 없는 항목이 있습니다.\n해당 행을 삭제하시겠습니까?",
                    "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    RemoveRows(emptyRows);
                    SaveSettings();
                }
                ResetScroll();
            }

            SaveCurrentRecordValues();

            int recordCount = _allRecords.Count;
            List<string[]> recordsToInsert;
            if (recordCount <= 1)
            {
                var singleValues = new string[_txtKeywords.Length];
                for (int i = 0; i < _txtKeywords.Length; i++)
                    singleValues[i] = _txtValues[i].Text.Trim();
                recordsToInsert = new List<string[]> { singleValues };
            }
            else
            {
                var dlgResult = MessageBox.Show($"{recordCount}건을 입력합니다.",
                    "알림", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (dlgResult != DialogResult.OK) return;
                recordsToInsert = _allRecords;
            }

            InsertRecordsGoogle(recordsToInsert);
        }

        /// <summary>
        /// 구글 시트에 레코드를 입력합니다. 로컬 엑셀과 동일한 UNIQUE/INSERT/UPDATE 로직.
        /// </summary>
        private void InsertRecordsGoogle(List<string[]> records)
        {
            string selectedSheet = cboGoogleSheet.SelectedItem?.ToString() ?? "";

            // COL 매핑 확인
            var colMappings = new List<(int rowIdx, int col)>();
            for (int i = 0; i < _txtKeywords.Length; i++)
            {
                string colStr = _txtColNums[i].Text.Trim();
                int colIdx = ColumnLetterToNumber(colStr);
                if (colIdx > 0) colMappings.Add((i, colIdx));
            }
            if (colMappings.Count == 0)
            {
                MessageBox.Show("컬럼(COL)이 지정된 항목이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 데이터가 있는 레코드만 필터
            var validRecords = new List<string[]>();
            foreach (var rec in records)
            {
                bool hasData = colMappings.Any(cm => cm.rowIdx < rec.Length && !string.IsNullOrEmpty(rec[cm.rowIdx]));
                if (hasData) validRecords.Add(rec);
            }
            if (validRecords.Count == 0)
            {
                MessageBox.Show("입력할 데이터가 없습니다.\n먼저 데이터를 입력하고 1차 가공을 해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // UNIQUE 선택 확인
            int uniqueIdx = -1;
            for (int i = 0; i < _rdoUnique.Length; i++)
                if (_rdoUnique[i].Checked) { uniqueIdx = i; break; }

            try
            {
                int maxCol = colMappings.Max(cm => cm.col);
                int lastDataRow = _googleService!.GetLastDataRow(selectedSheet, maxCol);

                // 데이터 유효성 검사 목록 로드 (드롭다운 선택항목 대소문자 매칭용)
                var validationLists = _googleService.GetColumnValidationLists(selectedSheet);

                int insertedCount = 0;
                int updatedCount = 0;
                int firstNewRow = -1;
                var updatedRows = new List<int>();
                var batchWrites = new List<(int row, Dictionary<int, string> colValues)>();

                foreach (var rec in validRecords)
                {
                    if (uniqueIdx >= 0 && uniqueIdx < rec.Length)
                    {
                        string uniqueValue = CleanSpecialChars(rec[uniqueIdx].Trim());
                        int uniqueCol = colMappings.FirstOrDefault(cm => cm.rowIdx == uniqueIdx).col;
                        if (uniqueCol <= 0 || string.IsNullOrEmpty(uniqueValue))
                        {
                            lastDataRow++;
                            var colValues = BuildColValues(rec, colMappings, validationLists);
                            batchWrites.Add((lastDataRow, colValues));
                            if (firstNewRow < 0) firstNewRow = lastDataRow;
                            insertedCount++;
                            continue;
                        }

                        int foundRow = _googleService.FindRowByValue(selectedSheet, uniqueCol, uniqueValue);

                        if (foundRow > 0)
                        {
                            // 기존 행 업데이트 (UNIQUE 컬럼 제외)
                            var colValues = new Dictionary<int, string>();
                            foreach (var (rowIdx, col) in colMappings)
                            {
                                if (rowIdx == uniqueIdx) continue;
                                string val = (rowIdx < rec.Length) ? rec[rowIdx].Trim() : "";
                                if (string.IsNullOrEmpty(val)) continue;
                                val = CleanSpecialChars(val);
                                // 데이터 유효성 검사 목록이 있으면 대소문자 무시 매칭
                                if (validationLists.TryGetValue(col, out var validList))
                                    val = MatchValidationValue(val, validList);
                                colValues[col] = val;
                            }
                            batchWrites.Add((foundRow, colValues));
                            updatedRows.Add(foundRow);
                            updatedCount++;
                        }
                        else
                        {
                            lastDataRow++;
                            var colValues = BuildColValues(rec, colMappings, validationLists);
                            batchWrites.Add((lastDataRow, colValues));
                            if (firstNewRow < 0) firstNewRow = lastDataRow;
                            insertedCount++;
                        }
                    }
                    else
                    {
                        lastDataRow++;
                        var colValues = BuildColValues(rec, colMappings, validationLists);
                        batchWrites.Add((lastDataRow, colValues));
                        if (firstNewRow < 0) firstNewRow = lastDataRow;
                        insertedCount++;
                    }
                }

                // 일괄 기록
                if (batchWrites.Count > 0)
                    _googleService.WriteRows(selectedSheet, batchWrites);

                // 결과 메시지
                var msgParts = new List<string>();
                if (insertedCount > 0)
                    msgParts.Add($"{firstNewRow}행부터 {insertedCount}건이 입력되었습니다.");
                if (updatedCount > 0)
                {
                    string rowList = string.Join(", ", updatedRows.Select(r => $"{r}행"));
                    msgParts.Add($"{rowList} 수정되었습니다.");
                }
                string msg = string.Join("\n", msgParts);
                MessageBox.Show(msg + "\n(구글 시트)", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetAfterInsert();
            }
            catch (Exception ex)
            {
                string msg = ex.Message ?? "";
                if (msg.Contains("403") || msg.Contains("permission", StringComparison.OrdinalIgnoreCase))
                    MessageBox.Show("구글 시트 접근 권한이 없습니다.\n서비스 계정 이메일을 시트에 편집자로 공유했는지 확인해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else if (msg.Contains("404") || msg.Contains("not found", StringComparison.OrdinalIgnoreCase))
                    MessageBox.Show("스프레드시트를 찾을 수 없습니다.\nURL을 확인해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    MessageBox.Show($"구글 시트 저장 오류:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 레코드 값을 컬럼 매핑에 따라 Dictionary로 변환합니다.
        /// </summary>
        private Dictionary<int, string> BuildColValues(string[] rec, List<(int rowIdx, int col)> colMappings, Dictionary<int, List<string>>? validationLists = null)
        {
            var colValues = new Dictionary<int, string>();
            foreach (var (rowIdx, col) in colMappings)
            {
                string val = (rowIdx < rec.Length) ? rec[rowIdx].Trim() : "";
                if (string.IsNullOrEmpty(val)) continue;
                val = CleanSpecialChars(val);
                // 데이터 유효성 검사 목록이 있으면 대소문자 무시 매칭
                if (validationLists != null && validationLists.TryGetValue(col, out var validList))
                    val = MatchValidationValue(val, validList);
                colValues[col] = val;
            }
            return colValues;
        }
    }
}
