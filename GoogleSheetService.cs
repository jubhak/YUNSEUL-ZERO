using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace DataParser
{
    /// <summary>
    /// Google Sheets API를 사용하여 스프레드시트에 데이터를 읽고 쓰는 서비스 클래스입니다.
    /// OAuth 2.0 사용자 로그인 방식을 사용합니다.
    /// </summary>
    public class GoogleSheetService : IDisposable
    {
        private SheetsService? _service;
        private string? _spreadsheetId;
        private static readonly string TokenFolder = Path.Combine(AppContext.BaseDirectory, "google_token");

        public bool IsConnected => _service != null && !string.IsNullOrEmpty(_spreadsheetId);
        public string? UserEmail { get; private set; }

        /// <summary>
        /// OAuth 2.0으로 사용자 구글 계정에 로그인합니다.
        /// 브라우저가 열리고 사용자가 로그인하면 토큰이 로컬에 저장됩니다.
        /// </summary>
        /// <param name="clientId">OAuth Client ID</param>
        /// <param name="clientSecret">OAuth Client Secret</param>
        public async Task<bool> LoginAsync(string clientId, string clientSecret)
        {
            var clientSecrets = new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            };

            var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                clientSecrets,
                new[] { SheetsService.Scope.Spreadsheets },
                "user",
                CancellationToken.None,
                new FileDataStore(TokenFolder, true));

            _service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "YUNSEUL-ZERO"
            });

            // 이메일 정보 가져오기 (토큰에서)
            UserEmail = "로그인됨";
            return true;
        }

        /// <summary>
        /// 이미 저장된 토큰으로 자동 로그인을 시도합니다.
        /// </summary>
        public async Task<bool> TryAutoLoginAsync(string clientId, string clientSecret)
        {
            if (!Directory.Exists(TokenFolder)) return false;
            var tokenFile = Path.Combine(TokenFolder, "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            if (!File.Exists(tokenFile)) return false;

            try
            {
                return await LoginAsync(clientId, clientSecret);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 저장된 토큰을 삭제하여 로그아웃합니다.
        /// </summary>
        public void Logout()
        {
            _service?.Dispose();
            _service = null;
            _spreadsheetId = null;
            UserEmail = null;
            if (Directory.Exists(TokenFolder))
            {
                try { Directory.Delete(TokenFolder, true); } catch { }
            }
        }

        /// <summary>
        /// 스프레드시트 URL을 설정합니다.
        /// </summary>
        public void SetSpreadsheet(string spreadsheetUrl)
        {
            _spreadsheetId = ExtractSpreadsheetId(spreadsheetUrl);
            if (string.IsNullOrEmpty(_spreadsheetId))
                throw new ArgumentException("유효한 구글 스프레드시트 URL 또는 ID가 아닙니다.");
        }

        /// <summary>
        /// 스프레드시트의 시트 이름 목록을 반환합니다.
        /// </summary>
        public List<string> GetSheetNames()
        {
            if (_service == null || string.IsNullOrEmpty(_spreadsheetId))
                return new List<string>();

            var request = _service.Spreadsheets.Get(_spreadsheetId);
            var spreadsheet = request.Execute();
            return spreadsheet.Sheets.Select(s => s.Properties.Title).ToList();
        }

        /// <summary>
        /// 지정된 시트에서 데이터가 있는 마지막 행 번호를 반환합니다 (1-based).
        /// </summary>
        public int GetLastDataRow(string sheetName, int maxCol)
        {
            if (_service == null || string.IsNullOrEmpty(_spreadsheetId))
                return 0;

            string colLetter = ColumnNumberToLetter(maxCol);
            string range = $"{sheetName}!A1:{colLetter}";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = request.Execute();

            if (response.Values == null || response.Values.Count == 0)
                return 0;

            int lastRow = 0;
            for (int i = 0; i < response.Values.Count; i++)
            {
                var row = response.Values[i];
                if (row != null && row.Any(cell => cell != null && !string.IsNullOrWhiteSpace(cell.ToString())))
                    lastRow = i + 1;
            }
            return lastRow;
        }

        /// <summary>
        /// 지정된 시트의 특정 컬럼에서 값을 검색하여 행 번호를 반환합니다 (1-based). 없으면 -1.
        /// </summary>
        public int FindRowByValue(string sheetName, int col, string searchValue)
        {
            if (_service == null || string.IsNullOrEmpty(_spreadsheetId))
                return -1;

            string colLetter = ColumnNumberToLetter(col);
            string range = $"{sheetName}!{colLetter}:{colLetter}";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = request.Execute();

            if (response.Values == null) return -1;

            for (int i = 0; i < response.Values.Count; i++)
            {
                var row = response.Values[i];
                if (row != null && row.Count > 0)
                {
                    string cellVal = row[0]?.ToString()?.Trim() ?? "";
                    if (cellVal.Equals(searchValue, StringComparison.OrdinalIgnoreCase))
                        return i + 1;
                }
            }
            return -1;
        }

        /// <summary>
        /// 여러 행에 데이터를 일괄 기록합니다 (효율적인 배치 처리).
        /// </summary>
        public void WriteRows(string sheetName, List<(int row, Dictionary<int, string> colValues)> rowsData)
        {
            if (_service == null || string.IsNullOrEmpty(_spreadsheetId))
                return;

            var requests = new List<Request>();
            int sheetId = GetSheetId(sheetName);

            foreach (var (row, colValues) in rowsData)
            {
                foreach (var (col, value) in colValues)
                {
                    if (string.IsNullOrEmpty(value)) continue;

                    requests.Add(new Request
                    {
                        UpdateCells = new UpdateCellsRequest
                        {
                            Start = new GridCoordinate
                            {
                                SheetId = sheetId,
                                RowIndex = row - 1,
                                ColumnIndex = col - 1
                            },
                            Rows = new List<RowData>
                            {
                                new RowData
                                {
                                    Values = new List<CellData>
                                    {
                                        new CellData
                                        {
                                            UserEnteredValue = new ExtendedValue { StringValue = value }
                                        }
                                    }
                                }
                            },
                            Fields = "userEnteredValue"
                        }
                    });
                }
            }

            if (requests.Count > 0)
            {
                var batchUpdate = new BatchUpdateSpreadsheetRequest { Requests = requests };
                _service.Spreadsheets.BatchUpdate(batchUpdate, _spreadsheetId).Execute();
            }
        }

        /// <summary>
        /// 시트의 데이터 유효성 검사 규칙(드롭다운 선택항목)을 컬럼별로 가져옵니다.
        /// 반환: Dictionary&lt;col(1-based), List&lt;string&gt; validValues&gt;
        /// </summary>
        public Dictionary<int, List<string>> GetColumnValidationLists(string sheetName)
        {
            var result = new Dictionary<int, List<string>>();
            if (_service == null || string.IsNullOrEmpty(_spreadsheetId))
                return result;

            try
            {
                var request = _service.Spreadsheets.Get(_spreadsheetId);
                request.IncludeGridData = false;
                var spreadsheet = request.Execute();

                var sheet = spreadsheet.Sheets.FirstOrDefault(s =>
                    s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet?.BasicFilter != null) { /* no-op */ }

                // 데이터 유효성 규칙은 시트 메타데이터에서 가져옴
                // Sheets API v4에서는 개별 셀의 dataValidation을 가져오려면 includeGridData가 필요
                var dataRequest = _service.Spreadsheets.Get(_spreadsheetId);
                dataRequest.IncludeGridData = true;
                dataRequest.Ranges = new[] { $"{sheetName}!1:1" }; // 첫 행만 읽어서 열 수 파악

                var fullSpreadsheet = dataRequest.Execute();
                var targetSheet = fullSpreadsheet.Sheets.FirstOrDefault(s =>
                    s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

                if (targetSheet?.Data == null || targetSheet.Data.Count == 0)
                    return result;

                // 전체 시트의 conditionalFormats/dataValidation을 가져오기 위해 더 넓은 범위 요청
                int colCount = targetSheet.Properties.GridProperties?.ColumnCount ?? 26;
                string colLetter = ColumnNumberToLetter(colCount);
                var fullRequest = _service.Spreadsheets.Get(_spreadsheetId);
                fullRequest.IncludeGridData = true;
                fullRequest.Ranges = new[] { $"{sheetName}!A1:{colLetter}2" }; // 첫 2행으로 유효성 규칙 확인

                var fullSheet = fullRequest.Execute();
                var dataSheet = fullSheet.Sheets.FirstOrDefault(s =>
                    s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

                if (dataSheet?.Data == null || dataSheet.Data.Count == 0)
                    return result;

                var gridData = dataSheet.Data[0];
                if (gridData.RowData == null) return result;

                // 각 열의 데이터 유효성 규칙 수집
                foreach (var rowData in gridData.RowData)
                {
                    if (rowData.Values == null) continue;
                    for (int colIdx = 0; colIdx < rowData.Values.Count; colIdx++)
                    {
                        var cell = rowData.Values[colIdx];
                        if (cell?.DataValidation == null) continue;
                        var dv = cell.DataValidation;
                        if (dv.Condition?.Type == "ONE_OF_LIST" && dv.Condition.Values != null)
                        {
                            int col1Based = colIdx + 1;
                            if (!result.ContainsKey(col1Based))
                            {
                                var validValues = dv.Condition.Values
                                    .Where(v => v.UserEnteredValue != null)
                                    .Select(v => v.UserEnteredValue)
                                    .ToList();
                                if (validValues.Count > 0)
                                    result[col1Based] = validValues;
                            }
                        }
                    }
                }
            }
            catch { /* 유효성 규칙을 가져올 수 없으면 무시 */ }

            return result;
        }

        private int GetSheetId(string sheetName)
        {
            if (_service == null || string.IsNullOrEmpty(_spreadsheetId))
                return 0;

            var spreadsheet = _service.Spreadsheets.Get(_spreadsheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s =>
                s.Properties.Title.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            return sheet?.Properties.SheetId ?? 0;
        }

        /// <summary>
        /// 스프레드시트 URL에서 ID를 추출합니다.
        /// </summary>
        private static string? ExtractSpreadsheetId(string urlOrId)
        {
            if (string.IsNullOrWhiteSpace(urlOrId)) return null;

            var match = Regex.Match(urlOrId, @"/spreadsheets/d/([a-zA-Z0-9_-]+)");
            if (match.Success) return match.Groups[1].Value;

            if (Regex.IsMatch(urlOrId.Trim(), @"^[a-zA-Z0-9_-]{20,}$"))
                return urlOrId.Trim();

            return null;
        }

        private static string ColumnNumberToLetter(int colNumber)
        {
            if (colNumber <= 0) return "A";
            string result = "";
            while (colNumber > 0)
            {
                colNumber--;
                result = (char)('A' + colNumber % 26) + result;
                colNumber /= 26;
            }
            return result;
        }

        public void Dispose()
        {
            _service?.Dispose();
            _service = null;
        }
    }
}
