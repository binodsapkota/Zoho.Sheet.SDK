using Zoho.Sheet.SDK.Core;
using Zoho.Sheet.SDK.Domain;
using RestSharp;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Zoho.Sheet.SDK.Services
{
    public class ZohoSheetClient : ISheetService
    {
        private readonly IAuthProvider _authProvider;
        private readonly RestClient _client;

        public ZohoSheetClient(IAuthProvider authProvider)
        {
            _authProvider = authProvider;
            _client = new RestClient("https://sheet.zoho.com/api/v2/");
        }

        private async Task<RestRequest> CreateRequestAsync(string resource, Method method, string contentType = "application/json")
        {
            var token = await _authProvider.GetAccessTokenAsync();

            var request = new RestRequest(resource, method);
            request.AddHeader("Authorization", $"Zoho-oauthtoken {token}");
            request.AddHeader("Content-Type", contentType);
            return request;
        }

        // ----------------------------------------------------
        // WORKBOOKS
        // ----------------------------------------------------
        public async Task<List<Workbook>> GetAllWorkbooksAsync(
            int startIndex = 1, int count = 50, string sortOption = "recently_modified")
        {
            var request = await CreateRequestAsync("workbooks", Method.Post, "application/x-www-form-urlencoded");

            request.AddParameter("method", "workbook.list");
            request.AddParameter("start_index", startIndex);
            request.AddParameter("count", count);
            request.AddParameter("sort_option", sortOption);

            var response = await _client.ExecuteAsync(request);
            if (!response.IsSuccessful)
                throw new Exception($"Zoho API Error: {response.StatusCode}: {response.Content}");

            var result = JObject.Parse(response.Content);
            var list = new List<Workbook>();

            foreach (var wb in result["workbooks"])
            {
                list.Add(new Workbook
                {
                    Id = wb["resource_id"].ToString(),
                    Name = wb["workbook_name"].ToString(),
                    Url = wb["workbook_url"].ToString(),
                    CreatedBy = wb["created_by"].ToString(),
                    CreatedTime = wb["created_time"].ToString(),
                    LastModifiedTime = wb["last_modified_time"].ToString()
                });
            }

            return list;
        }

        public async Task<Workbook> CreateWorkbookAsync(string name)
        {
            var token = await _authProvider.GetAccessTokenAsync();

            // Create workbook endpoint is *not* part of /workbooks route
            var client = new RestClient("https://sheet.zoho.com/api/v2/create?method=workbook.create");
            var request = new RestRequest("", Method.Post);

            request.AddHeader("Authorization", $"Zoho-oauthtoken {token}");
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            request.AddParameter("workbook_name", name);

            var response = await client.ExecuteAsync(request);
            if (!response.IsSuccessful)
                throw new Exception($"Error creating workbook: {response.StatusCode} - {response.Content}");

            dynamic json = JsonConvert.DeserializeObject(response.Content);

            return new Workbook
            {
                Id = json.resource_id,
                Name = json.workbook_name,
                Url = json.workbook_url
            };
        }

        public async Task<bool> DeleteWorkbookAsync(string workbookId)
        {
            var request = await CreateRequestAsync($"workbooks/{workbookId}", Method.Delete);
            var response = await _client.ExecuteAsync(request);

            return response.IsSuccessful;
        }

        // ----------------------------------------------------
        // SHEETS
        // ----------------------------------------------------
        public async Task<List<Domain.Sheet>> GetSheetsAsync(string workbookId)
        {
            var request = await CreateRequestAsync($"workbooks/{workbookId}/sheets", Method.Get);
            var response = await _client.ExecuteAsync(request);

            var json = JObject.Parse(response.Content);
            var list = new List<Domain.Sheet>();

            foreach (var sh in json["data"])
            {
                list.Add(new Domain.Sheet
                {
                    Id = sh["id"].ToString(),
                    Name = sh["name"].ToString()
                });
            }

            return list;
        }

        public async Task<Domain.Sheet> CreateSheetAsync(string workbookId, string sheetName)
        {
            var token = await _authProvider.GetAccessTokenAsync();

            // Endpoint requires method appended in URL
            var client = new RestClient("https://sheet.zoho.com/api/v2/");
            var request = new RestRequest($"{workbookId}?method=worksheet.insert", Method.Post);

            request.AddHeader("Authorization", $"Zoho-oauthtoken {token}");
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            // Form data
            request.AddParameter("worksheet_name", sheetName);

            var response = await client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error inserting worksheet: {response.StatusCode} - {response.Content}");

            var json = JObject.Parse(response.Content);

            // Get the newly created worksheet
            var newSheet = json["worksheet_names"]
                           .FirstOrDefault(s => s["worksheet_name"].ToString() == json["new_worksheet_name"].ToString());

            return new Domain.Sheet
            {
                Id = newSheet["worksheet_id"].ToString(),
                Name = newSheet["worksheet_name"].ToString()
            };
        }


        public async Task<bool> DeleteSheetAsync(string workbookId, string sheetName)
        {
            var token = await _authProvider.GetAccessTokenAsync();

            var client = new RestClient("https://sheet.zoho.com/api/v2/");
            var request = new RestRequest($"{workbookId}?method=worksheet.delete", Method.Post);

            request.AddHeader("Authorization", $"Zoho-oauthtoken {token}");
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            request.AddParameter("worksheet_name", sheetName);

            var response = await client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error deleting worksheet: {response.StatusCode} - {response.Content}");

            return true;
        }


        // ----------------------------------------------------
        // COLUMNS
        // ----------------------------------------------------
        public async Task<bool> AddColumnsAsync(string workbookId, string sheetId, List<string> columns)
        {
            var request = await CreateRequestAsync(
                $"workbooks/{workbookId}/sheets/{sheetId}/columns", Method.Post);

            request.AddJsonBody(new { columns });

            var response = await _client.ExecuteAsync(request);
            return response.IsSuccessful;
        }

        public async Task<bool> RemoveColumnsAsync(string workbookId, string sheetId, List<int> columnIndexes)
        {
            var request = await CreateRequestAsync(
                $"workbooks/{workbookId}/sheets/{sheetId}/columns", Method.Delete);

            request.AddJsonBody(new { columns = columnIndexes });

            var response = await _client.ExecuteAsync(request);
            return response.IsSuccessful;
        }

        // ----------------------------------------------------
        // ROWS
        // ----------------------------------------------------
        public async Task<bool> AddRowAsync(string workbookId, string sheetId, List<object> rowData)
        {
            var request = await CreateRequestAsync(
                $"workbooks/{workbookId}/sheets/{sheetId}/rows", Method.Post);

            request.AddJsonBody(new
            {
                data = new List<object[]> { rowData.ToArray() }
            });

            var response = await _client.ExecuteAsync(request);
            return response.IsSuccessful;
        }

        public async Task<bool> UpdateCellAsync(string workbookId, string sheetId, int row, int col, object value)
        {
            var request = await CreateRequestAsync(
                $"workbooks/{workbookId}/sheets/{sheetId}/rows", Method.Patch);

            request.AddJsonBody(new
            {
                data = new[]
                {
                    new { row, col, value }
                }
            });

            var response = await _client.ExecuteAsync(request);
            return response.IsSuccessful;
        }

        public async Task<bool> DeleteRowAsync(string workbookId, string sheetId, int rowIndex)
        {
            var request = await CreateRequestAsync(
                $"workbooks/{workbookId}/sheets/{sheetId}/rows/{rowIndex}", Method.Delete);

            var response = await _client.ExecuteAsync(request);
            return response.IsSuccessful;
        }

        public async Task<bool> RenameSheetAsync(string workbookId, string oldName, string newName)
        {
            // The `method` query parameter MUST be in the URL
            string resource = $"{workbookId}?method=worksheet.rename";

            // Create authorized request
            var token = await _authProvider.GetAccessTokenAsync();
            var client = new RestClient("https://sheet.zoho.com/api/v2/");
            var request = new RestRequest(resource, Method.Post);

            request.AddHeader("Authorization", $"Zoho-oauthtoken {token}");
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            // Form data
            request.AddParameter("old_name", oldName);
            request.AddParameter("new_name", newName);

            var response = await client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error renaming sheet: {response.StatusCode} - {response.Content}");

            return true;
        }



        #region table functions
        /// <summary>
        /// Lists all tables in the given workbook/sheet (resourceId)
        /// </summary>
        public async Task<List<Table>> ListTablesAsync(string resourceId)
        {
            var request = await CreateRequestAsync($"{resourceId}?method=table.list", Method.Post);

            // No data required
            request.AddParameter("", "");

            var response = await _client.ExecuteAsync(request);
            if (!response.IsSuccessful)
                throw new Exception($"Error listing tables: {response.Content}");

            var json = JObject.Parse(response.Content);
            var tables = new List<Table>();

            if (json["tables"] != null)
            {
                foreach (var t in json["tables"])
                {
                    tables.Add(new Table
                    {
                        TableId = t["table_id"].ToObject<int>(),
                        TableName = t["table_name"].ToString(),
                        StartRow = t["start_row"].ToObject<int>(),
                        StartColumn = t["start_column"].ToObject<int>(),
                        EndRow = t["end_row"].ToObject<int>(),
                        EndColumn = t["end_column"].ToObject<int>()
                    });
                }
            }

            return tables;
        }
        public async Task<Table> CreateTableAsync(
                    string resourceId,
                    string worksheetName,
                    int startRow,
                    int startColumn,
                    int endRow,
                    int endColumn,
                    bool containsHeader,
                    List<string> headerNames,
                    object tableStyle)
        {
            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.create", Method.Post, "application/x-www-form-urlencoded");

            // Serialize tableStyle object
            string tableStyleJson = JsonConvert.SerializeObject(tableStyle);

            // Add POST parameters
            request.AddParameter("worksheet_name", worksheetName);
            request.AddParameter("start_row", startRow);
            request.AddParameter("start_column", startColumn);
            request.AddParameter("end_row", endRow);
            request.AddParameter("end_column", endColumn);
            request.AddParameter("contains_header", containsHeader.ToString().ToLower());
            request.AddParameter("header_names", JsonConvert.SerializeObject(headerNames));
            request.AddParameter("table_style", tableStyleJson);

            // Execute request
            var response = await _client.ExecuteAsync(request);
            if (!response.IsSuccessful)
                throw new Exception($"Error creating table: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = JObject.Parse(response.Content);

            return new Table
            {
                TableName = json["table_name"].ToString(),
                TableId = json["table_id"].ToObject<int>(),
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = endRow,
                EndColumn = endColumn
            };
        }

        public async Task<bool> DeleteTableAsync(string resourceId, string tableName, bool clearFormat = true)
        {
            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.remove", Method.Post, "application/x-www-form-urlencoded");

            // Add POST parameters
            request.AddParameter("table_name", tableName);
            request.AddParameter("clear_format", clearFormat.ToString().ToLower());

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error deleting table: {response.StatusCode} - {response.Content}");

            // Optionally, you can parse and verify status
            var json = Newtonsoft.Json.Linq.JObject.Parse(response.Content);
            return json["status"] != null && json["status"].ToString().ToLower() == "success";
        }

        public async Task<bool> RenameTableHeadersAsync(
                string resourceId,
                string tableName,
                List<TableHeaderRename> headers)
        {
            if (headers == null || headers.Count == 0)
                throw new ArgumentException("Headers list cannot be null or empty.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.header.rename", Method.Post, "application/x-www-form-urlencoded");

            // Serialize header rename list
            string jsonData = JsonConvert.SerializeObject(headers);

            // Add parameters
            request.AddParameter("table_name", tableName);
            request.AddParameter("data", jsonData);

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error renaming table headers: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = JObject.Parse(response.Content);
            return json["status"] != null && json["status"].ToString().ToLower() == "success";
        }



        public async Task<List<TableRecord>> FetchTableRecordsAsync(
            string resourceId,
            string tableName,
            List<TableRecordCriteria> criteria,
            string criteriaPattern,
            List<string> columnNames,
            string renderOption = "formatted",
            int count = 50,
            bool isCaseSensitive = true)
        {
            if (criteria == null) criteria = new List<TableRecordCriteria>();
            if (columnNames == null) columnNames = new List<string>();

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.records.fetch", Method.Post, "application/x-www-form-urlencoded");

            // Add parameters
            request.AddParameter("table_name", tableName);
            request.AddParameter("criteria_json", JsonConvert.SerializeObject(criteria));
            request.AddParameter("criteria_pattern", criteriaPattern);
            request.AddParameter("column_names", string.Join(",", columnNames));
            request.AddParameter("render_option", renderOption);
            request.AddParameter("count", count);
            request.AddParameter("is_case_sensitive", isCaseSensitive.ToString().ToLower());

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error fetching table records: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = JObject.Parse(response.Content);
            var records = new List<TableRecord>();

            if (json["records"] != null)
            {
                foreach (var rec in json["records"])
                {
                    var record = new TableRecord
                    {
                        RowIndex = rec["row_index"].ToObject<int>()
                    };

                    foreach (var prop in rec.Children<JProperty>())
                    {
                        if (prop.Name != "row_index")
                            record.Data[prop.Name] = prop.Value.ToObject<object>();
                    }

                    records.Add(record);
                }
            }

            return records;
        }
        public async Task<int> UpdateTableRecordsAsync(
    string resourceId,
    string tableName,
    List<TableRecordCriteria> criteria,
    string criteriaPattern,
    Dictionary<string, object> newData,
    bool isCaseSensitive = true)
        {
            if (criteria == null || criteria.Count == 0)
                throw new ArgumentException("Criteria list cannot be null or empty.");

            if (newData == null || newData.Count == 0)
                throw new ArgumentException("New data cannot be null or empty.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.records.update", Method.Post, "application/x-www-form-urlencoded");

            // Serialize criteria and data
            string criteriaJson = JsonConvert.SerializeObject(criteria);
            string dataJson = JsonConvert.SerializeObject(newData);

            // Add parameters
            request.AddParameter("table_name", tableName);
            request.AddParameter("criteria_json", criteriaJson);
            request.AddParameter("criteria_pattern", criteriaPattern);
            request.AddParameter("is_case_sensitive", isCaseSensitive.ToString().ToLower());
            request.AddParameter("data", dataJson);

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error updating table records: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = JObject.Parse(response.Content);
            return json["no_of_affected_rows"]?.ToObject<int>() ?? 0;
        }

        public async Task<(int DeletedRows, int RemainingRows)> DeleteTableRecordsAsync(
    string resourceId,
    string tableName,
    List<TableRecordCriteria> criteria,
    string criteriaPattern,
    bool isCaseSensitive = true)
        {
            if (criteria == null || criteria.Count == 0)
                throw new ArgumentException("Criteria list cannot be null or empty.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.records.delete", Method.Post, "application/x-www-form-urlencoded");

            // Serialize criteria
            string criteriaJson = JsonConvert.SerializeObject(criteria);

            // Add parameters
            request.AddParameter("table_name", tableName);
            request.AddParameter("criteria_json", criteriaJson);
            request.AddParameter("criteria_pattern", criteriaPattern);
            request.AddParameter("is_case_sensitive", isCaseSensitive.ToString().ToLower());

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error deleting table records: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = JObject.Parse(response.Content);
            int deleted = json["no_of_rows_deleted"]?.ToObject<int>() ?? 0;
            int remaining = json["no_of_rows_remaining"]?.ToObject<int>() ?? 0;

            return (deleted, remaining);
        }

        public async Task<bool> InsertTableColumnsAsync(
    string resourceId,
    string tableName,
    List<string> columnNames,
    string insertAfterColumn)
        {
            if (columnNames == null || columnNames.Count == 0)
                throw new ArgumentException("Column names list cannot be null or empty.");

            if (string.IsNullOrEmpty(insertAfterColumn))
                throw new ArgumentException("Insert after column must be provided.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.columns.insert", Method.Post, "application/x-www-form-urlencoded");

            // Add parameters
            request.AddParameter("table_name", tableName);
            request.AddParameter("column_names", JsonConvert.SerializeObject(columnNames));
            request.AddParameter("insert_column_after", insertAfterColumn);

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error inserting table columns: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = Newtonsoft.Json.Linq.JObject.Parse(response.Content);
            return json["status"]?.ToString().ToLower() == "success";
        }
        public async Task<bool> DeleteTableColumnsAsync(
            string resourceId,
            string tableName,
            List<string> columnNames)
        {
            if (columnNames == null || columnNames.Count == 0)
                throw new ArgumentException("Column names list cannot be null or empty.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=table.columns.delete", Method.Post, "application/x-www-form-urlencoded");

            // Add parameters
            request.AddParameter("table_name", tableName);
            request.AddParameter("column_names", JsonConvert.SerializeObject(columnNames));

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error deleting table columns: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = Newtonsoft.Json.Linq.JObject.Parse(response.Content);
            return json["status"]?.ToString().ToLower() == "success";
        }
        #endregion
        #region worksheet Data Fetch

        public async Task<List<Dictionary<string, object>>> FetchWorksheetRecordsAsync(
    string resourceId,
    string worksheetName,
    string criteria,
    List<string> columnNames,
    int recordsStartIndex = 1,
    int count = 50,
    bool isCaseSensitive = true,
    string renderOption = "formatted",
    int headerRow = 1)
        {
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name must be provided.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=worksheet.records.fetch", Method.Post, "application/x-www-form-urlencoded");

            // Add parameters
            request.AddParameter("worksheet_name", worksheetName);
            request.AddParameter("header_row", headerRow);
            //request.AddParameter("criteria", criteria);
            request.AddParameter("column_names", string.Join(",", columnNames));
            request.AddParameter("render_option", renderOption);
            request.AddParameter("records_start_index", recordsStartIndex);
            request.AddParameter("count", count);
            request.AddParameter("is_case_sensitive", isCaseSensitive.ToString().ToLower());

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error fetching worksheet records: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = JObject.Parse(response.Content);
            var records = new List<Dictionary<string, object>>();

            if (json["records"] != null)
            {
                foreach (var rec in json["records"])
                {
                    var dict = new Dictionary<string, object>();
                    foreach (var prop in rec.Children<JProperty>())
                    {
                        dict[prop.Name] = prop.Value.ToObject<object>();
                    }
                    records.Add(dict);
                }
            }

            return records;
        }

        public async Task<bool> AddWorksheetRecordsAsync(
            string resourceId,
            string worksheetName,
            List<Dictionary<string, object>> records,
            int headerRow = 1)
        {
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name must be provided.");

            if (records == null || records.Count == 0)
                throw new ArgumentException("Records list cannot be null or empty.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=worksheet.records.add", Method.Post, "application/x-www-form-urlencoded");

            // Add parameters
            request.AddParameter("worksheet_name", worksheetName);
            request.AddParameter("header_row", headerRow);
            request.AddParameter("json_data", JsonConvert.SerializeObject(records));

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error adding worksheet records: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = JObject.Parse(response.Content);
            return json["status"]?.ToString().ToLower() == "success";
        }


        public async Task<int> UpdateWorksheetRecordsAsync(
            string resourceId,
            string worksheetName,
            string criteria,
            Dictionary<string, object> data,
            bool isCaseSensitive = true,
    int headerRow = 1)
        {
            if (string.IsNullOrEmpty(resourceId)) throw new ArgumentNullException(nameof(resourceId));
            if (string.IsNullOrEmpty(worksheetName)) throw new ArgumentNullException(nameof(worksheetName));
            if (data == null || data.Count == 0) throw new ArgumentNullException(nameof(data));

            var request = await CreateRequestAsync($"{resourceId}?method=worksheet.records.update", Method.Post);

            // Form data as x-www-form-urlencoded
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");

            request.AddParameter("worksheet_name", worksheetName);
            request.AddParameter("header_row", headerRow);
            request.AddParameter("criteria", criteria);
            request.AddParameter("is_case_sensitive", isCaseSensitive.ToString().ToLower());
            request.AddParameter("data", JsonConvert.SerializeObject(data));

            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error updating worksheet records: {response.StatusCode} - {response.Content}");

            dynamic json = JsonConvert.DeserializeObject(response.Content);

            return json.no_of_affected_rows != null ? (int)json.no_of_affected_rows : 0;
        }

        public async Task<(int deletedRows, int remainingRows)> DeleteWorksheetRecordsAsync(
            string resourceId,
            string worksheetName,
            string criteria = null,
            List<int> rowArray = null,
            int headerRow = 1,
            bool deleteRows = true)
        {
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name must be provided.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=worksheet.records.delete", Method.Post, "application/x-www-form-urlencoded");

            // Add parameters
            request.AddParameter("worksheet_name", worksheetName);
            request.AddParameter("header_row", headerRow);
            if (!string.IsNullOrEmpty(criteria))
                request.AddParameter("criteria", criteria);
            if (rowArray != null && rowArray.Count > 0)
                request.AddParameter("row_array", JsonConvert.SerializeObject(rowArray));
            request.AddParameter("delete_rows", deleteRows.ToString().ToLower());

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error deleting worksheet records: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = Newtonsoft.Json.Linq.JObject.Parse(response.Content);
            int deleted = json["no_of_rows_deleted"] != null ? (int)json["no_of_rows_deleted"] : 0;
            int remaining = json["no_of_rows_remaining"] != null ? (int)json["no_of_rows_remaining"] : 0;

            return (deleted, remaining);
        }
        public async Task<bool> InsertWorksheetColumnsAsync(
    string resourceId,
    string worksheetName,
    string insertAfterColumn,
    List<string> columnNames)
        {
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name must be provided.");

            if (string.IsNullOrEmpty(insertAfterColumn))
                throw new ArgumentException("Insert after column must be provided.");

            if (columnNames == null || columnNames.Count == 0)
                throw new ArgumentException("Column names cannot be null or empty.");

            // Create request
            var request = await CreateRequestAsync($"{resourceId}?method=records.columns.insert", Method.Post, "application/x-www-form-urlencoded");

            // Add parameters
            request.AddParameter("worksheet_name", worksheetName);
            request.AddParameter("insert_column_after", insertAfterColumn);
            request.AddParameter("column_names", JsonConvert.SerializeObject(columnNames));

            // Execute request
            var response = await _client.ExecuteAsync(request);

            if (!response.IsSuccessful)
                throw new Exception($"Error inserting worksheet columns: {response.StatusCode} - {response.Content}");

            // Parse response
            var json = Newtonsoft.Json.Linq.JObject.Parse(response.Content);
            return json["status"] != null && json["status"].ToString().ToLower() == "success";
        }
        #endregion

    }
    public class TableHeaderRename
    {
        public string OldName { get; set; }
        public string NewName { get; set; }
    }
}
