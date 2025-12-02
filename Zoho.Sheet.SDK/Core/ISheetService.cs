using System.Collections.Generic;
using System.Threading.Tasks;
using Zoho.Sheet.SDK.Domain;

namespace Zoho.Sheet.SDK.Core
{
    public interface ISheetService
    {
        // Workbooks
        Task<List<Workbook>> GetAllWorkbooksAsync(
            int startIndex = 1, int count = 50, string sortOption = "recently_modified");
        Task<Workbook> CreateWorkbookAsync(string name);
        Task<bool> DeleteWorkbookAsync(string workbookId);

        // Sheets
        Task<List<Domain.Sheet>> GetSheetsAsync(string workbookId);
        Task<Domain.Sheet> CreateSheetAsync(string workbookId, string sheetName);
        Task<bool> RenameSheetAsync(string workbookId, string oldName, string newName);
        Task<bool> DeleteSheetAsync(string workbookId, string sheetId);

        // Columns
        Task<bool> AddColumnsAsync(string workbookId, string sheetId, List<string> columns);
        Task<bool> RemoveColumnsAsync(string workbookId, string sheetId, List<int> columnIndexes);

        // Rows
        Task<bool> AddRowAsync(string workbookId, string sheetId, List<object> rowData);
        Task<bool> UpdateCellAsync(string workbookId, string sheetId, int row, int col, object value);
        Task<bool> DeleteRowAsync(string workbookId, string sheetId, int rowIndex);
    }
}
