using Microsoft.Extensions.Configuration;
using Zoho.Sheet.SDK.Domain;
using Zoho.Sheet.SDK.Services.Auth;
using Zoho.Sheet.SDK.Services;

// 1. Load configuration from appsettings.json
var configuration = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .Build();

// 2. Bind ZohoConfig section
var zohoConfig = new ZohoConfig
{
    ClientId = configuration["Zoho:ClientId"],
    ClientSecret = configuration["Zoho:ClientSecret"],
    AccessToken = configuration["Zoho:AccessToken"],
    RefreshToken = configuration["Zoho:RefreshToken"],
    DataCenter = configuration["Zoho:DataCenter"]
};

// 3. Create AuthProvider and SheetClient
var authProvider = new ZohoAuthProvider(zohoConfig);
var sheetClient = new ZohoSheetClient(authProvider);
// Step 1: Get device code if needed
if (string.IsNullOrEmpty(zohoConfig.RefreshToken))
{
    var deviceCodeInfo = await authProvider.GetDeviceCodeAsync();
    zohoConfig.DeviceCode = deviceCodeInfo.DeviceCode;
    Console.WriteLine($"User Code: {deviceCodeInfo.UserCode}");
    Console.WriteLine($"Verification URL: {deviceCodeInfo.VerificationUrl}");
    Console.WriteLine("Please authorize the device in your browser. And Press Enter");
    Console.ReadLine();
}
// 4. Use SDK
var workbooks = await sheetClient.GetAllWorkbooksAsync();
string resourceId= "";
foreach (var wb in workbooks)
{
    Console.WriteLine($"{wb.Name} ({wb.Id}) - Created by {wb.CreatedBy}");
    resourceId = wb.Id; // just store last one for testing
}

// Store IDs
string workbookId = "";
string worksheetName = "TestSheet1";
TestWorksheetRecordsAsync(resourceId).Wait();

try
{
    // ----------------------------------------------------
    // 1. CREATE WORKBOOK
    // ----------------------------------------------------
    Console.WriteLine("Creating workbook...");
    var workbook = await sheetClient.CreateWorkbookAsync("MyTestWorkbook");
    workbookId = workbook.Id;
    Console.WriteLine($"Created workbook: {workbook.Name}, ID: {workbookId}");

    // ----------------------------------------------------
    // 2. CREATE WORKSHEET
    // ----------------------------------------------------
    Console.WriteLine("Creating worksheet...");
    var sheet = await sheetClient.CreateSheetAsync(workbookId, worksheetName);
    Console.WriteLine($"Created worksheet: {sheet.Name}, ID: {sheet.Id}");

    // ----------------------------------------------------
    // 3. RENAME WORKSHEET
    // ----------------------------------------------------
    Console.WriteLine("Renaming worksheet...");
    var newSheetName = "RenamedSheet";

    bool renameResult = await sheetClient.RenameSheetAsync(workbookId, worksheetName, newSheetName);
    if (renameResult)
        Console.WriteLine($"Worksheet renamed to: {newSheetName}");
    else
        Console.WriteLine("Worksheet rename failed!");

    worksheetName = newSheetName; // update reference

    // ----------------------------------------------------
    // 4. DELETE WORKSHEET
    // ----------------------------------------------------
    Console.WriteLine("Deleting worksheet...");
    bool deleteSheet = await sheetClient.DeleteSheetAsync(workbookId, worksheetName);
    Console.WriteLine(deleteSheet ? "Worksheet deleted." : "Worksheet deletion failed!");

    // ----------------------------------------------------
    // 5. DELETE WORKBOOK
    // ----------------------------------------------------
    Console.WriteLine("Deleting workbook...");
    bool deleted = await sheetClient.DeleteWorkbookAsync(workbookId);
    Console.WriteLine(deleted ? "Workbook deleted." : "Workbook deletion failed!");
}
catch (Exception ex)
{
    Console.WriteLine("ERROR:");
    Console.WriteLine(ex.Message);
}

Console.WriteLine("----- ZOHO SHEET TEST COMPLETE -----");

async Task TestWorksheetRecordsAsync(string resourceId)
{
    try
    {
        // your workbook/resource id
        var worksheetName = "Sheet1";

        // ----------------------
        // 1. FETCH Records
        // ----------------------
        var fetchedRecords = await sheetClient.FetchWorksheetRecordsAsync(
            resourceId,
            worksheetName,
            headerRow: 1,
            criteria: "(\"Month\"=\"March\" OR \"Month\"=\"April\") AND \"Amount\">30",
            columnNames: new List<string> { "Month", "Amount" },
            renderOption: "formatted",
            recordsStartIndex: 1,
            count: 10,
            isCaseSensitive: true
        );

        Console.WriteLine($"Fetched {fetchedRecords.Count} records");

        // ----------------------
        // 2. ADD Records
        // ----------------------
        var newRecords = new List<Dictionary<string, object>>
        {
            new Dictionary<string, object> { { "Name", "Joe" }, { "Region", "South" }, { "Units", 284 } },
            new Dictionary<string, object> { { "Name", "Beth" }, { "Region", "East" }, { "Units", 290 } }
        };

        bool addSuccess = await sheetClient.AddWorksheetRecordsAsync(resourceId, worksheetName, newRecords);
        Console.WriteLine($"Add Records Success: {addSuccess}");

        // ----------------------
        // 3. UPDATE Records
        // ----------------------
        var updateData = new Dictionary<string, object> { { "Amount", 50 } };
        var updatedRows = await sheetClient.UpdateWorksheetRecordsAsync(
            resourceId: "aaaaabbbbbccccdddddeeee",
            worksheetName: "Sheet1",
            criteria: "\"Month\"=\"March\"",
            data: new Dictionary<string, object>
            {
                { "Month", "May" },
                { "Amount", 50 }
            }
        );
        Console.WriteLine($"Update Records Success: {updatedRows}");

        // ----------------------
        // 4. DELETE Records
        // ----------------------
        var deleteSuccess = await sheetClient.DeleteWorksheetRecordsAsync(
            resourceId,
            worksheetName,
            criteria: "\"Month\"=\"March\" AND \"Amount\">50",
            rowArray: new List<int> { 1, 2, 3, 4, 5 },
            deleteRows: true
        );
        Console.WriteLine($"Delete Records Success: {deleteSuccess}");

        // ----------------------
        // 5. INSERT Columns
        // ----------------------
        var insertColumns = new List<string> { "Phone", "Email" };
        bool insertSuccess = await sheetClient.InsertWorksheetColumnsAsync(
            resourceId,
            worksheetName,
            insertAfterColumn: "Region",
            columnNames: insertColumns
        );
        Console.WriteLine($"Insert Columns Success: {insertSuccess}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
    }
}
