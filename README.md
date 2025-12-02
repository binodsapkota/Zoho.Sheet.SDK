# Zoho.Sheet.SDK

Zoho.Sheet.SDK is a .NET SDK that provides a strongly-typed and developer-friendly interface to interact with Zoho Sheet APIs. It allows you to programmatically manage workbooks, worksheets, tables, columns, and records without manually handling HTTP requests.

Inspired by the Smartsheet SDK, this SDK provides a familiar experience for .NET developers to automate spreadsheets, integrate data, and build robust applications.

---

## Features

* Workbooks & Worksheets

  * Create, list, update, delete, and rename.
* Tables

  * Create, list, remove, and manage headers.
* Records

  * Add, update, fetch, and delete records in worksheets or tables.
* Columns

  * Insert, delete, or rename columns.
* Generic Mapping

  * Map sheet records to custom objects (TaskItem, Customer, etc.) for type-safe operations.
* OAuth2 Authentication

  * Supports Device Code and Refresh Token flows for secure API access.

---

## Installation

Install via NuGet:

```
Install-Package Zoho.Sheet.SDK
```

Or via .NET CLI:

```
dotnet add package Zoho.Sheet.SDK
```

---

## Configuration

Add your Zoho credentials in appsettings.json:

```
{
  "Zoho": {
    "ClientId": "1000.xxxxxxxx",
    "ClientSecret": "xxxxxxxxxxxxxxxx",
    "DeviceCode": "1000.xxxxxxxx",
    "DataCenter": "com"
  }
}
```

---

## Usage

### Initialize the SDK

```
using Zoho.Sheet.SDK.Services;
using Zoho.Sheet.SDK.Services.Auth;

// Initialize AuthProvider
var authProvider = new ZohoAuthProvider(
    clientId: "<CLIENT_ID>",
    clientSecret: "<CLIENT_SECRET>",
    deviceCode: "<DEVICE_CODE>"
);

// Initialize Sheet Client
var sheetClient = new ZohoSheetClient(authProvider);
```

### Workbooks

```
var workbooks = await sheetClient.GetAllWorkbooksAsync();
var workbook = await sheetClient.CreateWorkbookAsync("ProjectPlan");
bool deleted = await sheetClient.DeleteWorkbookAsync(workbook.Id);
```

### Worksheets

```
var sheet = await sheetClient.CreateSheetAsync(workbook.Id, "Tasks");
bool sheetDeleted = await sheetClient.DeleteSheetAsync(workbook.Id, sheet.Id);
```

### Records

```
var records = new List<object[]> {
    new object[] { "Task1", "John", true },
    new object[] { "Task2", "Alice", false }
};
await sheetClient.AddRowAsync(workbook.Id, sheet.Id, records[0]);

var fetched = await sheetClient.GetRecordsAsync(workbook.Id, sheet.Id, "Task");

await sheetClient.UpdateCellAsync(workbook.Id, sheet.Id, 1, 3, true);

await sheetClient.DeleteRowAsync(workbook.Id, sheet.Id, 2);
```

---

## Contributing

Contributions are welcome! Feel free to open issues or submit pull requests. Follow standard .NET coding conventions and write unit tests for new features.

---

## License

This project is licensed under the MIT License.

---

## Acknowledgements

* Inspired by the Smartsheet SDK
* Zoho Sheet API Documentation: [https://www.zoho.com/sheet/api/](https://www.zoho.com/sheet/api/)

