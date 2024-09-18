# Excel VBA Methods for AutoHelp Sales Automation

This repository contains a series of Visual Basic for Applications (VBA) methods created to streamline and automate tasks in Excel, specifically developed for use in a sales environment at AutoHelp. These scripts improve speed, accuracy, and efficiency in day-to-day operations, reducing the manual workload associated with data entry, reporting, and validation.

## Repository Structure

```plaintext
ExcelVBA/
├── Brands                 # Main automation scripts for all brands in the company
├── Masks                  # Additional helper functions for distribution of goods in relation to stock by store for the entire store network
├──  Pricing               # Pricing of an item according to its brand and tailored to the pricing conditions of a customer.
├── allConditions          # Excel Customizations.exportedUI_09.01.2020.exportedUI
├── Excel Custom...        # Exported customization file for Microsoft Excel's Ribbon UI
├── Headers.bas            # Costumizing the file for user use
├── Module1.bas            # Pulls an ABC analysis of the items and checks if a product is matched in the database.
├── OpenAllWorkbooks...bas # Open all workbooks with a specific name in a given folder
├── README.md              # This file
├── SplitOrders...bas      # Splits data from a specified range into separate Excel worksheets
└── dig2txt.bas            # Dig 2 txt
```
## Features

-   **Automated Data Entry**: Speed up the process of entering sales data into predefined Excel templates.
-   **Sales Report Generation**: Automatically generate comprehensive reports with minimal user input, ready for analysis.
-   **Data Validation**: Methods to ensure the accuracy and consistency of the entered sales data.
-   **Summary Metrics**: Generate quick insights from large datasets, providing metrics like total sales, average transaction value, and more.

## Getting Started

### Requirements

-   Microsoft Excel (2016 or later) with VBA enabled
-   Basic understanding of Excel and how to enable macros

### Installation

1.  Download the `.bas` files from the repository.
2.  Open Excel and press `Alt + F11` to launch the VBA editor.
3.  In the editor, go to `File > Import File` and select the `.bas` files you want to use.
4.  Save your Excel file as a macro-enabled workbook (`.xlsm`).

### How to Use the Methods

#### 1. **AutoFillSalesData()**

-   **Description**: Automatically populates sales data into a template.
-   **How it works**:
    1.  Fetches data from a source sheet.
    2.  Ensures data integrity before inserting it into the destination sheet.

#### 2. **GenerateSalesReport()**

-   **Description**: Generates a formatted sales report based on the latest sales data.
-   **How it works**:
    1.  Applies filters and sorting to the sales data.
    2.  Outputs a neatly formatted report with custom headers and footers.

#### 3. **ValidateData(cellRange As String)**

-   **Description**: Validates the data within a specific range to ensure correct formatting and values.
-   **How it works**:
    1.  Loops through the specified cell range.
    2.  Highlights invalid data and logs errors.

#### 4. **SummarizeMetrics()**

-   **Description**: Summarizes key metrics, such as total sales, average sales per transaction, and highest-selling products.
-   **How it works**:
    1.  Analyzes the entire sales dataset.
    2.  Outputs results in a predefined section of the workbook.

### Example Usage

### Example Usage

```vba
Sub ExampleUsage()
    ' Autofill sales data
    Call AutoFillSalesData
    
    ' Generate a sales report
    Call GenerateSalesReport
    
    ' Validate sales data within a range
    Call ValidateData("A2:A100")
    
    ' Summarize key sales metrics
    Call SummarizeMetrics
End Sub
```
## Contributing

Feel free to contribute to this repository by submitting pull requests. If you have improvements or additional features you'd like to add, please make sure to:

-   Include comments in your code.
-   Write clear descriptions for any new functions or methods.
-   Update this README file with any new documentation.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
