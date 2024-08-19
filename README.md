# Financial Statement Processor

Welcome to the **Financial Statement Processor**! This tool is designed to read financial statements in Excel format, process the data based on given projections, and generate a clean output file. The primary functionalities include data extraction, projection updates, and output generation. Below is a detailed guide on how to use this tool, including prerequisites, working, handling scenarios, and steps to start the service.

## Prerequisites

Before you start, make sure you have the following:

1. **Input Excel File**: Save your financial statement Excel file in the project directory with the filename `Financial Projections.xlsx` and sheet name as `Sheet1`.
2. **Node.js**: Ensure you have Node.js installed on your system.

## Getting Started

1. Ensure that the Excel file `Financial Projections.xlsx` is in the project directory.
2. Install the required dependencies with `npm install`.
3. Compile TypeScript code using `tsc`.
4. Start the service using `node index.js`.
5. Use the API endpoint to submit your projection updates.
6. Check the project directory for the cleaned Excel file named `cleanFinancialStatement.xlsx`.

## Steps to Start the Service

1. **Install Dependencies**:
   Open a terminal in the project directory and run:
   ```sh
   npm install
   ```

2. **Prepare the Excel File**:
   Place your financial statement Excel file in the project directory with the filename `Financial Projections.xlsx`.

3. **Compile TypeScript**:
   Project uses `TypeScript`, compile it by running:
   ```sh
   tsc
   ```

4. **Start the Service**:
   Start the service using Node.js with the command:
   ```sh
   node index.js
   ```

5. **Execute the API**:
   Once the service is running, you can execute the API to process and clean and process the financial data.

## How It Works

### API Request

You will use the API endpoint to submit your request. The API accepts JSON data with updated projection percentages.

**URL**: `http://localhost:8080/processNClean`

**Request Payload**:

```json
{
   "projectionFieldInterests": {
       "Product Sales": 10,
       "Cost of Goods Sold": 5
       // Add more fields and their respective projection percentages
   }
}
```

**Example cURL Request**:

```sh
curl --location 'http://localhost:8080/processNClean' \
--header 'Content-Type: application/json' \
--data '{
    "projectionFieldInterests": {
        "Product Sales": 10,
        "Cost of Goods Sold": 5
    }}'
```

**Response**:

```json
{
   "message": "New clean excel process and created 'cleanFinancialStatement.xlsx'",
   "data": {
       "historical": {
           "1719813600000": {
               "Product Sales": {
                   "value": 80000,
                   "month": "July 2024"
               },
               // More data...
           },
           // More historical data...
       },
       "projection": {
           "1722492000000": {
               "Product Sales": {
                   "value": 88000,
                   "month": "August 2024"
               },
               // More data...
           },
           // More projection data...
       }
   }
}
```

## Handling and Flexibility

### Auto Calculation of Totals

- The tool auto-calculates all totals based on preceding fields, helping to correct minor client errors in total calculations.

### Comment Handling

- The parser can handle comments in the Excel sheets with some limitations. Comments should be placed in intermediate columns as the parser will ignore them based on predefined rules.

### Presumptions

1. **Field Alignment**: Fields on the left will be in the same column and appear first when scanning from left to right.
2. **Headings**: Historical and projection headings should be in the given casing.
3. **Month Headings**: Month headings should be below the main headings and start in the same column as the first historical column.

### Scenarios

1. **Multiple Historical Columns**: The parser handles multiple columns of historical data. It identifies the most recent historical data to apply updated projections.

<img width="437" alt="image" src="https://github.com/user-attachments/assets/97a1d2f6-5bf8-4852-a008-26de87608184">

2. **Comment Handling**: Comments placed in intermediate columns are ignored. However, ensure comments do not interfere with field names or data alignment, as the parser assumes comments are irrelevant to the data extraction process.

<img width="495" alt="image" src="https://github.com/user-attachments/assets/3fafc721-3c45-45d2-9e4e-8b30dac9439a">

3. **Incorrect Totals**: The tool corrects minor errors in totals by auto-calculating them based on preceding fields. If totals are incorrect, the tool uses logical calculations to correct them, maintaining the integrity of the final output.

<img width="410" alt="image" src="https://github.com/user-attachments/assets/9c950005-8bb0-47c6-b883-68d8f8d9f88c">

4. **Updating % for Fields**: You can update projection percentages for specific fields using the `projectionFieldInterests` in the API request. The tool applies these updates to the most recent projection data and generates a clean output reflecting the new projections.(Note: Adding auto calculated fields to projectionFieldInterests would not affect the sheet, as those are auto-calculated on the fly).

<img width="158" alt="image" src="https://github.com/user-attachments/assets/c127a62a-35b2-48be-8a3c-593b2bc8a553">

Clean excle created for all above scenarios : 

<img width="478" alt="image" src="https://github.com/user-attachments/assets/941cd5c1-a865-445e-980b-f45b77639e30">

Happy processing! 🚀
