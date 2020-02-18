---
title: 'Office Scripts sample scenario: Internet Traffic Analysis'
description: 'A sample that takes raw internet traffic data in an Excel workbook and determines the origin location, before organizing that information into a table.'
ms.date: 02/18/2020
localization_priority: Normal
---

# Office Scripts sample scenario: Analyze web downloads

In this scenario, you are tasked with analyzing download reports from your company's website. The goal of this analysis is to determine if the web traffic is coming from the United States or elsewhere in the world.

Your colleagues upload the raw data to your workbook. Each week has its own worksheet. There is also the **Summary** worksheet with a table and chart that show week-over-week trends.

The script you've developed in this scenario runs on one of these weekly worksheets. It parses each IP address associated with a download and determines whether or not it came from the US. The answer is inserted in the worksheet as a boolean value ("TRUE" or "FALSE") and conditional formatting is applied to those cells. The IP address location results are totaled on the worksheet and copied to the summary table.

## Scripting skills covered

- Text parsing
- Subfunctions in scripts
- Conditional formatting
- Tables

## Demo video

This sample was demoed as part of the Office Add-ins developer community call for February 2020.

> [!VIDEO https://www.youtube.com/embed/eRJ71vQ8BaU?start=318]

## Setup instructions

1. Download [analyze-web-downloads.xlsx](analyze-web-downloads.xlsx) to your OneDrive.

2. Open the workbook with Excel for the web.

3. Under the **Automate** tab and open the **Code Editor**.

4. Press **New Script** and paste the following script into the task pane.

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the values of the active range of the active worksheet.
      const logRange = context.workbook.worksheets
        .getActiveWorksheet()
        .getUsedRange()
        .load("values");

      // Get the Summary worksheet and table.
      const summaryWorksheet = context.workbook.worksheets.getItem("Summary");
      const summaryTable = context.workbook.tables.getItem("Table1");

      // Get the range that will contain TRUE/FALSE if the IP address is from the US.
      const isUSColumn = logRange
        .getLastColumn()
        .getOffsetRange(0, 1)
        .load("address");

      // Get the values of all the US IP addresses.
      const ipRange = context.workbook.worksheets
        .getItem("USIPAddresses")
        .getUsedRange()
        .load("values");
      await context.sync();

      // Remove the first row.
      let topRow = logRange.values.shift();

      // Create a new array to contain Is US IP.
      let newCol = [[]];

      // Go through each row in worksheet and add Boolean.
      for (let i = 0; i < logRange.values.length; i++) {
        let curRowIP = logRange.values[i][1];
        newCol.push([findIP(ipRange.values, ipInt(curRowIP)) > 0 ? true : false]);
      }

      // Remove the empty column header and add proper heading.
      newCol.shift();
      newCol.unshift(["Is US IP"]);

      // Show the result in the console or write them to the spreadsheet.
      console.log(
        "IP Address: " + logRange.values[0][1],
        "IP as Integer: " + ipInt(logRange.values[0][1]),
        "Is US IP: " + newCol[1][0]
      );

      isUSColumn.values = newCol;
      addSummaryData();
      applyConditionalFormatting();

      // Get the calculated summary data.
      const summaryRange = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange("H2:K2")
        .load("values");
      await context.sync();

      // Add the corresponding row to the summary table.
      summaryTable.rows.add(null, summaryRange.values);

      // Function to apply conditional formatting to the new column.
      function applyConditionalFormatting() {
        // Add conditional formatting to the new column.
        let conditionalFormatTrue = isUSColumn.conditionalFormats.add(
          Excel.ConditionalFormatType.cellValue
        );
        let conditionalFormatFalse = isUSColumn.conditionalFormats.add(
          Excel.ConditionalFormatType.cellValue
        );
        // Set TRUE to light blue and FALSE to light orange.
        conditionalFormatTrue.cellValue.format.fill.color = "#8FA8DB";
        conditionalFormatTrue.cellValue.rule = {
          formula1: "=TRUE",
          operator: "EqualTo"
        };
        conditionalFormatFalse.cellValue.format.fill.color = "#F8CCAD";
        conditionalFormatFalse.cellValue.rule = {
          formula1: "=FALSE",
          operator: "EqualTo"
        };
      }

      // Adds the summary data to the current sheet and to the summary table.
      function addSummaryData() {
        // Add a summary row and table.
        let summaryHeader = [["Year", "Week", "US", "Other"]];
        let countTrueFormula =
          "=COUNTIF(" + isUSColumn.address + ', "=TRUE")/' + (newCol.length - 1);
        let countFalseFormula =
          "=COUNTIF(" + isUSColumn.address + ', "=FALSE")/' + (newCol.length - 1);

        let summaryContent = [
          [
            '=TEXT(A2,"YYYY")',
            '=TEXTJOIN(" ", FALSE, "Wk", WEEKNUM(A2))',
            countTrueFormula,
            countFalseFormula
          ]
        ];
        let summaryHeaderRow = context.workbook.worksheets
          .getActiveWorksheet()
          .getRange("H1:K1");
        let summaryContentRow = context.workbook.worksheets
          .getActiveWorksheet()
          .getRange("H2:K2");
        summaryHeaderRow.values = summaryHeader;
        summaryContentRow.values = summaryContent;
        let formats = [[".000", ".000"]];
        summaryContentRow
          .getOffsetRange(0, 2)
          .getResizedRange(0, -2).numberFormat = formats;
      }
    }

    // Translate an IP address into an integer.
    function ipInt(ip) {
      // Split the IP address into octets.
      let octets = ip.split(".");

      // Create a number for each octet and do the math to create the integer value of the IP address.
      let fullNum =
        // Define an arbitrary number for the last octet.
        111 +
        parseInt(octets[2]) * 256 +
        parseInt(octets[1]) * 65536 +
        parseInt(octets[0]) * 16777216;
      return fullNum;
    }

    // Return the row number where the ip address is found.
    function findIP(ipLookupTable: number[][], n: number) {
      for (let i = 0; i < ipLookupTable.length; i++) {
        if (ipLookupTable[i][0] <= n && ipLookupTable[i][1] >= n) {
          return i;
        }
      }
      return -1;
    }
    ```

5. Rename the script to **Analyze Web Downloads** and save it.

## Running the script

Run the **Analyze Web Downloads** script on one of the **Week\*\*** worksheets. This will apply the conditional formatting and location labelling on the current sheet. It also updates the **Summary** worksheet.

### Before running the script

![A worksheet that shows raw web traffic data.](../../images/scenario-analyze-web-downloads-before.png)

### After running the script

![A worksheet that shows formatted IP location information with the previous web traffic rows.](../../images/scenario-analyze-web-downloads-after.png)

![The summary table and chart that summarizes the worksheets on which the script has been run.](../../images/scenario-analyze-web-downloads-table.png)
