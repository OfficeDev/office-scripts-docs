---
title: 'Office Scripts sample scenario: Automated task reminders'
description: A sample that uses Power Automate and Adaptive Cards automate task reminders in a project management spreadsheet.
ms.date: 08/10/2023
ms.localizationpriority: medium
---

# Office Scripts sample scenario: Automated task reminders

In this scenario you're managing a project. You use an Excel worksheet to track your employees' status every month. You often need to remind people to fill out their status, so you've decided to automate that reminder process.

You'll create a Power Automate flow to message people with missing status fields and apply their responses to the spreadsheet. To do this, you'll develop a pair of scripts to handle the working with the workbook. The first script gets a list of people with blank statuses and the second script adds a status string to the right row. You'll also make use of [Teams Adaptive Cards](/microsoftteams/platform/task-modules-and-cards/what-are-cards) to have employees enter their status directly from the notification.

## Scripting skills covered

- Create flows in Power Automate
- Pass data to scripts
- Return data from scripts
- Teams Adaptive Cards
- Tables

## Prerequisites

This scenario uses [Power Automate](https://make.powerautomate.com) and [Microsoft Teams](https://www.microsoft.com/microsoft-365/microsoft-teams/group-chat-software). You will need both associated with the account that you use for developing Office Scripts. For free access to a Microsoft Developer subscription to learn about and work with these applications, consider joining the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program).

## Setup instructions

1. Download [task-reminders.xlsx](task-reminders.xlsx) to your OneDrive.

1. Open the workbook in Excel.

1. First, we need a script to get all the employees with status reports that are missing from the spreadsheet. Under the **Automate** tab, select **New Script** and paste the following script into the editor.

    ```TypeScript
    /**
     * This script looks for missing status reports in a project management table.
     *
     * @returns An array of Employee objects (containing their names and emails).
     */
    function main(workbook: ExcelScript.Workbook): Employee[] {
      // Get the first worksheet and the first table on that worksheet.
      let sheet = workbook.getFirstWorksheet()
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the data for the whole table.
      let bodyRangeValues = table.getRangeBetweenHeaderAndTotal().getValues();

      // Create the array of Employee objects to return.
      let people: Employee[] = [];

      // Loop through the table and check each row for completion.
      for (let i = 0; i < bodyRangeValues.length; i++) {
        let row = bodyRangeValues[i];
        if (row[STATUS_REPORT_INDEX] === "") {
          // Save the email to return.
          people.push({ name: row[NAME_INDEX].toString(), email: row[EMAIL_INDEX].toString() });
        }
      }

      // Log the array to verify we're getting the right rows.
      console.log(people);

      // Return the array of Employees.
      return people;
    }

    /**
     * An interface representing an employee.
     * An array of Employees will be returned from the script
     * for the Power Automate flow.
     */
    interface Employee {
      name: string;
      email: string;
    }
    ```

1. Save the script with the name **Get People**.

1. Next, we need a second script to process the status report cards and put the new information in the spreadsheet. In the Code Editor task pane, select **New Script** and paste the following script into the editor.

    ```TypeScript
    /**
     * This script applies the results of a Teams Adaptive Card about
     * a status update to a project management table.
     *
     * @param senderEmail - The email address of the employee updating their status.
     * @param statusReportResponse - The employee's status report.
     */
    function main(workbook: ExcelScript.Workbook,
      senderEmail: string,
      statusReportResponse: string) {

      // Get the first worksheet and the first table in that worksheet.
      let sheet = workbook.getFirstWorksheet();
      let table = sheet.getTables()[0];

      // Give the column indices names matching their expected content.
      const NAME_INDEX = 0;
      const EMAIL_INDEX = 1;
      const STATUS_REPORT_INDEX = 2;

      // Get the range and data for the whole table.
      let bodyRange = table.getRangeBetweenHeaderAndTotal();
      let tableRowCount = bodyRange.getRowCount();
      let bodyRangeValues = bodyRange.getValues();

      // Create a flag to denote success.
      let statusAdded = false;

      // Loop through the table and check each row for a matching email address.
      for (let i = 0; i < tableRowCount && !statusAdded; i++) {
        let row = bodyRangeValues[i];

        // Check if the row's email address matches.
        if (row[EMAIL_INDEX] === senderEmail) {
          // Add the Teams Adaptive Card response to the table.
          bodyRange.getCell(i, STATUS_REPORT_INDEX).setValues([
            [statusReportResponse]
          ]);
          statusAdded = true;
        }
      }

      // If successful, log the status update.
      if (statusAdded) {
        console.log(
          `Successfully added status report for ${senderEmail} containing: ${statusReportResponse}`
        );
      }
    }
    ```

1. Save the script with the name **Save Status**.

1. Now, we need to create the flow. Open the [Power Automate **Create** tab](https://make.powerautomate.com/create).

    > [!TIP]
    > If you haven't created a flow before, please check out our tutorial [Start using scripts with Power Automate](../../tutorials/excel-power-automate-manual.md) to learn the basics.

1. Create a new **Instant cloud flow**.

1. Choose **Manually trigger a flow** from the options and select **Create**.

1. The flow needs to call the **Get People** script to get all the employees with empty status fields. Select **New step**, then select **Excel Online (Business)**. Under **Actions**, select **Run script**. Provide the following entries for the flow step:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: task-reminders.xlsx *(Chosen through the file browser)*
    - **Script**: Get People

    :::image type="content" source="../../images/scenario-task-reminders-1.png" alt-text="The Power Automate flow showing the first Run script flow step.":::

1. Next, the flow needs to process each Employee in the array returned by the script. Select **New step**, search for **Post adaptive card and wait for a response**, and select that Microsoft Teams connector action.

1. Sending an Adaptive Card requires the card's [JSON](https://www.w3schools.com/whatis/whatis_json.asp) to be provided as the **Message**. You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to create custom cards. For this sample, use the following JSON.  

    ```json
    {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "size": "Medium",
          "weight": "Bolder",
          "text": "Update your Status Report"
        },
        {
          "type": "Image",
          "altText": "",
          "url": "https://i.imgur.com/f5RcuF3.png"
        },
        {
          "type": "TextBlock",
          "text": "This is a reminder to update your status report for this month's review. You can do so right here in this card, or by adding it directly to the spreadsheet.",
          "wrap": true
        },
        {
          "type": "Input.Text",
          "placeholder": "My status report for this month is...",
          "id": "response",
          "isMultiline": true
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "Submit",
          "id": "submit"
        }
      ]
    }
    ```

1. For the **Recipient** field, add **email** from the dynamic content (the selection will have the Excel logo by it). Adding **email** causes the flow step to be surrounded by an **Apply to each** block. That means the array will be iterated over by Power Automate.

1. Fill out the remaining fields as follows:

    - **Post as**: Flow bot
    - **Post in**: Chat with Flow bot
    - **Update message**: Thank you for submitting your status report. Your response has been successfully added to the spreadsheet.

    :::image type="content" source="../../images/scenario-task-reminders-2.png" alt-text="The Power Automate flow showing the completed adaptive card action.":::

1. In the **Apply to each** block, following the **Post adaptive card and wait for a response**, select **Add an action**. Select **Excel Online (Business)**. Under **Actions**, select **Run script**. Provide the following entries for the flow step:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: task-reminders.xlsx *(Chosen through the file browser)*
    - **Script**: Save Status
    - **senderEmail**: email *(dynamic content from Excel)*
    - **statusReportResponse**: response *(dynamic content from Teams)*

    :::image type="content" source="../../images/scenario-task-reminders-3.png" alt-text="The Power Automate flow showing the apply-to-each step.":::

1. Save the flow.

## Running the flow

To test the flow, make sure any table rows with blank status use an email address tied to a Teams account (you should probably use your own email address while testing). Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.

You should receive an Adaptive Card from Power Automate through Teams. Once you fill out the status field in the card, the flow will continue and update the spreadsheet with the status you provide.

### Before running the flow

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-before.png" alt-text="A worksheet with a status report containing one missing status entry.":::

### Receiving the Adaptive Card

:::image type="content" source="../../images/scenario-task-reminders-adaptive-card.png" alt-text="An Adaptive Card in Teams asking the employee for a status update.":::

### After running the flow

:::image type="content" source="../../images/scenario-task-reminders-spreadsheet-after.png" alt-text="A worksheet with a status report with a now-filled-in status entry.":::
