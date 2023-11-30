---
title: 'Tutorial: Send weekly email reminders based on spreadsheet data'
description: A tutorial that shows how to send reminder emails by running Office Scripts for Excel through Power Automate.
ms.date: 11/29/2023
ms.localizationpriority: high
---

# Tutorial: Send weekly email reminders based on spreadsheet data

This tutorial teaches you how to return information from an Office Script for Excel as part of an automated [Power Automate](https://make.powerautomate.com) workflow. You'll make a script that looks through a schedule and works with a flow to send reminder emails. This flow will run on a regular schedule, providing these reminders on your behalf.

> [!TIP]
> If you're new to Office Scripts, we recommend starting with [Tutorial: Create and format an Excel table](excel-tutorial.md).
>
> If you're new to Power Automate, we recommend starting with [Tutorial: Update a spreadsheet from a Power Automate flow](excel-power-automate-manual.md) and [Tutorial: Automatically save content from emails in a workbook](excel-power-automate-trigger.md).
>
> [Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript. If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## Prerequisites

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## Prepare the workbook

1. Download the workbook [on-call-rotation.xlsx](on-call-rotation.xlsx) to your OneDrive.

1. Open **on-call-rotation.xlsx** in Excel.

1. Add a row to the table with your name, email address, and start and end dates that overlap with the current date.

    > [!IMPORTANT]
    > The script you'll write uses the first matching entry in the table, so make sure your name is above any row with the current week.

    :::image type="content" source="../images/power-automate-return-tutorial-1.png" alt-text="A worksheet containing the on-call rotation table data.":::

## Create an Office Script

1. Go to the **Automate** tab and select **New Script**.

1. Name the script **Get On-Call Person**.

1. You should now have an empty script. You want a script that gets an email address from the spreadsheet. Change `main` to return a string, like this:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
    }
    ```

1. Next, you need to get all the data from the table. That lets the script look at each row. Add the following code inside the `main` function.

    ```TypeScript
    // Get the H1 worksheet.
    let worksheet = workbook.getWorksheet("H1");

    // Get the first (and only) table in the worksheet.
    let table = worksheet.getTables()[0];

    // Get the data from the table.
    let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    ```

1. The dates in the table are stored using [Excel's date serial number](https://support.microsoft.com/office/e7fe7167-48a9-4b96-bb53-5612a800b487). You need to convert those dates to JavaScript dates in order to compare them. Add the following helper function outside of the `main` function.

    ```TypeScript
    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

1. Now, you need to figure out which person is on call right now. Their row will have a start and end date surrounding the current date. The script will assume only one person is on call at a time. Scripts can return arrays to handle multiple values, but you can return the first matching email address for this tutorial. Add the following code to the end of the `main` function.

    ```TypeScript
    // Look for the first row where today's date is between the row's start and end dates.
    let currentDate = new Date();
    for (let row = 0; row < tableValues.length; row++) {
        let startDate = convertDate(tableValues[row][2] as number);
        let endDate = convertDate(tableValues[row][3] as number);
        if (startDate <= currentDate && endDate >= currentDate) {
            // Return the first matching email address.
            return tableValues[row][1].toString();
        }
    }
    ```

1. The final script should look like this:

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) : string {
        // Get the H1 worksheet.
        let worksheet = workbook.getWorksheet("H1");

        // Get the first (and only) table in the worksheet.
        let table = worksheet.getTables()[0];
    
        // Get the data from the table.
        let tableValues = table.getRangeBetweenHeaderAndTotal().getValues();
    
        // Look for the first row where today's date is between the row's start and end dates.
        let currentDate = new Date();
        for (let row = 0; row < tableValues.length; row++) {
            let startDate = convertDate(tableValues[row][2] as number);
            let endDate = convertDate(tableValues[row][3] as number);
            if (startDate <= currentDate && endDate >= currentDate) {
                // Return the first matching email address.
                return tableValues[row][1].toString();
            }
        }
    }

    // Convert the Excel date to a JavaScript Date object.
    function convertDate(excelDateValue: number) {
        let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
        return javaScriptDate;
    }
    ```

## Create an automated workflow with Power Automate

1. Sign in to the [Power Automate site](https://make.powerautomate.com).

1. In the menu that's displayed on the left side of the screen, select **Create**. This brings you to list of ways to create new workflows.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="The Create button in Power Automate.":::

1. Under the **Start from blank** section, select **Scheduled cloud flow**.

    :::image type="content" source="../images/power-automate-return-tutorial-2.png" alt-text="The Scheduled cloud flow button in Power Automate.":::

1. Next, set the schedule for this flow. Your spreadsheet has a new on-call assignment starting every Monday in the first half of 2024. Set the flow to run first thing Monday mornings. Use the following options to configure the flow to run on Monday each week.

    - **Flow name**: Notify On-Call Person
    - **Starting**: 11/27/23 at 1:00am
    - **Repeat every**: 1 Week
    - **On these days**: M

    :::image type="content" source="../images/power-automate-return-tutorial-3.png" alt-text="The Power Automate 'Build a scheduled cloud flow' dialog showing options. The options include flow name, time to start, how often to repeat, and one which day of the week to run the flow.":::

1. Select **Create**.

1. In the flow builder, select the **+** button and **Add an action**.

1. In the **Add an action** task pane, search for "Excel run script". Choose the **Excel Online (Business)** connector's **Run script** action. This action runs a script from your OneDrive on a workbook. If you want to use a script stored in your team's SharePoint library, you should use the **Run script from a SharePoint library** action.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="The action selection task pane showing actions for the Excel Online (Business) connector. The Run script action is highlighted.":::

1. You may be asked to sign in to your Microsoft 365 account. Do so to continue the tutorial.

1. Next, you'll select the workbook and script to use in the flow step. For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site. Specify the following parameters for the **Run script** action:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: on-call-rotation.xlsx *(Chosen through the file browser)*
    - **Script**: Get On-Call Person

    :::image type="content" source="../images/power-automate-return-tutorial-4.png" alt-text="The Power Automate connector settings for running a script.":::

1. In the flow builder, select the **+** button and **Add an action**.

1. End the flow by sending the reminder email. In the **Add an action** task pane, search for "send an email". Choose the **Office 365 Outlook** connector's **Send an email (V2)** action.

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="The action selection task pane showing actions for the Office 365 Outlook connector. The Send an email (V2) action is highlighted.":::

    > [!NOTE]
    > This tutorial uses Outlook. Feel free to use your preferred email service instead, though some options may be different.

1. For the **To** parameter, select the text box and select **Enter custom value**. Use the dynamic content control to add the email address returned by the script. This will be labelled **result** with the Excel icon next to it. You can provide whatever subject and body text you'd like.

    :::image type="content" source="../images/power-automate-return-tutorial-5.png" alt-text="The Power Automate Outlook connector settings for sending an email. The options include the file to send, the subject of the email, and the body of the email as well as advanced options.":::

1. Select **Save**.

## Test the script in Power Automate

Your flow will run every Monday morning. You can test the script now by selecting the **Test** button in the upper-right corner of the screen. Select **Manually**, then select **Run Test** to run the flow now and test the behavior. You may need to grant permissions to Excel and Outlook to continue.

:::image type="content" source="../images/power-automate-return-tutorial-6.png" alt-text="The Power Automate Test button.":::

> [!TIP]
> If your flow fails to send an email, double-check in the spreadsheet that a valid email is listed for the current date range at the top of the table.

## Next steps

Visit [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn more about connecting Office Scripts with Power Automate.

You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.
