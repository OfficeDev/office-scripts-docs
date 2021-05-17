---
title: 'Pass data to scripts in an automatically-run Power Automate flow'
description: 'A tutorial about running Office Scripts for Excel on the web through Power Automate when mail is received and passing flow data to the script.'
ms.date: 12/28/2020
localization_priority: Priority
---

# Pass data to scripts in an automatically-run Power Automate flow

This tutorial teaches you how to use an Office Script for Excel on the web with an automated [Power Automate](https://flow.microsoft.com) workflow. Your script will automatically run each time you receive an email, recording information from the email in an Excel workbook. Being able to pass data from other applications into an Office Script gives you a great deal of flexibility and freedom in your automated processes.

> [!TIP]
> If you are new to Office Scripts, we recommend starting with the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial. If you are new to Power Automate, we recommend starting with the [Call scripts from a manual Power Automate flow](excel-power-automate-manual.md) tutorial. [Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript. If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## Prerequisites

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## Prepare the workbook

Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-relative-references) like `Workbook.getActiveWorksheet` to access workbook components. So, we need a workbook and worksheet with consistent names for Power Automate to reference.

1. Create a new workbook named **MyWorkbook**.

2. Go to the **Automate** tab and select **All Scripts**.

3. Select **New Script**.

4. Replace the existing code with the following script and press **Run**. This will setup the workbook with consistent worksheet, table, and PivotTable names.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Add a new worksheet to store our email table
      let emailsSheet = workbook.addWorksheet("Emails");

      // Add data and create a table
      emailsSheet.getRange("A1:D1").setValues([
        ["Date", "Day of the week", "Email address", "Subject"]
      ]);
      let newTable = workbook.addTable(emailsSheet.getRange("A1:D2"), true);
      newTable.setName("EmailTable");

      // Add a new PivotTable to a new worksheet
      let pivotWorksheet = workbook.addWorksheet("Subjects");
      let newPivotTable = workbook.addPivotTable("Pivot", "EmailTable", pivotWorksheet.getRange("A3:C20"));

      // Setup the pivot hierarchies
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Day of the week"));
      newPivotTable.addRowHierarchy(newPivotTable.getHierarchy("Email address"));
      newPivotTable.addDataHierarchy(newPivotTable.getHierarchy("Subject"));
    }
    ```

## Create an Office Script

Let's create a script that logs information from an email. We want to know which days of the week we receive the most mail and how many unique senders are sending that mail. Our workbook has a table with **Date**, **Day of the week**, **Email address**, and **Subject** columns. Our worksheet also has a PivotTable that is pivoting on the **Day of the week** and **Email address** (those are the row hierarchies). The count of unique **Subjects** is the aggregated information being displayed (the data hierarchy). We'll have our script refresh that PivotTable after updating the email table.

1. From within the **Code Editor** task pane, select **New Script**.

2. The flow that we'll create later in the tutorial will send our script information about each email that's received. The script needs to accept that input through parameters in the `main` function. Replace the default script with the following script:

    ```TypeScript
    function main(
      workbook: ExcelScript.Workbook,
      from: string,
      dateReceived: string,
      subject: string) {

    }
    ```

3. The script needs access to the workbook's table and PivotTable. Add the following code to the body of the script, after the opening `{`:

    ```TypeScript
    // Get the email table.
    let emailWorksheet = workbook.getWorksheet("Emails");
    let table = emailWorksheet.getTable("EmailTable");
  
    // Get the PivotTable.
    let pivotTableWorksheet = workbook.getWorksheet("Subjects");
    let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");
    ```

4. The `dateReceived` parameter is of type `string`. Let's convert that to a [`Date` object](../develop/javascript-objects.md#date) so we can easily get the day of the week. After doing that, we'll need to map the day's number value to a more readable version. Add the following code to the end of your script, before the closing `}`:

    ```TypeScript
      // Parse the received date string to determine the day of the week.
      let emailDate = new Date(dateReceived);
      let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });
    ```

5. The `subject` string may include the "RE:" reply tag. Let's remove that from the string so that emails in the same thread have the same subject for the table. Add the following code to the end of your script, before the closing `}`:

    ```TypeScript
    // Remove the reply tag from the email subject to group emails on the same thread.
    let subjectText = subject.replace("Re: ", "");
    subjectText = subjectText.replace("RE: ", "");
    ```

6. Now that the email data has been formatted to our liking, let's add a row to the email table. Add the following code to the end of your script, before the closing `}`:

    ```TypeScript
    // Add the parsed text to the table.
    table.addRow(-1, [dateReceived, dayName, from, subjectText]);
    ```

7. Finally, let's make sure the PivotTable is refreshed. Add the following code to the end of your script, before the closing `}`:

    ```TypeScript
    // Refresh the PivotTable to include the new row.
    pivotTable.refresh();
    ```

8. Rename your script **Record Email** and press **Save script**.

Your script is now ready for a Power Automate workflow. It should look like the following script:

```TypeScript
function main(
  workbook: ExcelScript.Workbook,
  from: string,
  dateReceived: string,
  subject: string) {
  // Get the email table.
  let emailWorksheet = workbook.getWorksheet("Emails");
  let table = emailWorksheet.getTable("EmailTable");

  // Get the PivotTable.
  let pivotTableWorksheet = workbook.getWorksheet("Subjects");
  let pivotTable = pivotTableWorksheet.getPivotTable("Pivot");

  // Parse the received date string to determine the day of the week.
  let emailDate = new Date(dateReceived);
  let dayName = emailDate.toLocaleDateString("en-US", { weekday: 'long' });

  // Remove the reply tag from the email subject to group emails on the same thread.
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("RE: ", "");

  // Add the parsed text to the table.
  table.addRow(-1, [dateReceived, dayName, from, subjectText]);

  // Refresh the PivotTable to include the new row.
  pivotTable.refresh();
}
```

## Create an automated workflow with Power Automate

1. Sign in to the [Power Automate site](https://flow.microsoft.com).

2. In the menu that's displayed on the left side of the screen, press **Create**. This brings you to list of ways to create new workflows.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="The Power Automate Create button":::

3. In the **Start from blank** section, select **Automated flow**. This creates a workflow triggered by an event, such as receiving an email.

    :::image type="content" source="../images/power-automate-params-tutorial-1.png" alt-text="The Automated flow option in Power Automate":::

4. In the dialog window that appears, enter a name for your flow in the **Flow name** text box. Then select **When a new email arrives** from the list of options under **Choose your flow's trigger**. You may need to search for the option using the search box. Finally, press **Create**.

    :::image type="content" source="../images/power-automate-params-tutorial-2.png" alt-text="Part of the Power Automate flow showing the 'flow name' and the 'choose your flow's trigger' options. The flow name is 'Record Email Flow' and the trigger is the 'When a new email arrives in Outlook' option":::

    > [!NOTE]
    > This tutorial uses Outlook. Feel free to use your preferred email service instead, though some options may be different.

5. Press **New step**.

6. Select the **Standard** tab, then select **Excel Online (Business)**.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="Excel Online (Business) option in Power Automate":::

7. Under **Actions**, select **Run script**.

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="Run script action option in Power Automate":::

8. Next, you'll select the workbook, script, and script input arguments to use in the flow step. For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site. Specify the following settings for the **Run script** connector:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: MyWorkbook.xlsx *(Chosen through the file browser)*
    - **Script**: Record Email
    - **from**: From *(dynamic content from Outlook)*
    - **dateReceived**: Received Time *(dynamic content from Outlook)*
    - **subject**: Subject *(dynamic content from Outlook)*

    *Note that the parameters for the script will only appear once the script is selected.*

    :::image type="content" source="../images/power-automate-params-tutorial-3.png" alt-text="The Power Automate run script action showing the options that appear once the script is selected":::

9. Press **Save**.

Your flow is now enabled. It will automatically run your script each time you receive an email through Outlook.

## Manage the script in Power Automate

1. From the main Power Automate page, select **My flows**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="The My flows button in Power Automate":::

2. Select your flow. Here you can see the run history. You can refresh the page or press the refresh **All runs** button to update the history. The flow will trigger shortly after an email is received. Test the flow by sending yourself mail.

When the flow is triggered and successfully runs your script, you should see the workbook's table and PivotTable update.

:::image type="content" source="../images/power-automate-params-tutorial-4.png" alt-text="A worksheet showing the email table after the flow has run three times":::

:::image type="content" source="../images/power-automate-params-tutorial-5.png" alt-text="A worksheet showing the PivotTable after the flow has run three times":::

## Next steps

Complete the [Return data from a script to an automatically-run Power Automate flow](excel-power-automate-returns.md) tutorial. It teaches you how to return data from a script to the flow.

You can also check out the [Automated task reminders sample scenario](../resources/scenarios/task-reminders.md) to learn how to combine Office Scripts and Power Automate with Teams Adaptive Cards.
