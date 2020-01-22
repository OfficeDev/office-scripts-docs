---
title: 'Run Office Scripts in Excel on the web with Power Automate'
description: 'A tutorial about integrating Power Automate with Office Scripts using a manual trigger.'
ms.date: 01/21/2020
localization_priority: Normal
---

# Run Office Scripts in Excel on the web with Power Automate

This tutorial will teach you how to run an Office Script for Excel on the web through [Power Automate](https://flow.microsoft.com).

## Prerequisites

Before starting this tutorial, you'll need access to Office Scripts, which requires the following:

- [Excel on the web](https://www.office.com/launch/excel).
- Ask your administrator to [enable Office Scripts for your organization](https://support.office.com/article/office-scripts-settings-in-m365-19d3c51a-6ca2-40ab-978d-60fa49554dcf), which adds the **Automate** tab to the ribbon in Excel on the web.
- [Preview access to Power Automate](https://preview.flow.microsoft.com).

> [!IMPORTANT]
> This tutorial assumes you have completed the [Record, edit, and create Office Scripts in Excel on the web](excel-tutorial.md) tutorial.

## Prepare the workbook

Power Automate can't use relative references like `Workbook.getActiveWorksheet` to access workbook components. So, we need a workbook and worksheet with consistent names that Power Automate can reference.

1. Create a new workbook named **MyWorkbook**.

2. In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.

## Create an Office Script

1. Go to the **Automate** tab and select **Code Editor**.

2. Select **New Script**.

3. Replace the default script with the following script. This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.

    ```TypeScript
    async function main(context: Excel.RequestContext) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = context.workbook.worksheets.getItem("TutorialWorksheet")

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.values = [[date.toLocaleDateString()]];

      // Add the time string to B1.
      timeRange.values = [[date.toLocaleTimeString()]];
    }
    ```

4. Rename the script to **Set date and time**. Press the script name to change it.

5. Save the script by pressing **Save Script**.

## Create an automated workflow with Power Automate

1. Sign in to [Power Automate](https://preview.flow.microsoft.com).

2. In the menu that's displayed on the left side of the screen, press **Create**. This brings you to list of ways to create new workflows.

    ![The Create button in Power Automate.](../images/power-automate-tutorial-1.png)

3. In the **Start from blank** section, select **Instant flow**. This creates a manually activated workflow.

    ![The Instant flow option for creating a new workflow.](../images/power-automate-tutorial-2.png)

4. In the dialog window that appears, enter a name for your flow in the **Flow name** text box, select **Manually trigger a flow** from the list of options under **Choose how to trigger the flow**, and press **Create**.

    ![The manual trigger option for creating a new instant flow.](../images/power-automate-tutorial-3.png)

5. Press **New step**.

6. Select the **Standard** tab, then select **Excel Online (Business)**.

    ![The Power Automate option for Excel Online (Business).](../images/power-automate-tutorial-4.png)

7. Under **Actions**, select **Run script (preview)**.

    ![The Power Automate action option for Run script (preview).](../images/power-automate-tutorial-5.png)

8. Specify the following settings for the **Run script** connector:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: MyWorkbook.xlsx
    - **Script**: Set date and time

    ![The connector settings for running a script in Power Automate.](../images/power-automate-tutorial-6.png)

9. Press **Save**.

Your flow is now ready to be run through Power Automate. You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.

## Run the script through Power Automate

1. From the main Power Automate page, select **My flows**.

    ![The My flows button in Power Automate.](../images/power-automate-tutorial-7.png)

2. Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.

3. Press **Run**.

    ![The Run button in Power Automate.](../images/power-automate-tutorial-8.png)

4. A task pane will appear for running the flow. If you are asked to **Sign in** to Excel Online, do so by pressing **Continue**.

5. Press **Run flow**. This runs the flows, which runs the related Office Script.

6. Press **Done**. You should see the **Runs** section update accordingly.

7. Refresh the page to see the results of the Power Automate. If it succeeded, go to the workbook to see the updated cells. If it failed, verify the flow's settings and run it a second time.

    ![Power Automate output showing a successful flow run.](../images/power-automate-tutorial-9.png)

## Next steps

Visit the [Power Automate Documentation](https://docs.microsoft.com/power-automate) to learn ways to automate your Office Scripts.
