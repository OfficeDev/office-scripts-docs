---
title: 'Tutorial: Update a spreadsheet from a Power Automate flow'
description: A tutorial about how to use an Office Script in Power Automate through a manual trigger.
ms.date: 12/22/2025
ms.localizationpriority: high
---

# Tutorial: Update a spreadsheet from a Power Automate flow

This tutorial teaches you how to run an Office Script for Excel through [Power Automate](https://make.powerautomate.com). You'll make a script that updates the values of two cells with the current time. You'll then connect that script to a manually triggered Power Automate flow, so that the script is run whenever a button in Power Automate is selected. Once you understand the basic pattern, you can expand the flow to include other applications and automate more of your daily workflow.

> [!TIP]
> If you are new to Office Scripts, we recommend starting with [Tutorial: Create and format an Excel table](excel-tutorial.md). [Office Scripts use TypeScript](../overview/code-editor-environment.md) and this tutorial is intended for people with beginner to intermediate-level knowledge of JavaScript or TypeScript. If you're new to JavaScript, we recommend starting with the [Mozilla JavaScript tutorial](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

## Prerequisites

[!INCLUDE [Tutorial prerequisites](../includes/power-automate-tutorial-prerequisites.md)]

## Prepare the workbook

Power Automate shouldn't use [relative references](../testing/power-automate-troubleshooting.md#avoid-relative-references) like `Workbook.getActiveWorksheet` to access workbook components. So, you need a workbook and worksheet with consistent names that Power Automate can reference.

1. Create a new workbook named **MyWorkbook**.

1. In the **MyWorkbook** workbook, create a worksheet called **TutorialWorksheet**.

## Create an Office Script

1. Go to the **Automate** tab and select **New Script** > **Create in Code Editor**.

1. Replace the default script with the following script. This script adds the current date and time to the first two cells of the **TutorialWorksheet** worksheet.

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
      // Get the "TutorialWorksheet" worksheet from the workbook.
      let worksheet = workbook.getWorksheet("TutorialWorksheet");

      // Get the cells at A1 and B1.
      let dateRange = worksheet.getRange("A1");
      let timeRange = worksheet.getRange("B1");

      // Get the current date and time using the JavaScript Date object.
      let date = new Date(Date.now());

      // Add the date string to A1.
      dateRange.setValue(date.toLocaleDateString());

      // Add the time string to B1.
      timeRange.setValue(date.toLocaleTimeString());
    }
    ```

1. Rename the script to **Set date and time**. Select the script name to change it.

1. Save the script by selecting **Save Script**.

## Create an automated workflow with Power Automate

1. Sign in to the [Power Automate site](https://make.powerautomate.com).

1. In the menu that's displayed on the left side of the screen, select **Create**. This brings you to a list of ways to create new workflows.

    :::image type="content" source="../images/power-automate-tutorial-1.png" alt-text="The Power Automate 'Create' button.":::

1. In the **Start from blank** section, select **Instant flow**. This creates a manually activated workflow. You can also make scheduled flows and flows that start based on events. These are covered in the next tutorials.

    :::image type="content" source="../images/power-automate-tutorial-2.png" alt-text="The Power Automate Instant flow option for creating a new workflow.":::

1. In the dialog window that appears, enter a name for your flow in the **Flow name** text box. Under **Choose how to trigger the flow**, select **Manually trigger a flow** from the list of options. Select **Create** to finish the initial setup.

    :::image type="content" source="../images/power-automate-tutorial-3.png" alt-text="The Power Automate 'Manually trigger a flow' option.":::

    Note that a manually triggered flow is just one of many types of flows. In the next tutorial, you'll make a flow that automatically runs when you receive an email.

1. In the flow builder, select the **+** button and **Add an action**.

1. In the **Add an action** task pane, search for "Excel run script". Choose the **Excel Online (Business)** connector's **Run script** action. This action runs a script from your OneDrive on a workbook. If you want to use a script stored in your team's SharePoint library, you should use the **Run script from a SharePoint library** action.

    :::image type="content" source="../images/power-automate-tutorial-4.png" alt-text="The action selection task pane showing actions for the Excel Online (Business) connector. The Run script action is highlighted.":::

1. You may be asked to sign in to your Microsoft 365 account. Do so to continue the tutorial.

1. Next, you'll select the workbook and script to use in the flow step. For the tutorial, you'll use the workbook you created in your OneDrive, but you could use any workbook in a OneDrive or SharePoint site. Specify the following parameters for the **Run script** action:

    - **Location**: OneDrive for Business
    - **Document Library**: OneDrive
    - **File**: MyWorkbook.xlsx *(Chosen through the file browser)*
    - **Script**: Set date and time

    :::image type="content" source="../images/power-automate-tutorial-5.png" alt-text="The Power Automate connector settings for running a script.":::

1. Select **Save**.

    :::image type="content" source="../images/power-automate-tutorial-6.png" alt-text="The Save button in Power Automate.":::

Your flow is now ready to be run through Power Automate. You can test it using the **Test** button in the flow editor or follow the remaining tutorial steps to run the flow from your flow collection.

## Run the script through Power Automate

1. From the main Power Automate page, select **My flows**.

    :::image type="content" source="../images/power-automate-tutorial-7.png" alt-text="The My flows button in Power Automate.":::

1. Select **My tutorial flow** from the list of flows displayed in the **My flows** tab. This shows the details of the flow we previously created.

1. Select **Run**.

    :::image type="content" source="../images/power-automate-tutorial-8.png" alt-text="The Run button in Power Automate.":::

1. A task pane will appear for running the flow. If you are asked to **Sign in** to Excel Online, do so by selecting **Continue**.

1. Select **Run flow**. This runs the flow, which runs the related Office Script.

1. Select **Done**. You should see the run history update accordingly.

1. Refresh the page to see the results of the Power Automate. If it failed, verify the flow's settings and run it a second time.

    :::image type="content" source="../images/power-automate-tutorial-9.png" alt-text="Power Automate output showing a successful flow run.":::

1. Open the workbook to see the updated cells. You should see the current date in cell **A1** and the current time in cell **B1**. Power Automate uses Coordinated Universal Time (UTC), so the time will likely be offset from your current time zone.

    :::image type="content" source="../images/power-automate-tutorial-10.png" alt-text="The workbook showing date and time values in cells A1 and B1.":::

## Next steps

Complete [Tutorial: Automatically save content from emails in a workbook](excel-power-automate-trigger.md). It teaches you how to pass data from a workflow service to your Office Script and run the Power Automate flow when certain events occur.
