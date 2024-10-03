---
title: Platform limits and requirements with Office Scripts
description: Resource limits and browser support for Office Scripts when used with Excel.
ms.topic: limits-and-quotas
ms.date: 10/03/2024
ms.localizationpriority: medium
---

# Platform limits and requirements with Office Scripts

There are some platform limitations of which you should be aware when developing Office Scripts. This article details the browser support and data limits for Office Scripts for Excel.

## Platform support

[!INCLUDE [Platform requirements](../includes/platform-requirements.md)]

> [!NOTE]
> If you meet these requirements and are still not seeing the **Automate** tab, it's possible that your admin has disabled the feature or there's some other problem with your environment. Please follow the steps under [Automate tab not appearing or Office Scripts unavailable](../testing/troubleshooting.md#automate-tab-not-appearing-or-office-scripts-unavailable) to start using Office Scripts.

[!INCLUDE [Power Automate license support](../includes/power-automate-needs-business.md)]

## Data limits

There are limits on how much Excel data can be transferred at once and how many individual Power Automate transactions can be conducted.

### Excel

Excel on the web has the following limitations when making calls to the workbook through a script.

- Requests and responses are limited to **5MB**.
- A range is limited to **five million cells**.

When you exceed one of the Excel data limits, you receive this error message: "The response payload size has exceeded the limit."

If you're encountering errors when dealing with large datasets, try using multiple smaller ranges instead of larger ranges. For an example, see the [Write a large dataset](../resources/samples/write-large-dataset.md) sample. You can also use APIs like [Range.getSpecialCells](/javascript/api/office-scripts/excelscript/excelscript.range#excelscript-excelscript-range-getspecialcells-member(1)) to target specific cells instead of large ranges.

Excel limits that aren't specific to Office Scripts can be found in the article [Excel specifications and limits](https://support.microsoft.com/office/1672b34d-7043-467e-8e27-269d656771c3).

### Power Automate

The following limitations with the Power Automate platform are the ones you're most likely to encounter.

- Each user is limited to **1,600 calls** to the Run script action per day. This limit resets at 12:00 AM UTC.
- There's a **120-second timeout** for [synchronous Power Automate operations](/power-automate/limits-and-config#timeout). For long-running scripts, you must either [optimize your script](../develop/web-client-performance.md) or [split your Excel automation into multiple script calls](../resources/samples/write-large-dataset.md#sample-2-write-data-in-batches-from-a-power-automate-flow).
- The maximum size of parameters passed to the Run script action is **30,000,000 bytes (28.6MB)**.

Additional Power Automate platform usage limitations can be found in the following articles.

- [Limits and configuration in Power Automate](/power-automate/limits-and-config)
- [Known issues and limitations for the Excel Online (Business) connector](/connectors/excelonlinebusiness/#known-issues-and-limitations)

## Power Automate specific restrictions

There are a few important differences between running a script in the Excel application and running a script as part of a Power Automate flow.

### No external calls from a script

[!INCLUDE [External calls in Power Automate](../includes/external-calls-power-automate.md)]

### API behavior differences

Some APIs behave differently when run with Power Automate. Others fail due to their reliance on the Excel UI. The full lists are found in [Troubleshoot Office Scripts running in Power Automate](power-automate-troubleshooting.md).

### ISO strict Open XML workbooks aren't supported

The **Excel Business (Online)** connector's **Run script** action doesn't support workbooks with the [ISO strict version of the Excel Workbook file format](https://www.loc.gov/preservation/digital/formats/fdd/fdd000401.shtml). Flows with this type of workbook return a "BadGateway" error when trying to run a script. This is due to coauthoring restrictions. Please save workbooks as the standard Excel Workbook format for use with Power Automate.

## Teams support

[!INCLUDE [Teams support note](../includes/teams-support-note.md)]

## Government cloud support

Office Scripts aren't supported on [GCC High or above](/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/gcc-high-and-dod). [External calls](../develop/external-calls.md) from scripts may be affected by firewall settings in other government clouds.

## Third-party cookies for Excel on the web

Your browser needs third-party cookies enabled to show the **Automate** tab in Excel on the web. Check your browser settings if the tab isn't being displayed. If you're using a private browser session, you may need to re-enable this setting each time.

> [!NOTE]
> Some browsers refer to this setting as "all cookies", instead of "third-party cookies".

### How to adjust cookie settings in popular browsers

- [Chrome](https://support.google.com/chrome/answer/95647)
- [Edge](https://support.microsoft.com/microsoft-edge/597f04f2-c0ce-f08c-7c2b-541086362bd2)
- [Firefox](https://support.mozilla.org/kb/disable-third-party-cookies)
- [Safari](https://support.apple.com/guide/safari/manage-cookies-and-website-data-sfri11471/mac)

## Conditional Access

[Conditional Access](/azure/active-directory/conditional-access/overview) policies restrict access to SharePoint and OneDrive for [unmanaged devices](/sharepoint/control-access-from-unmanaged-devices). If your device isn't managed by the tenant, you may not have access to specific scripts, or may only be able to access them through the browser.

If you script is blocked by Conditional Access policies, you receive one of two error messages. These messages also surface in Power Automate if your flow is run from an unmanaged device.

- "Due to organizational policies, you can’t access this resource from this untrusted device."
- "We can't find this script. It may have been deleted by another user." (If your version of Excel is older.)

> [!IMPORTANT]
> Administrators should consider blocking all access to Power Automate from unmanaged devices. This process is detailed in the blog post [Control Access to Power Apps and Power Automate with Azure AD Conditional Access Policies](https://devblogs.microsoft.com/premier-developer/control-access-to-power-apps-and-power-automate-with-azure-ad-conditional-access-policies/).

## API support on older Excel versions

Some Office Scripts APIs may not be supported by Excel for Windows or Excel for Mac, especially older builds. These include newer APIs and APIs for web-only features. If a script contains unsupported APIs, the Code Editor displays a warning. If you try to run such a script, it won't run. Instead, the **Script Run Status** task pane displays a warning message that says, "This script currently must be run on Excel for the web. Open the workbook in the browser then try again, or contact the script owner for help."

Using an older version of Excel to open workbooks with scripts shared in them has no effect on the script itself.

## See also

- [Excel specifications and limits](https://support.microsoft.com/office/1672b34d-7043-467e-8e27-269d656771c3)
- [Troubleshoot Office Scripts](troubleshooting.md)
- [Undo the effects of Office Scripts](undo.md)
- [Improve the performance of your Office Scripts](../develop/web-client-performance.md)
