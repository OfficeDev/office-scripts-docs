---
title: 'Cross-reference Excel files with Power Automate'
description: 'Learn how to use Office Scripts and Power Automate to cross-reference and format an Excel file.'
ms.date: 06/25/2021
localization_priority: Normal
---

# Cross-reference Excel files with Power Automate

This solution shows how to compare data across two Excel files to find discrepancies. It uses Office Scripts to analyze data and Power Automate to communicate between the workbooks.

## Example scenario

You're an event coordinator who is scheduling speakers for upcoming conferences. You keep the event data in one spreadsheet and the speaker registrations in another. To ensure the two workbooks are kept in sync, you use a flow with Office Scripts to highlight any potential problems.

## Sample Excel files

Download the following files used in this solution to try it out yourself!

1. <a href="event-data.xlsx">event-data.xlsx</a>
1. <a href="speaker-registrations.xlsx">speaker-registrations.xlsx</a>

## Sample code: Get event data

```TypeScript
function main(workbook: ExcelScript.Workbook): string {
  // Get the first table in the "Keys" worksheet.
  let table = workbook.getWorksheet('Keys').getTables()[0];

  // Get the rows in the event table.
  let range = table.getRangeBetweenHeaderAndTotal();
  let rows = range.getValues();

  // Save each row as an EventData object. This lets them be passed through Power Automate.
  let records: EventData[] = [];
  for (let row of rows) {
    let [eventId, date, location, capacity] = row;
    records.push({
      eventId: eventId as string,
      date: date as number,
      location: location as string,
      capacity: capacity as number
    })
  }

  // Log the event data to the console and return it for a flow.
  let stringResult = JSON.stringify(records);
  console.log(stringResult);
  return stringResult;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## Sample code: Validate speaker registrations

```TypeScript
function main(workbook: ExcelScript.Workbook, keys: string): string {
  // Get the first table in the "Transactions" worksheet.
  let table = workbook.getWorksheet('Transactions').getTables()[0];

  // Clear the existing formatting in the table.
  let range = table.getRangeBetweenHeaderAndTotal();
  range.clear(ExcelScript.ClearApplyTo.formats);

  // Compare the data in the table to the keys passed into the script.
  let keysObject = JSON.parse(keys) as EventData[];
  let speakerSlotsRemaining = keysObject.map(value => value.capacity);
  let overallMatch = true;

  // Iterate over every row looking for differences from the other worksheet.
  let rows = range.getValues();
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i];
    let [eventId, date, location, capacity] = row;
    let match = false;

    // Look at each key provided for a matching Event ID.
    for (let keyIndex = 0; keyIndex < keysObject.length; keyIndex++) {
      let event = keysObject[keyIndex];
      if (event.eventId === eventId) {
        match = true;
        speakerSlotsRemaining[keyIndex]--;
        // If there's a match on the event ID, look for things that don't match and highlight them.
        if (event.date !== date) {
          overallMatch = false;
          range.getCell(i, 1).getFormat()
            .getFill()
            .setColor("FFFF00");
        }
        if (event.location !== location) {
          overallMatch = false;
          range.getCell(i, 2).getFormat()
            .getFill()
            .setColor("FFFF00");
        }

        break;
      }
    }

    // If no matching Event ID is found, highlight the Event ID's cell.
    if (!match) {
      overallMatch = false;
      range.getCell(i, 0).getFormat()
        .getFill()
        .setColor("FFFF00");
    }
  }

  

  // Choose a message to send to the user.
  let returnString = "All the data is in the right order.";
  if (overallMatch === false) {
    returnString = "Mismatch found. Data requires your review.";
  } else if (speakerSlotsRemaining.find(remaining => remaining < 0)){
    returnString = "Event potentially overbooked. Please review."
  }

  console.log("Returning: " + returnString);
  return returnString;
}

// An interface representing a row of event data.
interface EventData {
  eventId: string
  date: number
  location: string
  capacity: number
}
```

## Power Automate flow: Check for inconsistencies across the workbooks

This flow extracts the event information from the first workbook and uses that data to validate the second workbook.

1. Create a new **Instant cloud flow**.
1. Select **Manually trigger a flow** and press **Create**.
1. Add a **New step** that uses the **Excel Online (Business)** connector with the **Run script** action. Use the following values for the action:
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: event-data.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Get event data

    :::image type="content" source="../../images/cross-reference-flow-1.png" alt-text="The completed Excel Online (Business) connector for the first script in Power Automate.":::

1. Add a second **New step** that uses the **Excel Online (Business)** connector with the **Run script** action. Use the following values for the action:
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: speaker-registration.xlsx ([selected with the file chooser](../../testing/power-automate-troubleshooting.md#select-workbooks-with-the-file-browser-control))
    * **Script**: Validate speaker registration

    :::image type="content" source="../../images/cross-reference-flow-2.png" alt-text="The completed Excel Online (Business) connector for the second script in Power Automate.":::
1. This sample uses Outlook as the email client. You could use any email connector Power Automate supports. Add a **New step** that uses the **Office 365 Outlook** connector and the **Send and email (V2)** action. Use the following values for the action:
    * **To**: Your test email account (or personal email)
    * **Subject**: Event validation results
    * **Body**: result (_dynamic content from **Run script 2**_)

    :::image type="content" source="../../images/cross-reference-flow-3.png" alt-text="The completed Office 365 Outlook connector in Power Automate.":::
1. Save the flow and try it out. You should receive an email saying "Mismatch found. Data requires your review." This indicates there are differences between rows in **speaker-registrations** and rows in **event-data**. Open **speaker-registrations** to see several highlighted cells where there are potential problems with the speaker registration listings.
