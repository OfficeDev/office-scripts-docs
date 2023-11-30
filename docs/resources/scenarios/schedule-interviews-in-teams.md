---
title: Schedule interviews in Teams
description: Learn how to use Office Scripts to send a Teams meeting from Excel data.
ms.date: 11/30/2023
ms.localizationpriority: medium
---

# Office Scripts sample scenario: Schedule interviews in Teams

In this scenario, you're an HR recruiter scheduling interview meetings with candidates in Teams. You manage the interview schedule of candidates in an Excel file. You'll need to send the Teams meeting invite to both the candidate and interviewers. You then need to update the Excel file with the confirmation that Teams meetings have been sent.

The solution has three steps that are combined in a single Power Automate flow.

1. A script extracts data from a table and returns an array of objects as [JSON](https://www.w3schools.com/whatis/whatis_json.asp) data.
1. The data is then sent to the Teams **Create a Teams meeting** action to send invites.
1. The same JSON data is sent to another script to update the status of the invitation.

For more information about working with JSON, read [Use JSON to pass data to and from Office Scripts](../../develop/use-json.md).

## Scripting skills covered

* Power Automate flows
* Teams integration
* Table parsing

## Setup instructions

### Download the workbook

1. Download the sample workbook to your OneDrive.
    > [!div class="nextstepaction"]
    > [Download the sample workbook](hr-schedule.xlsx)

1. Open the workbook in Excel.

1. Change at least one of the email addresses to your own so that you receive an invite.

### Create the scripts

1. Under the **Automate** tab, select **New Script** and paste the following script into the editor. This will extract table data to schedule invites.

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

1. Name the script **Schedule Interviews** for the flow.

1. Create another new script with the following code. This will mark rows as invited.

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

1. Name the second script **Record Sent Invites** for the flow.

### Create the Power Automate flow

This flow run the interview scheduling scripts, send the Teams meetings, and record the activity back in the workbook.

1. Create a new **Instant cloud flow**.
1. Choose **Manually trigger a flow** and select **Create**.
1. In the flow builder, select the **+** button and **Add an action**. Use the **Excel Online (Business)** connector's **Run script** action. Complete the action with the following values.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: hr-interviews.xlsx *(Chosen through the file browser)*
    * **Script**: Schedule Interviews
    :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="The completed Run script action to get interview data from the workbook.":::

1. Add a action that uses the **Microsoft Teams** connector's **Create a Teams meeting** action. As you select dynamic content from the Excel connector, a **For each** block will be generated for your flow. Complete the connector with the following values.
    * **Subject**: Contoso Interview
    * **Message**: *Message* (dynamic content from **Run script**)
    * **Time zone**: Pacific Standard Time
    * **Start time**: *StartTime* (dynamic content from **Run script**)
    * **End time**: *FinishTime* (dynamic content from **Run script**)
    * **Calendar id**: Calendar
    * **Required attendees**: *CandidateEmail* ; *InterviewerEmail* (dynamic content from **Run script** - note the ';' separating the values)
    :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="The completed Teams action to schedule meetings.":::

1. In the same **For each** block, add another **Run script** action. Use the following values.
    * **Location**: OneDrive for Business
    * **Document Library**: OneDrive
    * **File**: hr-interviews.xlsx *(Chosen through the file browser)*
    * **Script**: Record Sent Invites
    * **invites**: *result* (dynamic content from **Run script**)
        * Press **[Switch input to entire array](../../testing/power-automate-troubleshooting.md#pass-entire-arrays-as-script-parameters)** first.
    :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="The completed Excel Online (Business) connector to record that invites have been sent.":::

1. Save the flow. The flow designer show look like the following image.

    :::image type="content" source="../../images/schedule-interviews-4.png" alt-text="A diagram of the completed flow that shows two steps leading to a For each control and two steps inside the For each control.":::

1. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.

## Training video: Send a Teams meeting from Excel data

[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8). His version uses a more robust script that handles changing columns and obsolete meeting times.
