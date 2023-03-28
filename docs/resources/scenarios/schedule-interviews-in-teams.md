---
title: Schedule interviews in Teams
description: Learn how to use Office Scripts to send a Teams meeting from Excel data.
ms.topic: sample
ms.date: 06/29/2021
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

## Sample Excel file

Download the file [hr-schedule.xlsx](hr-schedule.xlsx) used in this solution and try it out yourself! Be sure to change at least one of the email addresses so that you receive an invite.

## Sample code: Extract table data to schedule invites

Add this script to your script collection. Name it **Schedule Interviews** for the flow.

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

## Sample code: Mark rows as invited

Add this script to your script collection. Name it **Record Sent Invites** for the flow.

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

## Sample flow: Run the interview scheduling scripts and send the Teams meetings

1. Create a new **Instant cloud flow**.
1. Choose **Manually trigger a flow** and select **Create**.
1. Add a **New step** that uses the **Excel Online (Business)** connector and the **Run script** action. Complete the connector with the following values.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **File**: hr-interviews.xlsx *(Chosen through the file browser)*
    1. **Script**: Schedule Interviews
    :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="Screenshot of the completed Excel Online (Business) connector to get interview data from the workbook in Power Automate.":::
1. Add a **New step** that uses the **Create a Teams meeting** action. As you select dynamic content from the Excel connector, an **Apply to each** block will be generated for your flow. Complete the connector with the following values.
    1. **Calendar id**: Calendar
    1. **Subject**: Contoso Interview
    1. **Message**: **Message** (the Excel value)
    1. **Time zone**: Pacific Standard Time
    1. **Start time**: **StartTime** (the Excel value)
    1. **End time**: **FinishTime** (the Excel value)
    1. **Required attendees**: **CandidateEmail** ; **InterviewerEmail** (the Excel values)
    :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="Screenshot of the completed Teams connector to schedule meetings in Power Automate.":::
1. In the same **Apply to each** block, add another **Excel Online (Business)** connector with the **Run script** action. Use the following values.
    1. **Location**: OneDrive for Business
    1. **Document Library**: OneDrive
    1. **File**: hr-interviews.xlsx *(Chosen through the file browser)*
    1. **Script**: Record Sent Invites
    1. **invites**: **result** (the Excel value)
    :::image type="content" source="../../images/schedule-interviews-3.png" alt-text="Screenshot of the completed Excel Online (Business) connector to record that invites have been sent in Power Automate.":::
1. Save the flow and try it out. Use the **Test** button on the flow editor page or run the flow through your **My flows** tab. Be sure to allow access when prompted.

## Training video: Send a Teams meeting from Excel data

[Watch Sudhi Ramamurthy walk through a version of this sample on YouTube](https://youtu.be/HyBdx52NOE8). His version uses a more robust script that handles changing columns and obsolete meeting times.
