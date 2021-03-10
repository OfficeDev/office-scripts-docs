---
title: 'Seasons greetings'
description: 'Learn how to use Office Scripts to show a singing Christmas tree in Excel on the web.'
ms.date: 03/10/2021
localization_priority: Normal
---

# Seasons greetings

This is a script contributed by [Leslie Black](https://www.linkedin.com/in/lesblackconsultant/) in the spirit of the holiday season! It's a fun script that shows a singing Christmas tree in Excel on the web using Office Scripts.

Enjoy!

[![Watch the Seasons greetings script in action](../../images/community-seasons.png)](https://youtu.be/HBiGEkzmkgo "Seasons greetings script in action!")

## Script

```ts
/* By: Leslie Black  */

function main(workbook: ExcelScript.Workbook) {
  let HappyXmasTree = workbook.getWorksheet('HappyXmasTree')
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  FlashingStarandSmileFF0000(workbook) //red
  FlashingStarandSmileFFFF00(workbook) //yellow
  //FlashingStarandSmileFF0000(workbook) //red
  //FlashingStarandSmileFFFF00(workbook) //yellow
  //FlashingStarandSmileFF0000(workbook) //red
  //FlashingStarandSmileFFFF00(workbook) //yellow
  //FlashingStarandSmileFF0000(workbook) //red
  //FlashingStarandSmileFFFF00(workbook) //yellow
  //FlashingStarandSmileFF0000(workbook) //red
  //FlashingStarandSmileFFFF00(workbook) //yellow
  //FlashingStarandSmileFF0000(workbook) //red
  //FlashingStarandSmileFFFF00(workbook) //yellow
  //FlashingStarandSmileFF0000(workbook) //red
  //FlashingStarandSmileFFFF00(workbook) //yellow

  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  FlashingStarandSmileFF0000(workbook) //red
  FlashingStarandSmileFFFF00(workbook) //yellow
  Blink(workbook)

  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  //Unblink(workbook)

  //Blink(workbook)
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  //Unblink(workbook)

  //Blink(workbook)
  OuterEdgeFF0000(workbook) //red
  OuterEdgeFFFF00(workbook) //yellow
  Unblink(workbook)

  //Blink(workbook)
  //Unblink(workbook)






  function Blink(workbook: ExcelScript.Workbook) {
    //blink
    let selectedSheet = workbook.getWorksheet('HappyXmasTree');
    // Set fill color to C65911 for range HappyXmasTree!N16:Q17
    selectedSheet.getRange("N16:Q17")
      .getFormat()
      .getFill()
      .setColor("C65911");
    // Set fill color to C65911 for range HappyXmasTree!G16:J17
    selectedSheet.getRange("G16:J17")
      .getFormat()
      .getFill()
      .setColor("C65911");
  }

  function Unblink(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getWorksheet('HappyXmasTree');
    // Set fill color to FFFFFF for range HappyXmasTree!N16:N17
    selectedSheet.getRange("N16:N17")
      .getFormat()
      .getFill()
      .setColor("FFFFFF");
    // Set fill color to FFFFFF for range HappyXmasTree!O16:Q16
    selectedSheet.getRange("O16:Q16")
      .getFormat()
      .getFill()
      .setColor("FFFFFF");
    // Set fill color to FFFFFF for range HappyXmasTree!G16:H17
    selectedSheet.getRange("G16:H17")
      .getFormat()
      .getFill()
      .setColor("FFFFFF");
    // Set fill color to FFFFFF for range HappyXmasTree!I16:J16
    selectedSheet.getRange("I16:J16")
      .getFormat()
      .getFill()
      .setColor("FFFFFF");
    // Set fill color to FFFFFF for range HappyXmasTree!P17:Q17
    selectedSheet.getRange("P17:Q17")
      .getFormat()
      .getFill()
      .setColor("FFFFFF");
    // Set fill color to FFFFFF for range HappyXmasTree!J17
    selectedSheet.getRange("J17")
      .getFormat()
      .getFill()
      .setColor("FFFFFF");
  }

  function FlashingStarandSmileFF0000(workbook: ExcelScript.Workbook) {
    //red
    let selectedSheet = workbook.getWorksheet('HappyXmasTree');
    // Set fill color to FF0000 for range HappyXmasTree!L2:L6
    selectedSheet.getRange("L2:L6")
      .getFormat()
      .getFill()
      .setColor("FF0000");
    // Set fill color to FF0000 for range HappyXmasTree!K3:K5
    selectedSheet.getRange("K3:K5")
      .getFormat()
      .getFill()
      .setColor("FF0000");
    // Set fill color to FF0000 for range HappyXmasTree!M3:M5
    selectedSheet.getRange("M3:M5")
      .getFormat()
      .getFill()
      .setColor("FF0000");
    // Set fill color to FF0000 for range HappyXmasTree!N4
    selectedSheet.getRange("N4")
      .getFormat()
      .getFill()
      .setColor("FF0000");
    // Set fill color to FF0000 for range HappyXmasTree!J4
    selectedSheet.getRange("J4")
      .getFormat()
      .getFill()
      .setColor("FF0000");
    // Set fill color to 000000 for range HappyXmasTree!O26
    selectedSheet.getRange("O26")
      .getFormat()
      .getFill()
      .setColor("000000");
    // Set fill color to 000000 for range HappyXmasTree!I26
    selectedSheet.getRange("I26")
      .getFormat()
      .getFill()
      .setColor("000000") //black
  }

  function FlashingStarandSmileFFFF00(workbook: ExcelScript.Workbook) {
    //yellow
    let selectedSheet = workbook.getWorksheet('HappyXmasTree');
    // Set fill color to FF0000 for range HappyXmasTree!L2:L6
    selectedSheet.getRange("L2:L6")
      .getFormat()
      .getFill()
      .setColor("FFFF00");
    // Set fill color to FF0000 for range HappyXmasTree!K3:K5
    selectedSheet.getRange("K3:K5")
      .getFormat()
      .getFill()
      .setColor("FFFF00");
    // Set fill color to FF0000 for range HappyXmasTree!M3:M5
    selectedSheet.getRange("M3:M5")
      .getFormat()
      .getFill()
      .setColor("FFFF00");
    // Set fill color to FF0000 for range HappyXmasTree!N4
    selectedSheet.getRange("N4")
      .getFormat()
      .getFill()
      .setColor("FFFF00");
    // Set fill color to FF0000 for range HappyXmasTree!J4
    selectedSheet.getRange("J4")
      .getFormat()
      .getFill()
      .setColor("FFFF00");
    // Set fill color to 000000 for range HappyXmasTree!O26
    selectedSheet.getRange("O26")
      .getFormat()
      .getFill().clear
      //.setColor("000000");
    // Set fill color to 000000 for range HappyXmasTree!I26
    selectedSheet.getRange("I26")
      .getFormat()
      .getFill().clear
      //.setColor("000000")
  }

  //OuterEdgeClearFill()
  console.log('Routine finished')
  }

function OuterEdgeFFFF00(workbook: ExcelScript.Workbook) {
  // Set fill color to FFFF00 for range sheet!Q11
  let sheet = workbook.getWorksheet('HappyXmasTree')
  sheet.getRange("Q11")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!G11
  sheet.getRange("G11")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!R12
  sheet.getRange("R12")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!F12
  sheet.getRange("F12")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!S13
  sheet.getRange("S13")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!E13
  sheet.getRange("E13")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!T14
  sheet.getRange("T14")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!D14
  sheet.getRange("D14")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!C15
  sheet.getRange("C15")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!U15
  sheet.getRange("U15")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!T16:T17
  sheet.getRange("T16:T17")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!D16:D17
  sheet.getRange("D16:D17")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!C18
  sheet.getRange("C18")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!U18
  sheet.getRange("U18")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!T19
  sheet.getRange("T19")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!D19
  sheet.getRange("D19")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!L2:L6
  sheet.getRange("L2:L6")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!C21
  sheet.getRange("C21")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!U21
  sheet.getRange("U21")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!C23
  sheet.getRange("C23")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!U23
  sheet.getRange("U23")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!C25
  sheet.getRange("C25")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!U25
  sheet.getRange("U25")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!C27
  sheet.getRange("C27")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!U27
  sheet.getRange("U27")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!C29
  sheet.getRange("C29")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!U29
  sheet.getRange("U29")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!T30
  sheet.getRange("T30")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!D30
  sheet.getRange("D30")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!K3:K5
  sheet.getRange("K3:K5")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!M3:M5
  sheet.getRange("M3:M5")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!S31
  sheet.getRange("S31")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!E31
  sheet.getRange("E31")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!R32
  sheet.getRange("R32")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!F32
  sheet.getRange("F32")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!Q33
  sheet.getRange("Q33")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!G33
  sheet.getRange("G33")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!P34
  sheet.getRange("P34")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!H34
  sheet.getRange("H34")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!O35
  sheet.getRange("O35")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!I35
  sheet.getRange("I35")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!N36:N37
  sheet.getRange("N36:N37")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!J36:J37
  sheet.getRange("J36:J37")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!K37:M37
  sheet.getRange("K37:M37")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!N4
  sheet.getRange("N4")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!J4
  sheet.getRange("J4")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!K7
  sheet.getRange("K7")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!M7
  sheet.getRange("M7")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!N8
  sheet.getRange("N8")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!J8
  sheet.getRange("J8")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!O9
  sheet.getRange("O9")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!I9
  sheet.getRange("I9")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!P10
  sheet.getRange("P10")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  // Set fill color to FFFF00 for range sheet!H10
  sheet.getRange("H10")
    .getFormat()
    .getFill()
    .setColor("FFFF00");
}

function OuterEdgeFF0000(workbook: ExcelScript.Workbook) {
  // Set fill color to FFFF00 for range sheet!Q11
  let sheet = workbook.getWorksheet('HappyXmasTree')
  sheet.getRange("Q11")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!G11
  sheet.getRange("G11")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!R12
  sheet.getRange("R12")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!F12
  sheet.getRange("F12")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!S13
  sheet.getRange("S13")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!E13
  sheet.getRange("E13")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!T14
  sheet.getRange("T14")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!D14
  sheet.getRange("D14")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!C15
  sheet.getRange("C15")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!U15
  sheet.getRange("U15")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!T16:T17
  sheet.getRange("T16:T17")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!D16:D17
  sheet.getRange("D16:D17")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!C18
  sheet.getRange("C18")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!U18
  sheet.getRange("U18")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!T19
  sheet.getRange("T19")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!D19
  sheet.getRange("D19")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!L2:L6
  sheet.getRange("L2:L6")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!C21
  sheet.getRange("C21")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!U21
  sheet.getRange("U21")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!C23
  sheet.getRange("C23")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!U23
  sheet.getRange("U23")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!C25
  sheet.getRange("C25")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!U25
  sheet.getRange("U25")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!C27
  sheet.getRange("C27")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!U27
  sheet.getRange("U27")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!C29
  sheet.getRange("C29")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!U29
  sheet.getRange("U29")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!T30
  sheet.getRange("T30")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!D30
  sheet.getRange("D30")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!K3:K5
  sheet.getRange("K3:K5")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!M3:M5
  sheet.getRange("M3:M5")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!S31
  sheet.getRange("S31")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!E31
  sheet.getRange("E31")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!R32
  sheet.getRange("R32")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!F32
  sheet.getRange("F32")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!Q33
  sheet.getRange("Q33")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!G33
  sheet.getRange("G33")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!P34
  sheet.getRange("P34")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!H34
  sheet.getRange("H34")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!O35
  sheet.getRange("O35")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!I35
  sheet.getRange("I35")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!N36:N37
  sheet.getRange("N36:N37")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!J36:J37
  sheet.getRange("J36:J37")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!K37:M37
  sheet.getRange("K37:M37")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!N4
  sheet.getRange("N4")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!J4
  sheet.getRange("J4")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!K7
  sheet.getRange("K7")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!M7
  sheet.getRange("M7")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!N8
  sheet.getRange("N8")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!J8
  sheet.getRange("J8")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!O9
  sheet.getRange("O9")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!I9
  sheet.getRange("I9")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!P10
  sheet.getRange("P10")
    .getFormat()
    .getFill()
    .setColor("FF0000");
  // Set fill color to FF0000 for range sheet!H10
  sheet.getRange("H10")
    .getFormat()
    .getFill()
    .setColor("FF0000");
}
```
