# Sending Emails Using Office Script & Power Automate
By: Michael Huskey

## Background

### What was the problem I was trying to solve 🧐?
Because of the pandemic our team was spread out across the world. Part of the team in Southeastern Michigan, another group in Toronto, and the other part of the team in the UK. 

This made device testing for our team a group effort, but it was only part time. Everybody who was testing devices also had another full time job to do for the team as well. This meant that many times team members would forget to run their tests (myself included).

The old way to make sure everyone completed their tests was by checking an Excel File and seeing which tabs had not been filled out, and then sending corresponding Slack Messages to the team members.

### The Solution 💡 

> Microsoft Forms + Office Script + Power Automate + Outlook

Using the tools listed above I created a solution that would be able to automatically send reminder emails to teammates, which dramatically decreased our testing deliquency, eliminated the need for a team member to check and sped up the process of testing software versions.

## What I am going to show you in this Tutorial

I don't think my management would be too happy if I showed everything that I did to speed up our internal process, but I can show you this one part that did make all the difference and that is using `Office Script + Power Automate` to send out automated email reminders.

A look at the data. (Check Image Below 👇)

###<img width="1470" alt="Screen Shot 2021-07-21 at 2 33 46 PM" src="https://user-images.githubusercontent.com/40217812/126541384-16d1d111-57b7-4123-9cca-3f0e0c286fa7.png">

 1. Create & Instantiate Global Variables for my Script
Whenever I write an Office Script I will set up a global variables for my workbook, any worksheets I'm using, and column numbers of important data.
``` javascript
// Global Variables for Workbook & Sheets
let gWorkbook : ExcelScript.Workbook;
let sheet1 : ExcelScript.Worksheet;

// Column Variables

let name_COL = 0;
let email_COL = 1;
let device_COL = 2;
let latest_COL = 3;

let releaseNo_COL = 5;
let releaseDate_COL = 6;


function main(workbook: ExcelScript.Workbook) {
	// Instantiate Global Workbook & Sheets Variables
	gWorkbook = workbook
	sheet1 = gWorkbook.getWorksheet('Sheet1');

}
```
### 2. Create the Interface Object 
This is the data that will become accessible in Power Automate in steps after your script runs.
``` javascript
interface OutputData {
	name: string;
	emai: string;
	device: string;
	versionNo: string;
}
```
### 3. Set your Interface Object as what your `main` function will return and instantiate it in the main function.

``` javascript
function main(workbook: ExcelScript.Workbook): OutputData[] {
	// Instantiate Global Workbook & Sheets Variables
	gWorkbook = workbook
	sheet1 = gWorkbook.getWorksheet('Sheet1');

	var outputData : OutputData[] = [];

}
```

### 4. Write the rest of the script
I go through the tester data and if `releaseDate` - `latestTest` is greater than `daysDeliquent` it will add this user and their corresponding data to the output array.
``` javascript
// Global Variables for Workbook & Sheets
let gWorkbook: ExcelScript.Workbook;
let sheet1: ExcelScript.Worksheet;

// Column Variables

var name_COL = 0;
let email_COL = 1;
let device_COL = 2;
let latest_COL = 3;

let releaseNo_COL = 5;
let releaseDate_COL = 6;


function main(workbook: ExcelScript.Workbook): OutputData[] {
	// Instantiate Global Workbook & Sheets Variables
	gWorkbook = workbook
	sheet1 = gWorkbook.getWorksheet('Sheet1');

	var outputData : OutputData[] = [];
	
	// Main Code
	let latestRelease = GetLatestRelease();
	let releaseNum = latestRelease[0];
	let releaseDate = parseInt(latestRelease[1]) ;
	
	let testerData = GetTesterData().slice(1);

	let daysDeliquent = 3; // number days a tester has between release date and when they do a test.

	for(var i = 0; i < testerData.length; i++){
		var rowData = testerData[i];
		var name = rowData[name_COL];
		var email = rowData[email_COL];
		var device = rowData[device_COL];
		var lastTest = rowData[latest_COL];
		if(releaseDate - lastTest > daysDeliquent){
			var json = {name: name, email: email, device: device, release:releaseNum}
			outputData.push(json);
		}
	}
	console.log(outputData);
	return outputData;
}

function GetTesterData(){
	let length = sheet1.getUsedRange().getValues().length;
	let testerArray = removeBlanks(sheet1.getRangeByIndexes(0, name_COL, length, 4).getValues());
	return testerArray
}



function GetLatestRelease() {
	let length = sheet1.getUsedRange().getValues().length;
	let releaseArray = removeBlanks(sheet1.getRangeByIndexes(0, releaseNo_COL, length, 2).getValues());
	return releaseArray.pop()
}

/**
 * Will remove blank rows from the inputted array
 */
function removeBlanks(arr: (string | number | boolean)[][]) {
	let output: (string | number | boolean)[][] = [];
	for (var i = 0; i < arr.length; i++) {
		var rowData = arr[i];
		if (rowData[0] != "") {
			output.push(rowData);
		}
	}
	return output
}

/**
 * This is the data that will be accessible in the Power Automate Flow
 */
interface OutputData {
	name: string;
	emai: string;
	device: string;
	versionNo: string;
}
```
### 5. Open Power Auto
<img width="644" alt="Screen Shot 2021-07-21 at 2 57 39 PM" src="https://user-images.githubusercontent.com/40217812/126544421-4b57ad3a-2e90-4014-9d9f-50bbfb8ed856.png">
