---
title: 'Email a chart image'
description: 'Learn how to use Office Scripts and Power Automate to extract and email an image of an Excel chart.'
ms.date: 03/10/2021
localization_priority: Normal
---

# Output Excel table data as array of objects (as JSON) including hyperlink data for usage in Power Automate 

Often it is beneficial to extract Excel table data as array of objects (each item representing a row) in the form of a JSON using Office Scripts. This helps with extracting the data from Excel in the same format that is visible to the user. Columns such as date and date-time can then be fed into other systems using Power Automate flows. 

The second part of the demo sample involves doing the same as above, but this time selecting the hyperlink associated with one of the table columns. This allows data not readily available at the cell level to be surfaced in the array of object data (JSON). 

## Part-1 Input table data
![Input table](Table-Input.png) 

## Part-1 Output - Excel table data as JSON: array of rows 

```json
[{
	"Event ID": "E107",
	"Date": "2020-12-10",
	"Location": "Montgomery",
	"Capacity": "10",
	"Speakers": "Debra Berger"
}, {
	"Event ID": "E108",
	"Date": "2020-12-11",
	"Location": "Montgomery",
	"Capacity": "10",
	"Speakers": "Delia Dennis"
}, {
	"Event ID": "E109",
	"Date": "2020-12-12",
	"Location": "Montgomery",
	"Capacity": "10",
	"Speakers": "Diego Siciliani"
}, {
	"Event ID": "E110",
	"Date": "2020-12-13",
	"Location": "Boise",
	"Capacity": "25",
	"Speakers": "Gerhart Moller"
}, {
	"Event ID": "E111",
	"Date": "2020-12-14",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Speakers": "Grady Archie"
}, {
	"Event ID": "E112",
	"Date": "2020-12-15",
	"Location": "Fremont",
	"Capacity": "25",
	"Speakers": "Irvin Sayers"
}, {
	"Event ID": "E113",
	"Date": "2020-12-16",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Speakers": "Isaiah Langer"
}, {
	"Event ID": "E114",
	"Date": "2020-12-17",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Speakers": "Johanna Lorenz"
}]
```

## Part-2 Input table data
![Input table with hyperlink](table-hyper1.png) 
![Input table with hyperlink](table-hyper2.png) 

## Part-2 Output - Excel table data as JSON: array of rows 

```json
[{
	"Event ID": "E107",
	"Date": "2020-12-10",
	"Location": "Montgomery",
	"Capacity": "10",
	"Search link": "https://www.google.com/search?q=Montgomery",
	"Speakers": "Debra Berger"
}, {
	"Event ID": "E108",
	"Date": "2020-12-11",
	"Location": "Montgomery",
	"Capacity": "10",
	"Search link": "https://www.google.com/search?q=Montgomery",
	"Speakers": "Delia Dennis"
}, {
	"Event ID": "E109",
	"Date": "2020-12-12",
	"Location": "Montgomery",
	"Capacity": "10",
	"Search link": "https://www.google.com/search?q=Montgomery",
	"Speakers": "Diego Siciliani"
}, {
	"Event ID": "E110",
	"Date": "2020-12-13",
	"Location": "Boise",
	"Capacity": "25",
	"Search link": "https://www.google.com/search?q=Boise",
	"Speakers": "Gerhart Moller"
}, {
	"Event ID": "E111",
	"Date": "2020-12-14",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Search link": "https://www.google.com/search?q=salt+lake+city",
	"Speakers": "Grady Archie"
}, {
	"Event ID": "E112",
	"Date": "2020-12-15",
	"Location": "Fremont",
	"Capacity": "25",
	"Search link": "https://www.google.com/search?q=Fremont",
	"Speakers": "Irvin Sayers"
}, {
	"Event ID": "E113",
	"Date": "2020-12-16",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Search link": "https://www.google.com/search?q=salt+lake+city",
	"Speakers": "Isaiah Langer"
}, {
	"Event ID": "E114",
	"Date": "2020-12-17",
	"Location": "Salt Lake City",
	"Capacity": "20",
	"Search link": "https://www.google.com/search?q=salt+lake+city",
	"Speakers": "Johanna Lorenz"
}]
```


## Use this in Power Automate 

See a similar sample for how to use such a script in Power Auotmate: 

* Sample: https://github.com/sumurthy/officescripts-projects/tree/main/Chart%20and%20Table%20Images


## Office Scripts source code

1. [Return table data as array of objects/JSON](TableAsArrayOfObjects.ts)

To suit your needs/data, chage the `interface TableData` structure to match your table columns. Note that for column names that contains spaces, be sure to name your key within quotes as with `"Event ID"` in the sample. 

2. [Return table data as array of objects/JSON with hyperlink text](TableAsArrayOfObjectsWithHyperlink.ts)

Note that the script always extracts hyperlink from the 4th column (0-index) of the table. You can change that order or include multiple columns as hyperlink data by modofying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`

## Excel files used

[Excel file](Table-Data-With-Hyperlinks.xlsx)
