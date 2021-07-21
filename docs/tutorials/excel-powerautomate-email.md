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

<img width="1470" alt="Screen Shot 2021-07-21 at 12 22 28 PM" src="https://user-images.githubusercontent.com/40217812/126524112-0d95424d-333d-4249-9fef-d6773b4563d8.png">

#### 1. Create Global Variables for my Script.
Whenever I write an Office Script I will set up a global variables for my workbook, any worksheets I'm using, and column numbers of important data.
