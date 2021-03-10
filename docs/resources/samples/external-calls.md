---
title: 'Email a chart image'
description: 'Learn how to use Office Scripts and Power Automate to extract and email an image of an Excel chart.'
ms.date: 03/10/2021
localization_priority: Normal
---

# External API calls from Office Scripts

Office Scripts allows [limited external API call support](https://docs.microsoft.com/en-us/office/dev/scripts/develop/external-calls) as documented on the official website. 

_Few things to note_:

* There is no way to sign-in or use OAuth2 type of authentication flows. All keys/credentials have to be hardcoded (or read from another source). 
* There is no infrastructure to store API credentials / keys. This will have to be managed by the user. 
* External calls may result in sensitive data being exposed to undesirable endpoints or external data to be brought into internal workbooks. Your admin can establish firewall protection against such calls. Be sure sure to check with local policies prior to relying on external calls. 
* If a script uses an API call, it will not function in Power Automate scenario. You'll have to use Power Automates HTTP action or equivalent actions to pull or push data from/to external service. 
* External API call involves asynchronous API syntax and requires slighly advanced knowledge of the way async communication works. 
* Be sure to check the data throughput prior to taking dependency. For instance pulling down entire external data-set may not be best option and instead pagination should be used to get data in chunks. 

## Prior knowledge/useful resources

* [REST API](https://en.wikipedia.org/wiki/Representational_state_transfer) - most likely way you'll use the API call. 
* Knowing how [`async` `await`](https://developer.mozilla.org/en-US/docs/Learn/JavaScript/Asynchronous/Async_await) works. 
* Knowing how [`fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API/Using_Fetch) works. 

## Steps 

1. Mark you `main` function as an asynchronous function by adding `async` prefix. `async function main(workbook: ExcelScript.Workbook)`...
1. Which type of API call are you making? `GET`, `POST`, `PUT`, `DELETE`, `PATCH`? Refer to REST API material to understand. 
1. Obtain the service API endpoint, authentication requirements, headers, etc. 
1. Define the input or output `interface` to help with code completion and development time verification. See video for details. 
1. Code/test/optimize. You can separte your API call routine to a separte function to make it re-usable from other part of script or for code-reuse from a different script (copy paste becomes much easier this way). 

## Scenario 

![Get repos info](git.png)


## Resources used in the demo 

1. [Get repos Github API reference](https://docs.github.com/en/free-pro-team@latest/rest/reference/repos#list-repositories-for-a-user)
1. API call output: Go to browser or any HTTP interface and type in: https://api.github.com/users/{USERNAME}/repos by replacing {USERNAME} with your Github ID. 
1. Information fetched: repo.name, repo.size, repo.owner.id, repo.license?.name

## Script used 

* [API Call from Office Script](APICall.ts)

## Video

[![API Call video](v_api.png)](https://youtu.be/fulP29J418E)
