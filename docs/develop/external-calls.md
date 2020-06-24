---
title: 'External API call support in Office Scripts'
description: 'Support and guidance for making external API calls in an Office Script.'
ms.date: 06/24/2020
localization_priority: Normal
---

# External API call support in Office Scripts

The Office Scripts platform doesn't support calls to [external APIs](https://developer.mozilla.org/docs/Web/API). However, these calls can be run under the right circumstances. Scripts authors shouldn't expect consistent behavior when using external APIs during the platform's preview phase.

## External calls from Power Automate

Any external API calls fail when a script is run with Power Automate. This is a behavioral difference between running a script through the Excel client and through Power Automate. Be sure to check your scripts for such references before building them into a flow.

> [!WARNING]
> The mechanism used to block external API calls from the [Excel Online connector] in Power Automate is less secure than the rest of the platform. This is a known issue during the Office Scripts preview phase. It will be addressed before Office Scripts is turned on by default for organizations. If your data is highly sensitive and you're concerned with potential external calls from scripts transmitting data to external sources, your admin should either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web.

## Definition files for external APIs

The definition files for external APIs aren't included with Office Scripts. The use of such APIs generates compile-time errors for missing definitions. The APIs still run, as shown in the following script:

```typescript
async function main(context: Excel.RequestContext): Promise <void> {
  /* The following line of code generates the error:
   * "Cannot find name 'fetch'".
   * It will still run and return the JSON from the testing service.
   */
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');
  let json = await fetchResult.json();

  // Displays the content from https://jsonplaceholder.typicode.com/todos/1
  console.log(JSON.stringify(json));
}
```

> [!IMPORTANT]
> Some external API calls have inconsistent behavior in Office Scripts. The JavaScript runtime may close before the API call completes (or its `Promise` is fully resolved). Do not rely on external APIs for critical script scenarios.
