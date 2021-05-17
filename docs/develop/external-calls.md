---
title: 'External API call support in Office Scripts'
description: 'Support and guidance for making external API calls in an Office Script.'
ms.date: 05/14/2021
localization_priority: Normal
---

# External API call support in Office Scripts

Script authors shouldn't expect consistent behavior when using [external APIs](https://developer.mozilla.org/docs/Web/API) during the platform's preview phase. As such, do not rely on external APIs for critical script scenarios.

Calls to external APIs can be only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate).

> [!CAUTION]
> External calls may result in sensitive data being exposed to undesirable endpoints. Your admin can establish firewall protection against such calls.

## Configure your script for external calls

External calls are [asynchronous](https://developer.mozilla.org/docs/Learn/JavaScript/Asynchronous/Async_await) and require that your script is marked as `async`. Add the `async` prefix to your `main` function and have it return a `Promise`, as shown here:

```typescript
async function main(workbook: ExcelScript.Workbook) : Promise <void>
```

> [!NOTE]
> Scripts that return other information can return a `Promise` of that type. For example, if your script needs to return an `Employee` object, the return signature would be `: Promise <Employee>`

You'll need to learn the external service's interfaces to make calls to that service. If you are using `fetch` or [REST APIs](https://wikipedia.org/wiki/Representational_state_transfer), you need to determine the JSON structure of the returned data. For both input to and output from your script, consider making an `interface` to match the needed JSON structures. This gives the script more type safety. You can see an example of this in [Using fetch from Office Scripts](../resources/samples/external-fetch-calls.md).

### Limitations with external calls from Office Scripts

* There is no way to sign in or use OAuth2 type of authentication flows. All keys and credentials have to be hardcoded (or read from another source).
* There is no infrastructure to store API credentials and keys. This will have to be managed by the user.
* Document cookies, `localStorage` and `sessionStorage` objects are not supported. 
* External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks. Your admin can establish firewall protection against such calls. Be sure to check with local policies prior to relying on external calls.
* Be sure to check the amount of data throughput prior to taking a dependency. For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.

## Working with `fetch`

The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services. It is an `async` API, so you need to adjust the `main` signature of your script. Make the `main` function `async` and have it return a `Promise<void>`. You should also be sure to `await` the `fetch` call and `json` retrieval. This ensures those operations complete before the script ends.

Any JSON data retrieved by `fetch` must match an interface defined in the script. The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts). You should refer to the documentation for your service to see what the names and types of the returned properties are. Then, add the matching interface or interfaces to your script.

The following script uses `fetch` to retrieve JSON data from the test server in the given URL. Note the `JSONData` interface to store the data as a matching type.

```TypeScript
async function main(workbook: ExcelScript.Workbook): Promise<void> {
  // Retrieve sample JSON data from a test server.
  let fetchResult = await fetch('https://jsonplaceholder.typicode.com/todos/1');

  // Convert the returned data to the expected JSON structure.
  let json : JSONData = await fetchResult.json();

  // Display the content in a readable format.
  console.log(JSON.stringify(json));
}

/**
 * An interface that matches the returned JSON structure.
 * The property names match exactly.
 */
interface JSONData {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}
```

### Other `fetch` samples

* The [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md) sample shows how to get basic information about a user's GitHub repositories.
* The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the fetch command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.

## External calls from Power Automate

Any external API calls fail when a script is run with Power Automate. This is a behavioral difference between running a script through the Excel client and through Power Automate. Be sure to check your scripts for such references before building them into a flow.

You'll have to use [HTTP with Azure AD](/connectors/webcontents/) or other equivalent actions to pull data from or push it to an external service.

> [!WARNING]
> External calls made through the Power Automate [Excel Online connector](/connectors/excelonlinebusiness) fail in order to help uphold existing data loss prevention policies. However, scripts that are run through Power Automate are done so outside of your organization, and outside of your organization's firewalls. For additional protection from malicious users in this external environment, your admin can control the use of Office Scripts. Your admin can either disable the Excel Online connector in Power Automate or turn off Office Scripts for Excel on the web through the [Office Scripts administrator controls](/microsoft-365/admin/manage/manage-office-scripts-settings).

## See also

* [Using built-in JavaScript objects in Office Scripts](javascript-objects.md)
* [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md)
* [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md)
