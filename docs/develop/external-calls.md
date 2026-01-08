---
title: External API call support in Office Scripts
description: Support and guidance for making external API calls in an Office Script.
ms.date: 01/08/2026
ms.localizationpriority: medium
---

# External API call support in Office Scripts

Scripts support calls to external services. Use these services to supply data and other information to your workbook.

> [!CAUTION]
> External calls may result in sensitive data being exposed to undesirable endpoints. Your admin can establish Information Rights Management (IRM) or firewall protection against such calls.

> [!IMPORTANT]
> Calls to external APIs can only be made through the Excel application, not through Power Automate [under normal circumstances](#external-calls-from-power-automate). External calls are also not supported for scripts stored on a SharePoint site.

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
* Document cookies, `localStorage`, and `sessionStorage` objects are not supported.
* External calls may result in sensitive data being exposed to undesirable endpoints, or external data to be brought into internal workbooks. Your admin can establish firewall protection against such calls. Be sure to check with local policies prior to relying on external calls.
* Be sure to check the amount of data throughput prior to taking a dependency. For instance, pulling down the entire external dataset may not be the best option and instead pagination should be used to get data in chunks.
* When using `fetch` for local network access in Chromium-based web browsers, users need to allow the script when prompted. If the script still fails, contact your IT administrator.

## Retrieve information with `fetch`

The [fetch API](https://developer.mozilla.org/docs/Web/API/Fetch_API) retrieves information from external services. It is an `async` API, so you need to adjust the `main` signature of your script. Make the `main` function `async`. You should also be sure to `await` the `fetch` call and `json` retrieval. This ensures those operations complete before the script ends.

Any JSON data retrieved by `fetch` must match an interface defined in the script. The returned value must be assigned to a specific type because [Office Scripts do not support the `any` type](typescript-restrictions.md#no-any-type-in-office-scripts). You should refer to the documentation for your service to see what the names and types of the returned properties are. Then, add the matching interface or interfaces to your script.

> [!NOTE]
> If you use `fetch` to call an external resource with a [Cross-Origin Resource Sharing (CORS)](https://developer.mozilla.org/docs/Web/HTTP/Guides/CORS) policy, make sure that the [`Access-Control-Allow-Origin`](https://developer.mozilla.org/docs/Web/HTTP/Reference/Headers/Access-Control-Allow-Origin) header of that external resource uses the `*` directive. If the `Access-Control-Allow-Origin` header uses a specific `<origin>` directive, then your `fetch` call from Office Scripts may fail. The origin of the Office Scripts runtime may change without notice.

The following script uses `fetch` to retrieve JSON data from the test server in the given URL. Note the `JSONData` interface to store the data as a matching type.

```typescript
async function main(workbook: ExcelScript.Workbook) {
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
* Samples in the [Use JSON to pass data to and from Office Scripts](use-json.md) article show how to pass data to and from `fetch` commands as JSON.
* The [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md) demonstrates the `fetch` command being used to retrieve records from the National Oceanic and Atmospheric Administration's Tides and Currents database.
* The second sample in [Add images to a workbook](../resources/samples/add-image-to-workbook.md) contains a `fetch` call to get an image from a website.

## Restrict external calls with Information Rights Management (IRM)

You can [apply IRM settings](/microsoft-365/compliance/apply-irm-to-a-list-or-library) to a workbook to prevent external calls being made by scripts. Disable the [**Copy**/**EXTRACT** policy](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) to prevent this behavior.

## External calls from Power Automate

[!INCLUDE [External calls in Power Automate](../includes/external-calls-power-automate.md)]

## See also

* [Use JSON to pass data to and from Office Scripts](use-json.md)
* [Using built-in JavaScript objects in Office Scripts](javascript-objects.md)
* [Use external fetch calls in Office Scripts](../resources/samples/external-fetch-calls.md)
* [Office Scripts sample scenario: Graph water-level data from NOAA](../resources/scenarios/noaa-data-fetch.md)
