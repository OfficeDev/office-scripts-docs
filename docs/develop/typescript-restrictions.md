---
title: 'TypeScript restrictions in Office Scripts'
description: 'The specifics of the TypeScript compiler and linter used by the Office Scripts Code Editor.'
ms.date: 05/24/2021
localization_priority: Normal
---

# TypeScript restrictions in Office Scripts

Office Scripts use the TypeScript language. For the most part, any TypeScript or JavaScript code will work in Office Scripts. However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.

## No 'any' type in Office Scripts

Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred. However, Office Scripts requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any). Both explicit and implicit `any` are not allowed in Office Scripts. These cases are reported as errors.

### Explicit `any`

You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let value: any;`). The `any` type causes issues when processed by Excel. For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`. You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="The explicit `any` message in the Code Editor's hover text.":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="The explicit `any` error in the console window.":::

In the previous screenshot, `[2, 14] Explicit Any is not allowed` indicates that line #2, column #14 defines `any` type. This helps you locate the error.

To get around this issue, always define the type of the variable. If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html). This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).

### Implicit `any`

TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined. If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="The implicit `any` message in the Code Editor's hover text.":::

The most common case on any implicit `any` is in a variable declaration, such as `let value;`. There are two ways to avoid this:

* Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).
* Explicitly type the variable (`let value: number;`)

## No inheriting Office Script classes or interfaces

Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces. In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.

## Incompatible TypeScript functions

Office Scripts APIs cannot be used in the following:

* [Generator functions](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## `eval` is not supported

The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.

## Restricted identifers

The following words can't be used as identifiers in a script. They are reserved terms.

* `Excel`
* `ExcelScript`
* `console`

## Only arrow functions in array callbacks

Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods. You cannot pass any sort of identifier or "traditional" function to these methods.

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## Performance warnings

The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues. The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).

## External API calls

See [External API call support in Office Scripts](external-calls.md) for more information.

## See also

* [Scripting fundamentals for Office Scripts in Excel on the web](scripting-fundamentals.md)
* [Improve the performance of your Office Scripts](web-client-performance.md)
