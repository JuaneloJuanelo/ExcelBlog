# Hundreds of new JavaScript APIs are coming to Excel.

It's been 8 months since the last release of Excel JS APIs. During the period, we received lots of feedback regarding to the design and usage of APIs. We really appreciate your active paticipantion! The mission of providing richer and deeper capabilities of integrating with Excel via APIs remained unchagned and we spared no effort to it from now and then. And today we are excited to annouce more than one hundred of APIs coming to general availability in set 1.7!

## What's new in General Availablity?

- New Events in JavaScript. Excel JavaScript Events APIs provide a variety of event handlers that allow add-in to automatically run a designated function when a specific event occurs. You can design that function to perform whatever actions your scenario requires. From server side events supports the co-authoring scenarios, and add-in can tell the source of events from event parameters. In 1.7, we added support of events of onChanged, onSelectionChanged, onAdded and onActivated/onDeactivated.

- Customize chart elemetns. More chart related APIs were included in 1.7 to make the visulization customization more easier. With these APIs, add-in can change the appearance of chart itself, including font, boarder, title, line series, and point. And also create and customize appearance of chart related object, including axis, data label, legend and trendline.

- View and change the custome properties of a file, adding addtional information in the forms of key-value pairs.

- Access to document properties, e.g. author, subject, title. 

- More APIs to interact with Range and its format.
- Wide used style object with more than 20 properties. 
- Protect workbook from making changes.

## Paticipate in designing of new APIs in preview?
Excellent! Visit the Excel JavaScript API Open Specification, where you'll find documentation and code samples to help you learn more about the new API features now available in public preview.

Additionally, you can try out the new API features by using the built-in code snippets that are available in Script Lab. In Script Lab, you'll find samples that use preview APIs listed in the Preview APIs category.

## How to get a tryon?
Upgrade your Excel to the version 1801 (Build 9001.2171) or later. See [here](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets) for iOS/Mac/Online minimum version.

We encourage you to try out the new APIs and tell us what you think. You can post questions about the APIs on StackOverflow (with the office-js tag), make suggestions for the docs on GitHub, or suggest new API features on UserVoice.
