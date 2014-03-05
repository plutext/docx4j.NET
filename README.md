# docx4j.NET

.NET bindings for docx4j

## Motivation

Microsoft's Open XML SDK is the de facto way of working with docx/pptx/xlsx files in .NET environment.

However, it is quite a low level API; there is useful additional functionality in docx4j (see samples below), which can be utilised in a .NET environment.

Also, some docx4j users server side want to standardise on a single API, and so use docx4j client side (eg via VSTO AddIn).

## Users

You can install the NuGet package; see http://www.nuget.org/packages/docx4j.NET/

### Samples

Installing the NuGet package will add a dir src to your project; in src/samples you will see sample code for:
- docx to PDF
- docx to HTML
- interop with Open XML SDK
- mail merge (MERGEFIELD processing)
- content control data binding
All of those should run out of the box (provided you have set: Project Properties > Startup object)

For examples of how to do other stuff with docx4j, please see https://github.com/plutext/docx4j/tree/master/src/samples
Translating any of that code from Java to C# ought to be straightforward.

## Developers

You can clone this project.

The easiest way to add the needed references is still via NuGet.
