## Spreadsheet Editor

Spreadsheet Editor is a web application that can view and edit Excel workbook documents. It highlights commonly used features of [Aspose.Cells for Java](http://www.aspose.com/java/excel-component.aspx) and demonstrates how to use them to create, manipulate and render a spreadsheet.

![](http://www.aspose.com/blogs/wp-content/uploads/2014/09/2.png)

## Features

* Open a spreadsheet from any public URL.
* View a spreadsheet.
* View formated text in Cells.
* Edit cell values.
* Edit cell formulas.
* Add and delete rows.
* Show list of available worksheets.
* Switch between worksheets.
* Download the modified spreadsheet document.

## Upcoming features

Aspose.Cells for Java is very rich library and can be used to edit all aspects of an Excel file. Spreadsheet Editor is currently using only a small subset of Aspose.Cells for Java features and will expand in future releases. Here is a list of some prominent features that are expected in next versions of Spreadsheet Editor:

* Add and delete columns.
* Start with an empty spreadsheet.
* Upload file for editing.
* Edit text formating.
* View and edit cell styles.
* View merged cells.
* Merge and unmerge cells.
* Add and delete sheets.

## Setup

The project build process is managed using Maven. Run the following command to prepare a WAR file.

```
mvn package
```

## Aspose License

In addition to building from source code. You may need to add the license file for Aspose.Cells for Java component. By default, it will try to load `Aspose.Total.Java.lic` file from `src/main/resources/com/aspose/spreadsheeteditor` directory. You can change the default behavior by editing `WorksheetView.init()` method.

## Deployment

Spreadsheet Editor can be deployed on standard compliant Java application server that support CDI. It is tested on [Glassfish 4.0](http://glassfish.org/).
