## HTML5 Spreadsheet Editor by Aspose.Cells for Java v1.5

Html5 Spreadsheet Editor is a web application that can view and edit Microsoft Excel files. It highlights commonly used features of [Aspose.Cells for Java](http://www.aspose.com/java/excel-component.aspx) and demonstrates how to use them to create, manipulate and render a spreadsheet in your Java application.

![](https://cloud.githubusercontent.com/assets/9132329/5697305/5fe0d63c-9a0e-11e4-89d2-d14a90dd004e.png)

## What's new in 1.5

*	Insert cells
*	Delete cells
*	Clear cell formatting
*	Clear cell contents
*	Adjust column width
*	Adjust row height
*	Fixed the double scrollbar issue
*	Service-oriented modular design
*	Server-side cache for cell, row, column and worksheet objects
*	Differential updates to DOM to minimize the rendering time

## Highlights

### Opening files

You can open files using one of the following methods.

* Upload file from local computer
* Load files stored in online Dropbox storage
* Load file from a URL

#### Upload file from local computer

To upload a file form local computer, Click `Open from Computer` on `File` tab to open the `Browser` dialog. Go to your desired location of file and select it. Click `Open`. The file will be opened in the editor.

<img src="https://cloud.githubusercontent.com/assets/9132329/5697307/886486b2-9a0e-11e4-8e14-e97aeda7919a.png" width="50%" />

#### Load files from Dropbox

To open file from Dropbox, Click `Open from Dropbox` on `File` tab to open Dropbox file chooser. If you are not already signed-in to Dropbox, it will require you to signin to your Dropbox account. Choose your desired file from Dropbox folder and click `Choose` at the bottom. Your selected file will be opened from Dropbox.

<img src="https://cloud.githubusercontent.com/assets/9132329/5697310/b9434480-9a0e-11e4-9a46-feb8612a6e65.png" width="50%" />

#### Load file from a URL

Files can be directly opened from URLs. This allows you to edit any publicly available file on Internet. To open the file append `?url=location` parameter with the value of your desired location while loading the editor. For example:

`http://spreadsheet-editor.aspose.com/?url=http://example.com/Sample.xlsx`

<img src="https://cloud.githubusercontent.com/assets/9132329/5697316/d2d6629c-9a0e-11e4-9714-0d9215e33b5f.png" width="50%" />

### Create new files

To create a new empty spreadsheet click `New` on `File` tab.

<img src="https://cloud.githubusercontent.com/assets/9132329/5697305/5fe0d63c-9a0e-11e4-89d2-d14a90dd004e.png" width="50%" />

### Export and Save files

After editing files, you will want to save those changes. The editor allows you to export and download the modified files to local computer. To export the file click `Export as` on `File` tab and choose your desired format. The modified file will be exported for download. The following formats are supported for export:

* Excel 2007-2013 XLSX
* Excel 1997-2003 XLS
* Excel XLSM
* Excel XLSB
* Excel XLTX
* Excel XLTM
* SpreadsheetML
* Portable Document Format (PDF)
* OpenDocument Spreadsheet (ODS)

<img src="https://cloud.githubusercontent.com/assets/9132329/5697333/071a84ca-9a0f-11e4-9d25-b844247bbcf6.png" width="50%" />

### Supported formats

HTML5 Spreadsheet Editor can open files in the following formats:

* Excel 1997-2003 XLS
* Excel 2007-2013 XLSX
* XLSM
* XLSB
* XLTX
* SpreadsheetML
* CVS
* OpenDocument

### Editing Cells

You can edit any cell by a double-click. When you double-click a cell, it switches to edit-mode. To cancel editing, press the ESC key. To switch to normal mode, press ENTER. You can also press TAB to move to the next cell. You can specify static text and numbers. Formulas are supported too. To enter a formula, start the cell value with an equal sign (=). For example `=SUM(A1:A5)`. All formulas supported by Microsoft Excel are supported by HTML5 Spreadsheet Editor too.

<img src="https://cloud.githubusercontent.com/assets/9132329/5697338/21a78784-9a0f-11e4-831c-3ca27c5be757.png" width="50%" />

### Cell formatting

HTML5 Spreadsheet Editor supports rendering of the following text formatting:

* Bold
* Italic
* Underlines
* Font style
* Font size
* Text color
* Strike through
* Horizontal cell alignment: left, right, center, justified
* Color

Support for editing of the following formatting is available:

* Bold
* Italic
* Underlines
* Font style
* Font size
* Horizontal cell alignment: left, right, center, justified

<img src="https://cloud.githubusercontent.com/assets/9132329/5697344/437f477a-9a0f-11e4-9a9f-8a2261b14673.png" width="50%" />

### Rows and Columns

Adding and removing rows and columns is very easy. Click inside any cell and click `Add a Row Above` to add a new row right above the selected cell. You can remove a row by following same method and clicking `Delete Row`. You can:

* Add a row above the selected cell.
* Add a row below the selected cell.
* Add a column on the left of selected cell.
* Add a column on the right of selected cell.
* Delete the row including the selected cell.
* Delete the column including the selected cell.

<img src="https://cloud.githubusercontent.com/assets/9132329/5697351/5f4b2bfe-9a0f-11e4-9802-75ace6620494.png" width="50%" />

### Working with Sheets

Microsoft Excel allows multiple sheets in a single file. HTML5 Spreadsheet Editor allows you to edit and switch between sheets. On the Sheets tab we have a drop-down list of sheets. The selected sheet is the one which is opened by the editor. To switch to another sheet, select it from the list. The plus `+` button **adds a new sheet**, and minus `-` button **deletes the selected sheet**.

To **rename existing sheet**, click on the sheet name in the text box to edit it. When you are finished editing the name, press ENTER key, or click anywhere outside the box. The sheet will be renamed.

<img src="https://cloud.githubusercontent.com/assets/9132329/5697357/77a7bf14-9a0f-11e4-92ab-ccba33d86f49.png" width="50%" />

### Embed Anywhere

You can embed HTML5 Spreadsheet Editor in any website of your choice using IFRAME. Here is an example code:

```<
iframe src="http://spreadsheet-editor.aspose.com/?url=http://example.com/Sample.xlsx" width="800" height="600">
Your web browser does not support IFRAMEs
</iframe>
```

## Setup

It is very easy to manage the project using [NetBeans IDE](https://netbeans.org/). 

1. Download the source code.
2. Open the project in NetBeans IDE.
3. Click `Run`
4. Select `Glassfish` server as Application Server.

![](https://cloud.githubusercontent.com/assets/7696899/4933205/d49e9b94-6593-11e4-879d-5ef39755a232.png)

The project build process is managed using Maven. So you can prepare a WAR file from command line without any IDE. Use the following command to generate a WAR for deployment. The documentation of corresponding application server will help you how to deploy the generated WAR and its dependencies.

```
mvn package
```

## Aspose License

The project works without a license, with limitations. To remove limitations, you can acquire a free [temporary license](http://www.aspose.com/corporate/purchase/temporary-license.aspx) or [buy a full license](http://www.aspose.com/purchase/default.aspx).

By default, it will try to load `Aspose.Total.Java.lic` file from `src/main/resources/com/aspose/spreadsheeteditor` directory. Just copy the license file to this directory. You can change the default behavior by editing `AsposeLicense` class.

## Deployment

Spreadsheet Editor can be deployed on standard compliant Java application server that support CDI. It is tested on [Glassfish 4.0](http://glassfish.org/).
