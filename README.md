## Html5 Spreadsheet Editor by Aspose for Java v1.2

Html5 Spreadsheet Editor is a web application that can view and edit Microsoft Excel files. It highlights commonly used features of [Aspose.Cells for Java](http://www.aspose.com/java/excel-component.aspx) and demonstrates how to use them to create, manipulate and render a spreadsheet in your Java application.

![](https://cloud.githubusercontent.com/assets/7696899/4931559/f0dace90-657a-11e4-97c4-2632118fa552.png)

## Highlights

### Opening files

You can open files using one of the following methods.

* Upload file from local computer
* Load files stored in online Dropbox storage
* Load file from a URL

#### Upload file from local computer

To upload a file form local computer, Click `File`->`Open from Computer` to open the `Browser` dialog. Go to your desired location of file and select it. Click `Open`. The is loaded in the editor.

#### Load files from Dropbox

To open file from Dropbox, Click `File`->`Open from Dropbox` to open Dropbox file chooser. If you are not already signed-in to Dropbox, it will require you to signin to your Dropbox account. Choose your desired file from Dropbox folder and click `Choose` at the bottom. Your selected file will be opened from Dropbox.

#### Load file from a URL

Files can be directly opened from URLs. This allows you to edit any publicly available file on Internet. To open the file append `?url=location` parameter with the value of your desired location while loading the editor.

`http://spreadsheet-editor.aspose.com/?url=http://example.com/Sample.xlsx`

### Create new files

To create a new empty spreadsheet click `New` in `File` menu.

### Saving files

After editing files, you will want to save those changes. The editor allows you to save the modified files to local computer. To save the files click `File` and than click `Save`. The modified file will be available for download.

### Supported formats

HTML5 Spreadsheet Editor supports following formats:

* XLS
* XLSX
* XLSM
* XLSB
* XLTX
* SpreadsheetML
* CVS
* OpenDocument

### Editing Cells

You can edit any cell by a double-click. When you double-click a cell, it switches to edit-mode. To cancel editing, press the ESC key. To swich to normal mode, press ENTER. You can also press TAB to move to the next cell. You can specify static text and numbers. Formulas are supported too. To enter a formula, start the cell value with an equal sign (=). For example `=SUM(A1:A5)`. All formulas supported by Microsoft Excel are supported by HTML5 Spreadsheet Editor too.

### Rows and Columns

Adding and removing rows is very easy. Each row starts with a number which is the row ID. Click on a row ID to select the entire row. Right-click and click `Add a Row Below` to add a new row right below the selected row. You can remove a row by following same method and clicking `Delete Row`.

### Working with Sheets

Microsoft Excel allows multiple sheets in a single file. HTML5 Spreadsheet Editor allows you to edit and switch between sheets. At the top right-hand corner there is a drop-down list of sheets. The selected sheet is the one which opened by the editor. To switch to another sheet, select it from the list. The plus `+` button adds a new sheet, and minus `-` button deletes the selected sheet.

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

By default, it will try to load `Aspose.Total.Java.lic` file from `src/main/resources/com/aspose/spreadsheeteditor` directory. Just copy the license file to this directory. You can change the default behavior by editing `WorksheetView.init()` method.

## Deployment

Spreadsheet Editor can be deployed on standard compliant Java application server that support CDI. It is tested on [Glassfish 4.0](http://glassfish.org/).
