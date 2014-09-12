package com.aspose.spreadsheeteditor;

import com.aspose.cells.Workbook;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import javax.annotation.PostConstruct;
import javax.faces.application.FacesMessage;
import javax.faces.context.FacesContext;
import javax.faces.view.ViewScoped;
import javax.inject.Named;
import org.primefaces.event.CellEditEvent;
import org.primefaces.event.FileUploadEvent;
import org.primefaces.event.SelectEvent;

@Named(value = "worksheet")
@ViewScoped
public class WorksheetView implements Serializable {

    private com.aspose.cells.Workbook asposeWorkbook;
    private Format outputFormat = Format.XLSX;
    private URL sourceUrl;
    private ArrayList<Row> cachedRows;
    private ArrayList<Column> cachedColumns;
    private Row selectedRow;

    @PostConstruct
    private void init() {
        if (!com.aspose.cells.License.isLicenseSet()) {
            try (InputStream i = WorksheetView.class.getResourceAsStream("Aspose.Total.Java.lic")) {
                new com.aspose.cells.License().setLicense(i);
            } catch (IOException | com.aspose.cells.CellsException x) {
                x.printStackTrace();
            }
        }

        String u = FacesContext.getCurrentInstance().getExternalContext().getRequestParameterMap().get("url");
        if (u != null) {
            try {
                this.sourceUrl = new URL(u);
            } catch (MalformedURLException x) {
                sendMessage("The specified URL is invalid", u);
            }
        }

        if (this.sourceUrl != null) {
            this.load();
        }
    }

    public URL getSourceUrl() {
        return sourceUrl;
    }

    public void setSourceUrl(URL sourceUrl) {
        this.sourceUrl = sourceUrl;
    }

    public String getActiveSheet() {
        if (this.isLoaded()) {
            int i = getAsposeWorkbook().getWorksheets().getActiveSheetIndex();
            return getAsposeWorkbook().getWorksheets().get(i).getName();
        }
        return null;
    }

    public void setActiveSheet(String name) {
        int i = getAsposeWorkbook().getWorksheets().get(name).getIndex();
        getAsposeWorkbook().getWorksheets().setActiveSheetIndex(i);
        this.cachedColumns = null;
        this.cachedRows = null;
    }

    public boolean isLoaded() {
        return this.asposeWorkbook != null;
    }

    public Format getOutputFormat() {
        return this.outputFormat;
    }

    public void setOutputFormat(Format outputFormat) {
        this.outputFormat = outputFormat;
    }

    public void save() {
        int saveFormat = -1;

        switch (this.outputFormat) {
            case XLSX:
                saveFormat = com.aspose.cells.SaveFormat.XLSX;
                break;
            case XLS:
                saveFormat = com.aspose.cells.SaveFormat.EXCEL_97_TO_2003;
                break;
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            getAsposeWorkbook().save(out, saveFormat);
        } catch (Exception x) {
            throw new RuntimeException(x);
        }

        FacesContext.getCurrentInstance().getExternalContext().addResponseHeader("Content-Type", "application/octet-stream");
        FacesContext.getCurrentInstance().getExternalContext().addResponseHeader("Content-Disposition", "attachment; filename=Spreadsheet.xlsx");

        try (OutputStream responseStream = FacesContext.getCurrentInstance().getExternalContext().getResponseOutputStream()) {
            responseStream.write(out.toByteArray());
            responseStream.flush();
        } catch (IOException x) {
            throw new RuntimeException(x);
        }
    }

    private void load() {
        try (InputStream i = this.sourceUrl.openStream()) {
            this.asposeWorkbook = new com.aspose.cells.Workbook(i);
            System.out.println("File reloaded by Aspose.Cells at " + new java.util.Date().toString());

            sendMessage("Workbook loaded", this.sourceUrl.toString());
        } catch (IOException iox) {
            sendMessage("Sorry, there was a problem opening the specified file.", iox.getMessage());
        } catch (Exception x) {
            throw new RuntimeException(x);
        }

    }

    public com.aspose.cells.Workbook getAsposeWorkbook() {
        return this.asposeWorkbook;
    }

    private void cacheColumns() {
        this.cachedColumns = new ArrayList<>();
        int count = 0;
        try {
            count = getAsposeWorksheet().getCells().getMaxColumn() + 1;
        } catch (Exception x) {
        }
        count = count + 26 - (count % 26);

        for (int i = 0; i < count; i++) {
            Column c = new Column(i, com.aspose.cells.CellsHelper.columnIndexToName(i));
            this.cachedColumns.add(c);
        }
    }

    private void cacheRows() {
        if (this.cachedColumns == null) {
            this.cacheColumns();
        }

        this.cachedRows = new ArrayList<>();
        int rowCount = 0;
        try {
            rowCount = getAsposeWorksheet().getCells().getMaxRow() + 1;
        } catch (Exception x) {
        }
        rowCount = rowCount + 10 - (rowCount % 10);

        for (int i = 0; i < rowCount; i++) {
            Row.Builder rowBuilder = new Row.Builder().setId(i);

            for (int j = 0; j < this.cachedColumns.size(); j++) {
                Cell.Builder cellBuilder = new Cell.Builder();
                try {
                    com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(i, j);
                    c.calculate(true, null);
                    cellBuilder
                            .setFormula(c.getFormula())
                            .setHtml(c.getHtmlString())
                            .setValue(c.getStringValueWithoutFormat())
                            .setColumnId(j)
                            .setRowId(i);
                } catch (Exception x) {
//                    x.printStackTrace();
                }
                rowBuilder.setCell(com.aspose.cells.CellsHelper.columnIndexToName(j), cellBuilder.build());
            }

            this.cachedRows.add(rowBuilder.build());
        }
    }

    private com.aspose.cells.Worksheet getAsposeWorksheet() {
        return getAsposeWorkbook().getWorksheets().get(getAsposeWorkbook().getWorksheets().getActiveSheetIndex());
    }

    private com.aspose.cells.Worksheet getAsposeWorksheet(int index) {
        return getAsposeWorkbook().getWorksheets().get(index);
    }

    public List<String> getSheets() {
        ArrayList<String> list = new ArrayList<>();
        if (this.isLoaded()) {
            for (int i = 0; i < getAsposeWorkbook().getWorksheets().getCount(); i++) {
                list.add(String.valueOf(getAsposeWorkbook().getWorksheets().get(i).getName()));
            }
        }
        return list;
    }

    public List<Column> getColumns() {
        if (this.cachedColumns == null) {
            cacheColumns();
        }

        return this.cachedColumns;
    }

    public List<Row> getRows() {
        if (this.cachedRows == null) {
            this.cacheRows();
        }

        return this.cachedRows;
    }

    public void onFileUpload(FileUploadEvent e) {
        try (InputStream i = e.getFile().getInputstream()) {
            this.asposeWorkbook = new Workbook(i);
        } catch (IOException iox) {
            sendMessage("Could not read the file from source", e.getFile().getFileName());
        } catch (Exception x) {
            sendMessage("Could not load the workbook", e.getFile().getFileName());
        }
    }

    public void onCellEdit(CellEditEvent e) {
        Cell oldCell = (Cell) e.getOldValue();
        Cell newCell = (Cell) e.getNewValue();

        try {
            com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(newCell.getRowId(), newCell.getColumnId());
            if (newCell.getFormula() != null) {
                c.setFormula(newCell.getFormula(), null);
            } else {
                c.putValue(newCell.getValue(), true);
            }
            this.cachedRows = null;

            sendMessage("Cell Changed", String.format("Old: %s New: %s", oldCell.getValue(), newCell.getValue()));

        } catch (com.aspose.cells.CellsException x) {
            x.printStackTrace();
        }
    }

    public void onRowSelect(SelectEvent e) {
        this.selectedRow = (Row) e.getObject();
        sendMessage("Row selected", String.valueOf(this.selectedRow.getId()));
    }

    public void addRowAbove() {
        if (this.selectedRow == null) {
            sendMessage("No row selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().insertRows(this.selectedRow.getId(), 1, true);
            this.cachedRows = null;

            sendMessage("Added new above", String.valueOf(this.selectedRow.getId()));
        } catch (com.aspose.cells.CellsException cx) {
            cx.printStackTrace();
        }
    }

    public void addRowBelow() {
        if (this.selectedRow == null) {
            sendMessage("No row selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().insertRows(this.selectedRow.getId() + 1, 1, true);
            this.cachedRows = null;

            sendMessage("Added new below", String.valueOf(this.selectedRow.getId()));
        } catch (com.aspose.cells.CellsException cx) {
            cx.printStackTrace();
        }

    }

    public void deleteRow() {
        if (this.selectedRow == null) {
            sendMessage("No row selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().deleteRows(this.selectedRow.getId(), 1, true);
            this.cachedRows = null;

            sendMessage("Deleted row", String.valueOf(this.selectedRow.getId()));
        } catch (com.aspose.cells.CellsException cx) {
            cx.printStackTrace();
        }
    }

    public void setCell(int columnId, int rowId, Cell cell) {
        try {
            getAsposeWorksheet().getCells().get(rowId, columnId).putValue(cell.getValue());
            this.cachedRows = null;
        } catch (com.aspose.cells.CellsException x) {
            x.printStackTrace();
        }
    }

    private void sendMessage(String summary, String details) {
        FacesContext.getCurrentInstance().addMessage(null, new FacesMessage(summary, details));
    }
}
