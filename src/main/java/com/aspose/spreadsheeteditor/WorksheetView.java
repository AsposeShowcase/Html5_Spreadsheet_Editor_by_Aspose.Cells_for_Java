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
import org.primefaces.context.RequestContext;
import org.primefaces.event.CellEditEvent;
import org.primefaces.event.FileUploadEvent;
import org.primefaces.event.SelectEvent;

@Named(value = "worksheet")
@ViewScoped
public class WorksheetView implements Serializable {

    private com.aspose.cells.Workbook asposeWorkbook;
    private Format outputFormat = Format.XLSX;
    private String sourceUrl;
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
                this.sourceUrl = new URL(u).toString();
            } catch (MalformedURLException x) {
                sendMessageDialog("The specified URL is invalid", u);
            }
        }

        this.loadFromUrl();
    }

    public String getSourceUrl() {
        return sourceUrl;
    }

    public void setSourceUrl(String sourceUrl) {
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

    public void loadBlank() {
        this.asposeWorkbook = new com.aspose.cells.Workbook();
        this.cachedColumns = null;
        this.cachedRows = null;
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

    public void loadFromUrl() {
        if (this.sourceUrl == null) {
            return;
        }

        try (InputStream i = new URL(this.sourceUrl).openStream()) {
            this.asposeWorkbook = new com.aspose.cells.Workbook(i);
            System.out.println("File reloaded by Aspose.Cells at " + new java.util.Date().toString());

            sendMessage("Workbook loaded", this.sourceUrl.toString());
        } catch (MalformedURLException murlx) {
            sendMessage("The specified URL is invalid.", this.sourceUrl);
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

            try {
                c.setWidth(getAsposeWorksheet().getCells().getColumnWidthPixel(i));
            } catch (com.aspose.cells.CellsException | NullPointerException x) {
                c.setWidth(getDefaultColumnWidth());
            }

            this.cachedColumns.add(c);
        }
    }

    public int getDefaultColumnWidth() {
        try {
            return getAsposeWorksheet().getCells().getStandardWidthPixels();
        } catch (com.aspose.cells.CellsException | NullPointerException x) {
            return 64;
        }
    }

    public int getColumnWidth(int columnId) {
        if (isLoaded()) {
            return getAsposeWorksheet().getCells().getColumnWidthPixel(columnId);
        } else {
            return 0;
        }

    }

    public int getRowHeight(int rowId) {
        if (isLoaded()) {
            return getAsposeWorksheet().getCells().getRowHeightPixel(rowId);
        } else {
            return 0;
        }
    }

    public Cell getCell(int column, int row) {
        if (isLoaded()) {
            return createCellFromAsposeCell(getAsposeWorksheet().getCells().get(row, column));
        } else {
            return new Cell().setColumnId(column).setRowId(row);
        }
    }

    public Cell createCellFromAsposeCell(com.aspose.cells.Cell asposeCell) {
        Cell result = new Cell()
                .setColumnId(asposeCell.getColumn())
                .setRowId(asposeCell.getRow());

        try {
            asposeCell.calculate(true, null);
        } catch (com.aspose.cells.CellsException cx) {
            sendMessage("Cell recalculation failure", cx.getMessage());
        }

        try {
            result
                    .setFormula(asposeCell.getFormula())
                    .setValue(asposeCell.getStringValueWithoutFormat());
        } catch (Exception x) {
            sendMessage("Cell value error", x.getMessage());
        }

        StringBuilder style = new StringBuilder();

        try {
//            com.aspose.cells.Color cellBgColor = asposeCell.getStyle().getBackgroundColor();
//            style.append("background-color: rgb(")
//                    .append(cellBgColor.getR() & 0xFF)
//                    .append(",")
//                    .append(cellBgColor.getG() & 0xFF)
//                    .append(",")
//                    .append(cellBgColor.getB() & 0xFF)
//                    .append(");");
//            com.aspose.cells.Color cellFgColor = asposeCell.getStyle().getForegroundColor();
//            style.append("color: rgb(")
//                    .append(cellFgColor.getR())
//                    .append(",")
//                    .append(cellFgColor.getG())
//                    .append(",")
//                    .append(cellFgColor.getB())
//                    .append(");");
            com.aspose.cells.Font font = asposeCell.getStyle().getFont();
            style.append("font-family: '").append(font.getName()).append("';");

            if (asposeCell.getStyle().getFont().isItalic()) {
                style.append("font-style: italic;");
            }

            if (asposeCell.getStyle().getFont().isBold()) {
                style.append("font-weight: bold;");
            }

            switch (asposeCell.getStyle().getFont().getUnderline()) {
                case com.aspose.cells.FontUnderlineType.ACCOUNTING:
                case com.aspose.cells.FontUnderlineType.DASH:
                case com.aspose.cells.FontUnderlineType.DASHED_HEAVY:
                case com.aspose.cells.FontUnderlineType.DASH_DOT_DOT_HEAVY:
                case com.aspose.cells.FontUnderlineType.DASH_DOT_HEAVY:
                case com.aspose.cells.FontUnderlineType.DOTTED:
                case com.aspose.cells.FontUnderlineType.DOTTED_HEAVY:
                case com.aspose.cells.FontUnderlineType.DOT_DASH:
                case com.aspose.cells.FontUnderlineType.DOT_DOT_DASH:
                case com.aspose.cells.FontUnderlineType.DOUBLE:
                case com.aspose.cells.FontUnderlineType.DOUBLE_ACCOUNTING:
                case com.aspose.cells.FontUnderlineType.HEAVY:
                case com.aspose.cells.FontUnderlineType.SINGLE:
                case com.aspose.cells.FontUnderlineType.WAVE:
                case com.aspose.cells.FontUnderlineType.WAVY_DOUBLE:
                case com.aspose.cells.FontUnderlineType.WAVY_HEAVY:
                case com.aspose.cells.FontUnderlineType.WORDS:
                    style.append("text-decoration: underline;");
                    break;
            }

            switch (asposeCell.getStyle().getFont().getStrikeType()) {
                case com.aspose.cells.TextStrikeType.SINGLE:
                case com.aspose.cells.TextStrikeType.DOUBLE:
                    style.append("text-decoration: line-through;");
                    break;
            }

            switch (asposeCell.getStyle().getFont().getCapsType()) {
                case com.aspose.cells.TextCapsType.ALL:
                    style.append("text-transform: capitalize;");
                    break;
                case com.aspose.cells.TextCapsType.SMALL:
                    style.append("font-variant: small-caps;");
                    break;
            }

//            if (asposeCell.getStyle().getFont().getSize()) {
            style
                    .append("font-size: ")
                    .append(asposeCell.getStyle().getFont().getSize())
                    .append("pt;");
//            }

            switch (asposeCell.getStyle().getHorizontalAlignment()) {
                case com.aspose.cells.TextAlignmentType.GENERAL:
                case com.aspose.cells.TextAlignmentType.LEFT:
                    style.append("text-align: left;");
                    break;
                case com.aspose.cells.TextAlignmentType.RIGHT:
                    style.append("text-align: right;");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER_ACROSS:
                case com.aspose.cells.TextAlignmentType.CENTER:
                    style.append("text-align: center;");
                    break;
                case com.aspose.cells.TextAlignmentType.JUSTIFY:
                    style.append("text-align: justify;");
                    break;
            }

            switch (asposeCell.getStyle().getVerticalAlignment()) {
                case com.aspose.cells.TextAlignmentType.TOP:
                    style.append("vertical-align: top;");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER:
                    style.append("vertical-align: middle;");
                    break;
                case com.aspose.cells.TextAlignmentType.BOTTOM:
                    style.append("vertical-align: bottom;");
                    break;
            }

            int cellRotationAngle = asposeCell.getStyle().getRotationAngle();
            cellRotationAngle = 0; // TODO
            style.append("writing-mode: lr-bt ;")
                    .append("transform: rotate(-").append(cellRotationAngle).append("deg);")
                    .append("-webkit-transform: rotate(-").append(cellRotationAngle).append("deg);")
                    .append("-moz-transform: rotate(-").append(cellRotationAngle).append("deg);")
                    .append("-ms-transform: rotate(-").append(cellRotationAngle).append("deg);")
                    .append("-o-transform: rotate(-").append(cellRotationAngle).append("deg);");

            com.aspose.cells.Color cellTextColor = asposeCell.getStyle().getFont().getColor();
            style.append("color: rgb(")
                    .append(cellTextColor.getR() & 0xFF)
                    .append(",")
                    .append(cellTextColor.getG() & 0xFF)
                    .append(",")
                    .append(cellTextColor.getB() & 0xFF)
                    .append(");");

        } catch (Exception x) {
            sendMessage("Cell style error", x.getMessage());
        }

        result.setStyle(style.toString());

        if (asposeCell.getMergedRange() != null) {
            if (asposeCell.getColumn() == asposeCell.getMergedRange().getFirstColumn()) {
//                result.setColspan(asposeCell.getMergedRange().getColumnCount());
            }
            if (asposeCell.getRow() == asposeCell.getMergedRange().getFirstRow()) {
//                result.setRowspan(asposeCell.getMergedRange().getRowCount());
            }
        }

        return result;
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
                String columnName = com.aspose.cells.CellsHelper.columnIndexToName(j);
                try {
                    com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(i, j);

//                    if (c.getMergedRange() == null) {
                    rowBuilder.setCell(columnName, createCellFromAsposeCell(c));
//                    } else {
//                        if (c.getColumn() == c.getMergedRange().getFirstColumn() && c.getRow() == c.getMergedRange().getFirstRow()) {
//                            rowBuilder.setCell(columnName, createCellFromAsposeCell(c));
//                        }
//                    }
                } catch (Exception x) {
//                    x.printStackTrace();
                }
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
            this.cachedColumns = null;
            this.cachedRows = null;
        } catch (IOException iox) {
            sendMessage("Could not read the file from source", e.getFile().getFileName());
        } catch (Exception x) {
            sendMessage("Could not load the workbook", e.getFile().getFileName());
        }
    }

    public void onAddNewSheet() {
        if (isLoaded()) {
            try {
                int i = getAsposeWorkbook().getWorksheets().add();
                getAsposeWorkbook().getWorksheets().setActiveSheetIndex(i);
                this.cachedColumns = null;
                this.cachedRows = null;
            } catch (com.aspose.cells.CellsException cx) {
                sendMessage("New Worksheet", cx.getMessage());
            }
        }
    }

    public void onRemoveActiveSheet() {
        if (isLoaded()) {
            try {
                int i = getAsposeWorkbook().getWorksheets().getActiveSheetIndex();
                getAsposeWorkbook().getWorksheets().removeAt(i);
                if (getAsposeWorkbook().getWorksheets().getCount() == 0) {
                    int j = getAsposeWorkbook().getWorksheets().add();
                    getAsposeWorkbook().getWorksheets().setActiveSheetIndex(j);
                }
                this.cachedColumns = null;
                this.cachedRows = null;
            } catch (com.aspose.cells.CellsException cx) {
                sendMessage("Could not remove sheet", cx.getMessage());
            }
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

    private void sendMessageDialog(String summary, String details) {
        RequestContext.getCurrentInstance().showMessageInDialog(new FacesMessage(summary, details));
    }
}

