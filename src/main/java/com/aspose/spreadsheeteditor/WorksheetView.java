package com.aspose.spreadsheeteditor;

import com.aspose.cells.Workbook;
import java.io.ByteArrayInputStream;
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
import org.primefaces.event.ColumnResizeEvent;
import org.primefaces.event.FileUploadEvent;
import org.primefaces.model.DefaultStreamedContent;
import org.primefaces.model.StreamedContent;

@Named(value = "worksheet")
@ViewScoped
public class WorksheetView implements Serializable {

    private com.aspose.cells.Workbook asposeWorkbook;
    private Format outputFormat = Format.XLSX;
    private String sourceUrl;
    private ArrayList<Row> cachedRows;
    private ArrayList<Column> cachedColumns;
    private int currentColumnId;
    private int currentRowId;
    private String currentCellClientId;
    private boolean boldOptionEnabled;
    private boolean italicOptionEnabled;
    private boolean underlineOptionEnabled;
    private String fontSelectionOption;
    private int fontSizeOption;
    private String alignSelectionOption;

    static {
        AsposeLicense.load();
    }

    @PostConstruct
    private void init() {
        String requestedSourceUrl = FacesContext.getCurrentInstance().getExternalContext().getRequestParameterMap().get("url");
        if (requestedSourceUrl != null) {
            try {
                this.sourceUrl = new URL(requestedSourceUrl).toString();
                this.loadFromUrl();
            } catch (MalformedURLException x) {
                sendMessageDialog("The specified URL is invalid", requestedSourceUrl);
            }
        }
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

    /**
     * Switch to `name`d sheet. If there does not exist any sheet with the given
     * name, the active sheet is renamed to the given name.
     *
     * The rule is derived from the following use-case.
     *
     * If the user select a sheet from drop-down menu, this means that the sheet
     * already exist. So we can switch to that sheet. If the sheet does not
     * exist, we can say that the user has not selected it from drop-down but
     * directly modified the name of existing in the input text box.
     *
     * @param name Worksheet name
     */
    public void setActiveSheet(String name) {
        com.aspose.cells.Worksheet w = getAsposeWorkbook().getWorksheets().get(name);
        if (w != null) {
            int i = w.getIndex();
            getAsposeWorkbook().getWorksheets().setActiveSheetIndex(i);
        } else {
            getAsposeWorksheet().setName(name);
        }

        purge();
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
        purge();
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

    public StreamedContent getOutputFile(int saveFormat) {
        byte[] buf;
        String ext = null;
        switch (saveFormat) {
            case com.aspose.cells.SaveFormat.EXCEL_97_TO_2003:
                ext = "xls";
                break;
            case com.aspose.cells.SaveFormat.XLSX:
                ext = "xlsx";
                break;
            case com.aspose.cells.SaveFormat.XLSM:
                ext = "xlsm";
                break;
            case com.aspose.cells.SaveFormat.XLSB:
                ext = "xlsb";
                break;
            case com.aspose.cells.SaveFormat.XLTX:
                ext = "xltx";
                break;
            case com.aspose.cells.SaveFormat.XLTM:
                ext = "xltm";
                break;
            case com.aspose.cells.SaveFormat.SPREADSHEET_ML:
                ext = "xml";
                break;
            case com.aspose.cells.SaveFormat.PDF:
                ext = "pdf";
                break;
            case com.aspose.cells.SaveFormat.ODS:
                ext = "ods";
                break;
        }

        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            getAsposeWorkbook().save(out, saveFormat);
            buf = out.toByteArray();
        } catch (Exception x) {
            throw new RuntimeException(x);
        }

        return new DefaultStreamedContent(new ByteArrayInputStream(buf), "application/octet-stream", "Spreadsheet." + ext);
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
//            style.append("background-color:")
//                    .append(asposeColorToCssColor(cellBgColor, false))
//                    .append(";");

            com.aspose.cells.Color cellFgColor = asposeCell.getStyle().getForegroundColor();
            style.append("background-color:")
                    .append(asposeColorToCssColor(cellFgColor, false))
                    .append(";");

            com.aspose.cells.Font font = asposeCell.getStyle().getFont();
            style.append("font-family:'").append(font.getName()).append("';");

            if (asposeCell.getStyle().getFont().isItalic()) {
                result.addClass("i").setItalic(true);
            }

            if (asposeCell.getStyle().getFont().isBold()) {
                result.addClass("b").setBold(true);
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
                    result.addClass("u").setUnderline(true);
                    break;
            }

            switch (asposeCell.getStyle().getFont().getStrikeType()) {
                case com.aspose.cells.TextStrikeType.SINGLE:
                case com.aspose.cells.TextStrikeType.DOUBLE:
                    result.addClass("sts");
                    break;
            }

            switch (asposeCell.getStyle().getFont().getCapsType()) {
                case com.aspose.cells.TextCapsType.ALL:
                    result.addClass("uc");
                    break;
                case com.aspose.cells.TextCapsType.SMALL:
                    result.addClass("sc");
                    break;
            }

//            if (asposeCell.getStyle().getFont().getSize()) {
            style
                    .append("font-size:")
                    .append(asposeCell.getStyle().getFont().getSize())
                    .append("pt;");
//            }

            switch (asposeCell.getStyle().getHorizontalAlignment()) {
                case com.aspose.cells.TextAlignmentType.GENERAL:
                case com.aspose.cells.TextAlignmentType.LEFT:
                    result.addClass("al");
                    break;
                case com.aspose.cells.TextAlignmentType.RIGHT:
                    result.addClass("ar");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER_ACROSS:
                case com.aspose.cells.TextAlignmentType.CENTER:
                    result.addClass("ac");
                    break;
                case com.aspose.cells.TextAlignmentType.JUSTIFY:
                    result.addClass("aj");
                    break;
            }

            switch (asposeCell.getStyle().getVerticalAlignment()) {
                case com.aspose.cells.TextAlignmentType.TOP:
                    result.addClass("at");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER:
                    result.addClass("am");
                    break;
                case com.aspose.cells.TextAlignmentType.BOTTOM:
//                    style.append("vertical-align: bottom;");
                    result.addClass("ab");
                    break;
            }

//            int cellRotationAngle = asposeCell.getStyle().getRotationAngle();
//            cellRotationAngle = 0; // TODO
//            style.append("writing-mode: lr-bt ;")
//                    .append("transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-webkit-transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-moz-transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-ms-transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-o-transform: rotate(-").append(cellRotationAngle).append("deg);");
            com.aspose.cells.Color cellTextColor = asposeCell.getStyle().getFont().getColor();
            style.append("color:")
                    .append(asposeColorToCssColor(cellTextColor, true))
                    .append(";");

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

    private String asposeColorToCssColor(com.aspose.cells.Color color, boolean emptyIsBlack) {
        int r, g, b;

        if (color.isEmpty()) {
            if (emptyIsBlack) {
                r = g = b = 0;
            } else {
                r = g = b = 255;
            }
        } else {
            r = color.getR() & 0xFF;
            g = color.getG() & 0xFF;
            b = color.getB() & 0xFF;
        }

        return new StringBuffer("rgb(")
                .append(r & 0xFF)
                .append(",")
                .append(g & 0xFF)
                .append(",")
                .append(b & 0xFF)
                .append(")").toString();
    }

    private void cacheRows() {
        if (this.cachedColumns == null) {
            this.cacheColumns();
        }

        this.cachedRows = new ArrayList<>();
        int rowCount = 20;
        try {
            rowCount = 20 + getAsposeWorksheet().getCells().getMaxRow() + 1;
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

    public void applyCellFormatting() {
        com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(currentRowId, currentColumnId);
        com.aspose.cells.Style s = c.getStyle();

        s.getFont().setBold(boldOptionEnabled);
        s.getFont().setItalic(italicOptionEnabled);
        s.getFont().setUnderline(underlineOptionEnabled ? com.aspose.cells.FontUnderlineType.SINGLE : com.aspose.cells.FontUnderlineType.NONE);
        s.getFont().setName(fontSelectionOption);
        s.getFont().setSize(fontSizeOption);
        switch (alignSelectionOption) {
            case "al":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.LEFT);
                break;
            case "ac":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.CENTER);
                break;
            case "ar":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.RIGHT);
                break;
            case "aj":
                s.setHorizontalAlignment(com.aspose.cells.TextAlignmentType.JUSTIFY);
                break;
        }

        c.setStyle(s);
        RequestContext.getCurrentInstance().update(currentCellClientId);
        purge();
    }

    public int getCurrentColumnId() {
        return currentColumnId;
    }

    public void setCurrentColumnId(int currentColumnId) {
        this.currentColumnId = currentColumnId;
    }

    public int getCurrentRowId() {
        return currentRowId;
    }

    public void setCurrentRowId(int currentRowId) {
        this.currentRowId = currentRowId;
    }

    public String getCurrentCellClientId() {
        return currentCellClientId;
    }

    public void setCurrentCellClientId(String currentCellClientId) {
        this.currentCellClientId = currentCellClientId;
    }

    public boolean isBoldOptionEnabled() {
        return boldOptionEnabled;
    }

    public void setBoldOptionEnabled(boolean boldOptionEnabled) {
        this.boldOptionEnabled = boldOptionEnabled;
    }

    public boolean isItalicOptionEnabled() {
        return italicOptionEnabled;
    }

    public void setItalicOptionEnabled(boolean italicOptionEnabled) {
        this.italicOptionEnabled = italicOptionEnabled;
    }

    public boolean isUnderlineOptionEnabled() {
        return underlineOptionEnabled;
    }

    public void setUnderlineOptionEnabled(boolean underlineOptionEnabled) {
        this.underlineOptionEnabled = underlineOptionEnabled;
    }

    public String getFontSelectionOption() {
        return fontSelectionOption;
    }

    public void setFontSelectionOption(String fontSelectionOption) {
        this.fontSelectionOption = fontSelectionOption;
    }

    public int getFontSizeOption() {
        return fontSizeOption;
    }

    public void setFontSizeOption(int fontSizeOption) {
        this.fontSizeOption = fontSizeOption;
    }

    public String getAlignSelectionOption() {
        return alignSelectionOption;
    }

    public void setAlignSelectionOption(String alignSelectionOption) {
        this.alignSelectionOption = alignSelectionOption;
    }

    public void onFileUpload(FileUploadEvent e) {
        try (InputStream i = e.getFile().getInputstream()) {
            this.asposeWorkbook = new Workbook(i);
            purge();
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
                purge();
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
                purge();
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

//            sendMessage("Cell Changed", String.format("Old: %s New: %s", oldCell.getValue(), newCell.getValue()));
        } catch (com.aspose.cells.CellsException x) {
            x.printStackTrace();
        }
    }

    public void onColumnResize(ColumnResizeEvent e) {
        int columnId = com.aspose.cells.CellsHelper.columnNameToIndex(e.getColumn().getHeaderText());
        if (isLoaded()) {
            try {
                getAsposeWorksheet().getCells().setColumnWidthPixel(columnId, e.getWidth());
            } catch (com.aspose.cells.CellsException cx) {
                sendMessage("Could not resize column", cx.getMessage());
            }
            purge();
        }
    }

    public void addRowAbove() {
        if (getCurrentRowId() < 0) {
            sendMessage("No cell selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().insertRows(getCurrentRowId(), 1, true);
            this.cachedRows = null;
        } catch (com.aspose.cells.CellsException cx) {
            sendMessage("Could not add row", cx.getMessage());
        }
    }

    public void addRowBelow() {
        if (getCurrentRowId() < 0) {
            sendMessage("No cell selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().insertRows(getCurrentRowId() + 1, 1, true);
            this.cachedRows = null;
        } catch (com.aspose.cells.CellsException cx) {
            sendMessage("Could not add row", cx.getMessage());
        }
    }

    public void deleteRow() {
        if (getCurrentRowId() < 0) {
            sendMessage("No cell selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().deleteRows(getCurrentRowId(), 1, true);
            this.cachedRows = null;
        } catch (com.aspose.cells.CellsException cx) {
            sendMessage("Could not delete row", cx.getMessage());
        }
    }

    public void addColumnBefore() {
        if (getCurrentColumnId() < 0) {
            sendMessage("No cell selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().insertColumns(getCurrentColumnId(), 1, true);
            purge();
        } catch (com.aspose.cells.CellsException cx) {
            sendMessage("Could not add column", cx.getMessage());
        }
    }

    public void addColumnAfter() {
        if (getCurrentColumnId() < 0) {
            sendMessage("No cell selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().insertColumns(getCurrentColumnId() + 1, 1, true);
            purge();
        } catch (com.aspose.cells.CellsException cx) {
            sendMessage("Could not add column", cx.getMessage());
        }
    }

    public void deleteColumn() {
        if (getCurrentColumnId() < 0) {
            sendMessage("No cell selected", null);
            return;
        }

        try {
            getAsposeWorksheet().getCells().deleteColumns(getCurrentColumnId(), 1, true);
            purge();
        } catch (com.aspose.cells.CellsException cx) {
            sendMessage("Could not delete column", cx.getMessage());
        }
    }

    public List<String> getFonts() {
        ArrayList<String> list = new ArrayList<>();
        if (isLoaded()) {
            // TODO: get list of fonts used by the workboook
        }
        return list;
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

    public void updatePartialView() {
        String id = FacesContext.getCurrentInstance().getExternalContext().getRequestParameterMap().get("id");
        if (id != null) {
            RequestContext.getCurrentInstance().update(id);
        }
    }

    public void purge() {
        this.cachedColumns = null;
        this.cachedRows = null;
    }
}
