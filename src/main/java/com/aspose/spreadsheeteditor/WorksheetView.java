package com.aspose.spreadsheeteditor;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import javax.annotation.PostConstruct;
import javax.faces.context.FacesContext;
import javax.faces.view.ViewScoped;
import javax.inject.Inject;
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

    private String loadedWorkbook;
    private Format outputFormat = Format.XLSX;
    private String sourceUrl;
    private int currentColumnId;
    private int currentRowId;
    private String currentCellClientId;

    @Inject
    private CellsService cells;

    @Inject
    private MessageService msg;

    @Inject
    private LoaderService loader;

    @Inject
    FormattingService formatting;

    @PostConstruct
    private void init() {
        String requestedSourceUrl = FacesContext.getCurrentInstance().getExternalContext().getRequestParameterMap().get("url");
        if (requestedSourceUrl != null) {
            try {
                this.sourceUrl = new URL(requestedSourceUrl).toString();
                this.loadFromUrl();
            } catch (MalformedURLException x) {
                msg.sendMessageDialog("The specified URL is invalid", requestedSourceUrl);
            }
        }
    }

    public String getLoadedWorkbook() {
        return this.loadedWorkbook;
    }

    public void setLoadedWorkbook(String id) {
        this.loadedWorkbook = id;
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
        return this.loadedWorkbook != null;
    }

    public Format getOutputFormat() {
        return this.outputFormat;
    }

    public void setOutputFormat(Format outputFormat) {
        this.outputFormat = outputFormat;
    }

    public void loadBlank() {
        this.loadedWorkbook = loader.fromBlank();
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

        loadedWorkbook = loader.fromUrl(this.sourceUrl);
    }

    public com.aspose.cells.Workbook getAsposeWorkbook() {
        return loader.get(loadedWorkbook);
    }

    public int getDefaultColumnWidth() {
        try {
            return getAsposeWorksheet().getCells().getStandardWidthPixels();
        } catch (com.aspose.cells.CellsException | NullPointerException x) {
            return 64;
        }
    }

    public List<Integer> getColumnWidth() {
        return cells.getColumnWidth(loadedWorkbook);
    }

    public List<Integer> getRowHeight() {
        return cells.getRowHeight(loadedWorkbook);
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
        return cells.getColumns(loadedWorkbook);
    }

    public List<Row> getRows() {
        return cells.getRows(loadedWorkbook);
    }

    public void applyCellFormatting() {
        com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(currentRowId, currentColumnId);
        com.aspose.cells.Style s = c.getStyle();

        s.getFont().setBold(formatting.isBoldOptionEnabled());
        s.getFont().setItalic(formatting.isItalicOptionEnabled());
        s.getFont().setUnderline(formatting.isUnderlineOptionEnabled() ? com.aspose.cells.FontUnderlineType.SINGLE : com.aspose.cells.FontUnderlineType.NONE);
        s.getFont().setName(formatting.getFontSelectionOption());
        s.getFont().setSize(formatting.getFontSizeOption());
        switch (formatting.getAlignSelectionOption()) {
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

    public void onFileUpload(FileUploadEvent e) {
        try {
            loadedWorkbook = loader.fromInputStream(e.getFile().getInputstream(), e.getFile().getFileName());
            purge();
        } catch (IOException x) {
            throw new RuntimeException(x);
        }
    }

    public void onAddNewSheet() {
        if (isLoaded()) {
            try {
                int i = getAsposeWorkbook().getWorksheets().add();
                getAsposeWorkbook().getWorksheets().setActiveSheetIndex(i);
                purge();
            } catch (com.aspose.cells.CellsException cx) {
                msg.sendMessage("New Worksheet", cx.getMessage());
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
                msg.sendMessage("Could not remove sheet", cx.getMessage());
            }
        }
    }

    public void onCellEdit(CellEditEvent e) {
        Cell oldCell = (Cell) e.getOldValue();
        Cell newCell = (Cell) e.getNewValue();
        int columnId = newCell.getRowId();
        int rowId = newCell.getColumnId();

        try {
            com.aspose.cells.Cell c = getAsposeWorksheet().getCells().get(rowId, columnId);
            if (newCell.getFormula() != null) {
                c.setFormula(newCell.getFormula(), null);
            } else {
                c.putValue(newCell.getValue(), true);
            }
            cells.putCell(loadedWorkbook, columnId, rowId, newCell);
        } catch (com.aspose.cells.CellsException x) {
            x.printStackTrace();
        }
    }

    public void onColumnResize(ColumnResizeEvent e) {
        if (!isLoaded()) {
            return;
        }

        int columnId = com.aspose.cells.CellsHelper.columnNameToIndex(e.getColumn().getHeaderText());
        try {
            getAsposeWorksheet().getCells().setColumnWidthPixel(columnId, e.getWidth());
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Could not resize column", cx.getMessage());
            return;
        }

        reloadColumnWidth(columnId);
    }

    public void addRowAbove() {
        try {
            getAsposeWorksheet().getCells().insertRows(currentRowId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Could not add row", cx.getMessage());
            return;
        }

        purge();
        reloadRowHeight(currentRowId);
    }

    public void addRowBelow() {
        if (getCurrentRowId() < 0) {
            msg.sendMessage("No cell selected", null);
            return;
        }

        int newRowId = currentRowId + 1;

        try {
            getAsposeWorksheet().getCells().insertRows(newRowId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Could not add row", cx.getMessage());
            return;
        }

//        Row r = new Row.Builder().setId(newRowId).build();
//        cells.getRows(loadedWorkbook).add(newRowId, r);
//
//        for (int i = 0; i < cells.getColumns(loadedWorkbook).size(); i++) {
//            r.putCell(i, cells.fromBlank(i, newRowId));
//        }
        purge();
        reloadRowHeight(newRowId);
    }

    public void deleteRow() {
        try {
            getAsposeWorksheet().getCells().deleteRows(currentRowId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Could not delete row", cx.getMessage());
            return;
        }

        cells.getRows(loadedWorkbook).remove(currentRowId);
        getRowHeight().remove(currentRowId);
        purge();
    }

    public void addColumnBefore() {
        try {
            getAsposeWorksheet().getCells().insertColumns(getCurrentColumnId(), 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Could not add column", cx.getMessage());
            return;
        }

        reloadColumnWidth(currentColumnId);
        purge();
    }

    public void addColumnAfter() {
        int newColumnId = currentColumnId + 1;
        try {
            getAsposeWorksheet().getCells().insertColumns(newColumnId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Could not add column", cx.getMessage());
            return;
        }

        reloadColumnWidth(newColumnId);
        purge();
    }

    public void deleteColumn() {
        try {
            getAsposeWorksheet().getCells().deleteColumns(currentColumnId, 1, true);
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Could not delete column", cx.getMessage());
            return;
        }

        cells.getColumns(loadedWorkbook).remove(currentColumnId);
        getRowHeight().remove(currentColumnId);
        purge();
    }

    public void addCellShiftRight() {
        if (!isLoaded()) {
            return;
        }

        com.aspose.cells.CellArea a = new com.aspose.cells.CellArea();
        a.StartColumn = a.EndColumn = currentColumnId;
        a.StartRow = a.EndRow = currentRowId;
        getAsposeWorksheet().getCells().insertRange(a, com.aspose.cells.ShiftType.RIGHT);
        purge();
    }

    public void addCellShiftDown() {
        if (!isLoaded()) {
            return;
        }

        com.aspose.cells.CellArea a = new com.aspose.cells.CellArea();
        a.StartColumn = a.EndColumn = currentColumnId;
        a.StartRow = a.EndRow = currentRowId;
        getAsposeWorksheet().getCells().insertRange(a, com.aspose.cells.ShiftType.DOWN);
        purge();
    }

    public void removeCellShiftUp() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().deleteRange(currentRowId, currentColumnId, currentRowId, currentColumnId, com.aspose.cells.ShiftType.UP);
        purge();
    }

    public void removeCellShiftLeft() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().deleteRange(currentRowId, currentColumnId, currentRowId, currentColumnId, com.aspose.cells.ShiftType.LEFT);
        purge();
    }

    public void clearCurrentCellFormatting() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().clearFormats(currentRowId, currentColumnId, currentRowId, currentColumnId);
        reloadCell(currentColumnId, currentRowId);
        RequestContext.getCurrentInstance().update(currentCellClientId);
    }

    public void clearCurrentCellContents() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().clearContents(currentRowId, currentColumnId, currentRowId, currentColumnId);
        reloadCell(currentColumnId, currentRowId);
        RequestContext.getCurrentInstance().update(currentCellClientId);
    }

    public void clearCurrentCell() {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().clearRange(currentRowId, currentColumnId, currentRowId, currentColumnId);
        reloadCell(currentColumnId, currentRowId);
        RequestContext.getCurrentInstance().update(currentCellClientId);
    }

    public int getCurrentColumnWidth() {
        return getColumnWidth().get(currentColumnId);
    }

    public void setCurrentColumnWidth(int width) {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().setColumnWidthPixel(currentColumnId, width);
        reloadColumnWidth(currentColumnId);
        RequestContext.getCurrentInstance().update("sheet");

    }

    public int getCurrentRowHeight() {
        return getRowHeight().get(currentRowId);
    }

    public void setCurrentRowHeight(int height) {
        if (!isLoaded()) {
            return;
        }

        getAsposeWorksheet().getCells().setRowHeightPixel(currentRowId, height);
        reloadRowHeight(currentRowId);
        RequestContext.getCurrentInstance().update("sheet");
    }

    public List<String> getFonts() {
        ArrayList<String> list = new ArrayList<>();
        if (isLoaded()) {
            // TODO: get list of fonts used by the workboook
        }
        return list;
    }

    public void updatePartialView() {
        String id = FacesContext.getCurrentInstance().getExternalContext().getRequestParameterMap().get("id");
        if (id != null) {
            RequestContext.getCurrentInstance().update(id);
        }
    }

    private void reloadColumnWidth(int columnId) {
        int width = getAsposeWorksheet().getCells().getColumnWidthPixel(columnId);
        getColumnWidth().remove(columnId);
        getColumnWidth().add(columnId, width);
    }

    private void reloadRowHeight(int rowId) {
        int height = getAsposeWorksheet().getCells().getRowHeightPixel(rowId);
        getRowHeight().remove(rowId);
        getRowHeight().add(rowId, height);
    }

    private void reloadCell(int columnId, int rowId) {
        com.aspose.cells.Cell a = getAsposeWorksheet().getCells().get(rowId, columnId);
        Cell c = cells.fromAsposeCell(a);
        cells.getRows(loadedWorkbook).get(rowId).putCell(columnId, c);
    }

    public void purge() {
        loader.buildCellsCache(loadedWorkbook);
    }
}
