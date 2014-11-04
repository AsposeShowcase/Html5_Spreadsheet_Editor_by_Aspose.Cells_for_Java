package com.aspose.spreadsheeteditor;

import java.io.Serializable;

public class Cell implements Serializable {

    private String value;
    private String formula;
    private String style;
    private int columnId;
    private int rowId;
    private int colspan = 0;
    private int rowspan = 0;

    public void init() {
    }

    public Cell setValue(String value) {
        this.value = value;
        return this;
    }

    public Cell setFormula(String formula) {
        this.formula = formula;
        return this;
    }

    public String getValue() {
        return value;
    }

    public String getFormula() {
        return formula;
    }

    public String getStyle() {
        return style;
    }

    public Cell setStyle(String style) {
        this.style = style;
        return this;
    }

    public int getColumnId() {
        return columnId;
    }

    public Cell setColumnId(int columnId) {
        this.columnId = columnId;
        return this;
    }

    public int getRowId() {
        return rowId;
    }

    public Cell setRowId(int rowId) {
        this.rowId = rowId;
        return this;
    }

    public int getColspan() {
        return colspan;
    }

    public Cell setColspan(int colspan) {
        this.colspan = colspan;
        return this;
    }

    public int getRowspan() {
        return rowspan;
    }

    public Cell setRowspan(int rowspan) {
        this.rowspan = rowspan;
        return this;
    }
}
