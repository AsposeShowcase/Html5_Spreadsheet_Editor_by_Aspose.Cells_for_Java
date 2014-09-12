package com.aspose.spreadsheeteditor;

import java.io.Serializable;

public class Cell implements Serializable {

    private String value;
    private String html;
    private String formula;
    private int columnId;
    private int rowId;
    private int colspan = 0;
    private int rowspan = 0;

    public void init() {
    }

    public void setValue(String value) {
        this.value = value;
    }

    public void setHtml(String html) {
        this.html = html;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public String getValue() {
        return value;
    }

    public String getHtml() {
        return html;
    }

    public String getFormula() {
        return formula;
    }

    public int getColumnId() {
        return columnId;
    }

    public void setColumnId(int columnId) {
        this.columnId = columnId;
    }

    public int getRowId() {
        return rowId;
    }

    public void setRowId(int rowId) {
        this.rowId = rowId;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public static class Builder {

        protected Cell instance;

        public Builder() {
            this.instance = new Cell();
        }

        public Builder setColumnId(int columnId) {
            this.instance.columnId = columnId;
            return this;
        }

        public Builder setRowId(int rowId) {
            this.instance.rowId = rowId;
            return this;
        }

        public Builder setValue(String value) {
            this.instance.value = value;
            return this;
        }

        public Builder setHtml(String html) {
            this.instance.html = html;
            return this;
        }

        public Builder setFormula(String formula) {
            this.instance.formula = formula;
            return this;
        }

        public Builder setColspan(int colspan) {
            this.instance.colspan = colspan;
            return this;
        }

        public Builder setRowspan(int rowspan) {
            this.instance.rowspan = rowspan;
            return this;
        }

        public Cell build() {
            return this.instance;
        }

    }

}
