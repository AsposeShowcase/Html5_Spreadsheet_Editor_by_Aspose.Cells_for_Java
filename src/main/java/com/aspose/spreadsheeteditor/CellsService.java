package com.aspose.spreadsheeteditor;

import java.util.HashMap;
import java.util.List;
import javax.enterprise.context.ApplicationScoped;
import javax.inject.Inject;
import javax.inject.Named;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "cells")
@ApplicationScoped
public class CellsService {

    private HashMap<String, List<Cell>> cells = new HashMap<>();
    private HashMap<String, List<Column>> columns = new HashMap<>();
    private HashMap<String, List<Row>> rows = new HashMap<>();
    private HashMap<String, List<Integer>> columnWidth = new HashMap<>();
    private HashMap<String, List<Integer>> rowHeight = new HashMap<>();

    @Inject
    private MessageService msg;

    public void put(String key, List<Cell> cells) {
        this.cells.put(key, cells);
    }

    public List<Cell> get(String key) {
        return this.cells.get(key);
    }

    public List<Column> getColumns(String key) {
        return columns.get(key);
    }

    public void putColumns(String key, List<Column> c) {
        columns.put(key, c);
    }

    public List<Row> getRows(String key) {
        return rows.get(key);
    }

    public void putRows(String key, List<Row> r) {
        rows.put(key, r);
    }

    public Cell getCell(String key, int column, int row) {
        return rows.get(key).get(row).getCellsMap().get(com.aspose.cells.CellsHelper.columnIndexToName(column));
    }

    public void putCell(String key, int column, int row, Cell c) {
        rows.get(key).get(row).getCellsMap().put(com.aspose.cells.CellsHelper.columnIndexToName(column), c);
    }

    public List<Integer> getColumnWidth(String key) {
        return columnWidth.get(key);
    }

    public void putColumnWidth(String key, List<Integer> c) {
        columnWidth.put(key, c);
    }

    public List<Integer> getRowHeight(String key) {
        return rowHeight.get(key);
    }

    public void putRowHeight(String key, List<Integer> r) {
        rowHeight.put(key, r);
    }

    public Cell fromBlank(int columnId, int rowId) {
        return new Cell().setColumnId(columnId).setRowId(rowId);
    }

    public Cell fromAsposeCell(com.aspose.cells.Cell a) {
        // a = Aspose.Cells' definition of a cell
        // c = Spreassheet's definition of a cell

        Cell c = new Cell()
                .setColumnId(a.getColumn())
                .setRowId(a.getRow());

        try {
            a.calculate(true, null);
        } catch (com.aspose.cells.CellsException cx) {
            msg.sendMessage("Cell recalculation failure", cx.getMessage());
        }

        try {
            c.setFormula(a.getFormula())
                    .setValue(a.getStringValueWithoutFormat());
        } catch (Exception x) {
            msg.sendMessage("Cell value error", x.getMessage());
        }

        StringBuilder style = new StringBuilder();

        try {
//            com.aspose.cells.Color cellBgColor = a.getStyle().getBackgroundColor();
//            style.append("background-color:")
//                    .append(asposeColorToCssColor(cellBgColor, false))
//                    .append(";");

            com.aspose.cells.Color cellFgColor = a.getStyle().getForegroundColor();
            style.append("background-color:")
                    .append(asposeColorToCssColor(cellFgColor, false))
                    .append(";");

            com.aspose.cells.Font font = a.getStyle().getFont();
            style.append("font-family:'").append(font.getName()).append("';");

            if (a.getStyle().getFont().isItalic()) {
                c.addClass("i").setItalic(true);
            }

            if (a.getStyle().getFont().isBold()) {
                c.addClass("b").setBold(true);
            }

            switch (a.getStyle().getFont().getUnderline()) {
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
                    c.addClass("u").setUnderline(true);
                    break;
            }

            switch (a.getStyle().getFont().getStrikeType()) {
                case com.aspose.cells.TextStrikeType.SINGLE:
                case com.aspose.cells.TextStrikeType.DOUBLE:
                    c.addClass("sts");
                    break;
            }

            switch (a.getStyle().getFont().getCapsType()) {
                case com.aspose.cells.TextCapsType.ALL:
                    c.addClass("uc");
                    break;
                case com.aspose.cells.TextCapsType.SMALL:
                    c.addClass("sc");
                    break;
            }

//            if (a.getStyle().getFont().getSize()) {
            style
                    .append("font-size:")
                    .append(a.getStyle().getFont().getSize())
                    .append("pt;");
//            }

            switch (a.getStyle().getHorizontalAlignment()) {
                case com.aspose.cells.TextAlignmentType.GENERAL:
                case com.aspose.cells.TextAlignmentType.LEFT:
                    c.addClass("al");
                    break;
                case com.aspose.cells.TextAlignmentType.RIGHT:
                    c.addClass("ar");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER_ACROSS:
                case com.aspose.cells.TextAlignmentType.CENTER:
                    c.addClass("ac");
                    break;
                case com.aspose.cells.TextAlignmentType.JUSTIFY:
                    c.addClass("aj");
                    break;
            }

            switch (a.getStyle().getVerticalAlignment()) {
                case com.aspose.cells.TextAlignmentType.TOP:
                    c.addClass("at");
                    break;
                case com.aspose.cells.TextAlignmentType.CENTER:
                    c.addClass("am");
                    break;
                case com.aspose.cells.TextAlignmentType.BOTTOM:
//                    style.append("vertical-align: bottom;");
                    c.addClass("ab");
                    break;
            }

//            int cellRotationAngle = a.getStyle().getRotationAngle();
//            cellRotationAngle = 0; // TODO
//            style.append("writing-mode: lr-bt ;")
//                    .append("transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-webkit-transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-moz-transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-ms-transform: rotate(-").append(cellRotationAngle).append("deg);")
//                    .append("-o-transform: rotate(-").append(cellRotationAngle).append("deg);");
            com.aspose.cells.Color cellTextColor = a.getStyle().getFont().getColor();
            style.append("color:")
                    .append(asposeColorToCssColor(cellTextColor, true))
                    .append(";");

        } catch (Exception x) {
            msg.sendMessage("Cell style error", x.getMessage());
        }

        c.setStyle(style.toString());

        if (a.getMergedRange() != null) {
            if (a.getColumn() == a.getMergedRange().getFirstColumn()) {
//                c.setColspan(a.getMergedRange().getColumnCount());
            }
            if (a.getRow() == a.getMergedRange().getFirstRow()) {
//                c.setRowspan(a.getMergedRange().getRowCount());
            }
        }

        return c;
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
}
