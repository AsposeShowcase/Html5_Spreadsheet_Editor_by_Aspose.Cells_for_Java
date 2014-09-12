package com.aspose.spreadsheeteditor;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author saqib
 */
public class Row implements Serializable {

    private int id;
    private HashMap<String, Cell> cellsMap = new HashMap<>();
    private ArrayList<Cell> cellsList = new ArrayList<>();

    public int getId() {
        return id;
    }

    public Map<String, Cell> getCellsMap() {
        return this.cellsMap;
    }

    public List<Cell> getCellsList() {
        return this.cellsList;
    }

    public Cell[] getCells_() {
        Cell[] cellArray = new Cell[this.cellsMap.keySet().size()];
        for (String k : this.cellsMap.keySet()) {
            cellArray[com.aspose.cells.CellsHelper.columnNameToIndex(k)] = this.cellsMap.get(k);
        }
        return cellArray;
    }

    public static class Builder {

        protected Row instance;

        public Builder() {
            this.instance = new Row();
        }

        public int getId() {
            return this.instance.id;
        }

        public Builder setId(int id) {
            this.instance.id = id;
            return this;
        }

        public Builder setCell(String column, Cell cell) {
            this.instance.cellsMap.put(column, cell);
            
            int columnId = com.aspose.cells.CellsHelper.columnNameToIndex(column);
            this.instance.cellsList.ensureCapacity(columnId + 1);
            this.instance.cellsList.add(columnId, cell);
            return this;
        }

        public Row build() {
            return instance;
        }
    }
}
