package com.aspose.spreadsheeteditor;

import javax.el.ValueExpression;
import javax.faces.component.UIComponent;
import javax.faces.context.FacesContext;
import javax.faces.convert.Converter;
import javax.faces.convert.FacesConverter;
import javax.inject.Inject;

@FacesConverter("cellEditorConverter")
public class CellConverter implements Converter {

    @Inject
    private WorksheetView worksheet;

    public WorksheetView getWorksheet() {
        return worksheet;
    }

    public void setWorksheet(WorksheetView worksheet) {
        this.worksheet = worksheet;
    }

    @Override
    public Object getAsObject(FacesContext context, UIComponent component, String value) {
        int columnId = (Integer) ((ValueExpression) component.getPassThroughAttributes().get("data-columnId")).getValue(context.getELContext());
        int rowId = (Integer) ((ValueExpression) component.getPassThroughAttributes().get("data-rowId")).getValue(context.getELContext());

        Cell cell = new Cell().setColumnId(columnId).setRowId(rowId);
        if (value.trim().startsWith("=")) {
            cell.setFormula(value.trim());
        } else {
            cell.setValue(value);
        }

//        worksheet.setCell(columnId, rowId, builder.build());
//        RequestContext.getCurrentInstance().update(":form:sheet");
        return cell;
    }

    @Override
    public String getAsString(FacesContext context, UIComponent component, Object value) {
        Cell cell = (Cell) value;

        if (cell.getFormula() != null) {
            return cell.getFormula();
        } else {
            return cell.getValue();
        }
    }
}
