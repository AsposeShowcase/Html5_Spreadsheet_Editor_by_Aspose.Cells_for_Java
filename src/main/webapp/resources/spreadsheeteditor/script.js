function showAboutDialog() {
    PF('aboutDialog').show();
    return false;
}

function showFileUploadDialog() {
    PF('fileUploadDialog').chooseButton.find('input[type=file]').click();
    return false;
}

function showColumnWidthDialog() {
    PF('columnWidthDialog').show();
    return false;
}

function hideColumnWidthDialog() {
    PF('columnWidthDialog').hide();
    return false;
}

function showRowHeightDialog() {
    PF('rowHeightDialog').show();
    return false;
}

function hideRowHeightDialog() {
    PF('rowHeightDialog').hide();
    return false;
}

PrimeFaces.widget.DataTable = PrimeFaces.widget.DataTable.extend({
    bindEditEvents: function() {
        var $this = this;
        // From primefaces' datatable.js

        var cellSelector = '> tr > td.ui-editable-column';
        this.tbody.off('click.datatable-cell', cellSelector)
                .on('click.datatable-cell', cellSelector, null, function(e) {
                    var columnId = $(this).find('.ui-cell-editor-input input').attr('data-columnid');
                    var rowId = $(this).find('.ui-cell-editor-input input').attr('data-rowid');
                    var clientId = $(this).find('.ui-cell-editor').attr('id');

                    PF('currentColumnIdValue').jq.val(columnId);
                    PF('currentRowIdValue').jq.val(rowId);
                    PF('currentCellClientIdValue').jq.val(clientId);
                    if ($(this).find('.ui-cell-editor-output div').hasClass('b')) {
                        PF('boldOptionButton').check();
                    } else {
                        PF('boldOptionButton').uncheck();
                    }
                    if ($(this).find('.ui-cell-editor-output div').hasClass('i')) {
                        PF('italicOptionButton').check();
                    } else {
                        PF('italicOptionButton').uncheck();
                    }
                    if ($(this).find('.ui-cell-editor-output div').hasClass('u')) {
                        PF('underlineOptionButton').check();
                    } else {
                        PF('underlineOptionButton').uncheck();
                    }
                    PF('fontOptionSelector').selectValue($(this).find('.ui-cell-editor-output div').css('font-family').slice(1, -1));
                    PF('fontSizeOptionSelector').selectValue($(this).find('.ui-cell-editor-output div')[0].style.fontSize.replace('pt', ''));
                    ['al', 'ac', 'ar', 'aj'].forEach(function(v) {
                        if ($(this).find('.ui-cell-editor-output div').hasClass(v)) {
                            // TODO: save the value to PF('alignOptionSelector');
                        }
                    });
                    PF('currentColumnWidthValue').jq.val($(this).width());
                    PF('currentRowHeightValue').jq.val($(this).height());

                    $($this.selectedCell).removeClass('sheet-selected-cell');
                    $(this).addClass('sheet-selected-cell');
                    $this.selectedCell = this;
                });

        this.tbody.off('dblclick.datatable-cell', cellSelector)
                .on('dblclick.datatable-cell', cellSelector, null, function(e) {
                    $this.incellClick = true;
                    var cell = $(this);
                    if (!cell.hasClass('ui-cell-editing')) {
                        $this.showCellEditor($(this));
                    }
                });

        $(document).off('click.datatable-cell-blur' + this.id)
                .on('click.datatable-cell-blur' + this.id, function(e) {
                    if (!$this.incellClick && $this.currentCell && !$this.contextMenuClick) {
                        $this.saveCell($this.currentCell);
                    }
                    $this.incellClick = false;
                    $this.contextMenuClick = false;
                });

    }
});

function updateCurrentCells() {
    var columnId = PF('currentColumnIdValue').jq.val();
    var rowId = PF('currentRowIdValue').jq.val();
    var sel = '.ui-datatable .ui-cell-editor-input input[data-columnid=' + columnId + '][data-rowid=' + rowId + ']';

    updatePartialView([
        {name: 'id', value: $(sel).closest('.ui-cell-editor').attr('id')}
    ]);
}

setInterval(function() {
    if (typeof $ === 'undefined' && !$('#sheet .ui-datatable').length) {
        return;
    }

    $('#sheet .ui-datatable').height($(window).height() - $('#sheet .ui-datatable').position().top - 1);
}, 786);

$(document).ready(function() {
    [
        '.ui-button.cell-formatting',
        '.ui-selectonemenu-panel.cell-formatting .ui-selectonemenu-list-item',
        '.ui-selectonebutton.cell-formatting .ui-button',
    ].forEach(function(sel) {
        $(document).on('click', sel, function(e) {
            PF('applyFormattingButton').getJQ().click();
        });
    });
});

