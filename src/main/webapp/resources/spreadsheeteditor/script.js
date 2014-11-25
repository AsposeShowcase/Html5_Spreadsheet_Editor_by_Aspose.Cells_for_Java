function showAboutDialog() {
    PF('aboutDialog').show();
}

function showFileUploadDialog() {
    PF('fileUploadDialog').chooseButton.find('input[type=file]').click();
}

(function($) {
    $(document).on('focus', '.ui-datatable .ui-cell-editor-input input[type=text]', function(e) {
        var columnId = $(this).data('columnid');
        var rowId = $(this).data('rowid');

        PF('currentColumnIdValue').jq.val(columnId);
        PF('currentRowIdValue').jq.val(rowId);
    });
})(jQuery);

function fillCurrentColumnRowValues(targetContainerSelector) {
    var columnId = PF('currentColumnIdValue').jq.val();
    var rowId = PF('currentRowIdValue').jq.val();

    $(targetContainerSelector).find('.current-columnId-holder').val(columnId);
    $(targetContainerSelector).find('.current-rowId-holder').val(rowId);
}

