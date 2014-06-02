$(function () {
    bindCodeEditors();
    bindHoverDialogs();
    bindSearchTextBoxes();
});

function bindCodeEditors() {
    var editor = ace.edit('scriptpreview');
    editor.setTheme('ace/theme/monokai');
    editor.setReadOnly(true);
    editor.setBehavioursEnabled(true);
    editor.getSession().setMode('ace/mode/sql');
}

function bindHoverDialogs() {
    $('a.dialog').on({
        click: function () {
            $('a.dialog.selected').removeClass('selected');
            $(this).addClass('selected');
        },
        mouseenter: function () {
            toggleHoverDialog($(this), true);
        },
        mouseleave: function () {
            toggleHoverDialog($(this), false);
        }
    });
}

function toggleHoverDialog(link, show) {
    var dialog = link.next('div.dialog');

    var previewId = link.attr('data-hoverdialog-previewid');
    if (previewId && previewId != '') {
        var preview = $('#' + previewId);
        if (show) {
            preview.html(dialog.html());
            bindCodeEditors();
        }
        else {
            var selectedLink = $('a.dialog.selected');
            
            if (selectedLink.length > 0) {
                toggleHoverDialog(selectedLink, true);
            }
            else {
                preview.html('');
            }
        }
    }
    else {
        if (show) {
            dialog.show();
        } else {
            dialog.hide();
        }
    }
}

function bindSearchTextBoxes() {
    $('input[type="text"].search').on('keyup', function () {
        clearSearchResults();

        var searchTerm = $(this).val();
        var searchRegEx = new RegExp(searchTerm, 'ig');
        var replaceTerm = '<span class="searchresult">' + searchTerm + '</span>';

        $('.searchable').each(function () {
            var searchableElement = $(this);
            searchableElement.html(searchableElement.html().replace(searchTerm, replaceTerm));
        });
    });
}

function clearSearchResults() {
    $('span.searchresult').each(function () {
        var searchResult = $(this);
        searchResult.replaceWith(searchResult.text());
    });
}