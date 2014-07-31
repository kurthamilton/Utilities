RegExp.escape = function (str) {
    var specials = new RegExp("[.*+?|()\\[\\]{}\\\\]", "g"); // .*+?|()[]{}\
    return str.replace(specials, "\\$&");
}

var scriptSearch;

$(function () {
    bindCodeEditors();
    scriptSearch = scriptSearch || Window.ScriptSearch();    
});

var codeEditor;

function bindCodeEditors() {

    var provider = '';

    if (provider === 'ace') {
        if (!codeEditor) {
            codeEditor = ace.edit('editor');
            codeEditor.setTheme('ace/theme/monokai');
            codeEditor.setReadOnly(true);
            codeEditor.getSession().setMode('ace/mode/sql');
        }

        $(document).bind('codeUpdate', function (ev, args) {
            codeEditor.setValue(args.value)
        });
    }
    else if (provider === 'SyntaxHighlighter') {
        /* shBrushSql.js customisations:
        added funcs: error_message() max min @@trancount
        added keywords: catch go if nvarchar print raiserror try while
        remove keywords: max min
        added operators: exists
        added regex:
        { regex: /\/\*(.|\n|\r)*?\*\/$/gm, css: 'comments' }, // multiline comments
        */
        SyntaxHighlighter.all();
    }
}

(function () {
    Window.ScriptSearch = function () {
        
        bindHoverDialogs();
        bindSearchTextBoxes();

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
            var editorId = link.attr('data-hoverdialog-editorid');
            if (previewId && previewId != '') {
                var preview = $('#' + previewId);
                var editor = preview;
                if (editorId && editorId != '') {
                    editor = $('#' + editorId);
                }

                if (show) {
                    var html = dialog.html();
                    editor.html(html);
                    preview.show();
                }
                else {
                    var selectedLink = $('a.dialog.selected');

                    if (selectedLink.length > 0) {
                        toggleHoverDialog(selectedLink, true);
                    }
                    else {
                        preview.hide();
                        editor.html('');
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
                var searchRegEx = new RegExp(RegExp.escape(searchTerm), 'ig');
                var replaceTerm = '<span class="searchresult">$&</span>';

                $('.searchable').each(function () {
                    var searchableElement = $(this);
                    var filterItem = searchableElement.closest('.filter-item');

                    if (searchTerm !== '') {
                        var html = searchableElement.html();
                        var match = html.match(searchRegEx);
                        if (match) {
                            searchableElement.html(html.replace(searchRegEx, replaceTerm));
                        }

                        if (filterItem.length > 0) {
                            if (match) {
                                filterItem.show();
                            } else {
                                filterItem.hide();
                            }
                        }
                    }
                    else {
                        if (filterItem.length > 0) {
                            filterItem.show();
                        }
                    }
                });
            });
        }

        function clearSearchResults() {
            $('span.searchresult').each(function () {
                var searchResult = $(this);
                searchResult.replaceWith(searchResult.text());
            });
        }        
    };
})();

function selectContent(element) {
    if (element) {
        if (document.selection) {
            var range = document.body.createTextRange();
            range.moveToElementText(element);
            range.select();
        } else if (window.getSelection) {
            var range = document.createRange();
            range.selectNode(element);
            window.getSelection().addRange(range);
        }
    }
}