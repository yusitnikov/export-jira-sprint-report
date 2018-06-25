// ==UserScript==
// @name        Jira sprint report download
// @description User Script for exporting Jira Sprint Reports as Excel files
// @author      yusitnikov
// @version     1.1
// @updateURL   https://github.com/yusitnikov/export-jira-sprint-report/raw/master/export-jira-sprint-report.user.js
// @match       https://*.atlassian.net/*
// @run-at      document-end
// @grant       none
// ==/UserScript==

(function() {
    var script = document.createElement('script');
    script.src = 'https://cdn.rawgit.com/jmaister/excellentexport/master/dist/excellentexport.js';
    script.type = 'text/javascript';
    document.getElementsByTagName('head')[0].appendChild(script);
})();

setInterval(function() {
    // Check that page is loaded and we are on the Sprint Report page
    if (!window.$ || !window.ExcellentExport || $('#subnav-title').text() !== 'Sprint Report') {
        return;
    }

    // Check that we didn't add the Export button yet
    if ($('#ys-export-sprint-report').size()) {
        return;
    }

    // Add the export button
    var $li = $('<li id="ys-export-sprint-report"><a download="Sprint Report.xlsx" href="#">Export to Excel</a></li>'),
        $a = $li.find('a');
    $li.find('a').click(function() {
        function map(a, f) {
            return Array.prototype.map.call(a, f);
        }

        try {
            return ExcellentExport.convert(
                {
                    anchor: this,
                    filename: 'Sprint Report',
                    format: 'xlsx'
                },
                map(
                    $('#ghx-report-content').find('.ghx-sprint-report-table table'),
                    function(table) {
                        var $table = $(table),
                            $header = $table.prev();
                        return {
                            // Sheet name couldn't be more than 31 letters
                            name: $header.find('h4').text().trim().replace('Issues completed outside of this sprint', 'Completed outside of sprint').substr(0, 31),
                            from: {
                                array: map(
                                    // table -> thead/tbody -> tr
                                    $table.children().children(),
                                    function(tr) {
                                        return map(
                                            $(tr).children(),
                                            function(td) {
                                                return $(td).text().trim();
                                            }
                                        );
                                    }
                                )
                            }
                        };
                    }
                )
            );
        } catch (ex) {
            alert('Unexpected error: ' + ex);
            return false;
        }
    });
    $('#ghx-sprint-actions-content').find('ul').append($li);
}, 500);
