// ==UserScript==
// @name        Jira sprint report download
// @description User Script for exporting Jira Sprint Reports as Excel files
// @author      yusitnikov
// @version     1.3
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
    // Check that page is loaded
    if (!window.$ || !window.ExcellentExport) {
        return;
    }

    // Check that we didn't add the Export button yet
    if ($('#ys-export-report').size()) {
        return;
    }

    var baseUrl = location.protocol + '//' + location.host,
        reportName = $('#subnav-title').text();
    switch (reportName) {
        case 'Sprint Report':
            addExportButton(map(
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
                                function(tr, trIndex) {
                                    var $td = $(tr).children();
                                    var cells = map(
                                        $td,
                                        function(td) {
                                            return $(td).text().trim();
                                        }
                                    );
                                    var addedAfterSprint = cells[0].endsWith(' *');
                                    cells.splice(1, 0, trIndex === 0 ? 'Link' : baseUrl + '/' + $td.first().find('a').attr('href'), trIndex === 0 ? 'Added After Sprint Start' : addedAfterSprint);
                                    if (addedAfterSprint) {
                                        cells[0] = cells[0].substr(0, cells[0].length - 2);
                                    }
                                    return cells;
                                }
                            )
                        }
                    };
                }
            ));
            break;
        case 'Burndown Chart':
            var $tr = $('#ghx-chart-data').find('table tbody tr'),
                prevRow = [],
                rows = map(
                    $tr,
                    function(tr, trIndex) {
                        var $td = $(tr).children();
                        var cells = map(
                            $td,
                            function(td) {
                                return $(td).text().trim();
                            }
                        );
                        cells[0] = cells[0] || prevRow[0];
                        cells.splice(2, 0, baseUrl + $td.eq(1).find('a').attr('href'));
                        if (trIndex === 0) {
                            cells[1] = 'Many';
                            cells[2] = '';
                            cells[5] = '0';
                            cells[8] = cells[10];
                        }
                        prevRow = cells;
                        return cells;
                    }
                ),
                $startRow = $tr.first().find('td'),
                $startIssues = $startRow.eq(1).find('a'),
                $startTimes = $startRow.eq(7).find('div'),
                startRows = map(
                    $startIssues,
                    function(issue, issueIndex) {
                        var $issue = $(issue);
                        return [
                            $issue.text(),
                            baseUrl + $issue.attr('href'),
                            $startTimes.eq(issueIndex).text()
                        ];
                    }
                );
            rows.unshift(
                [ '',     '',      '',     '',           '',             'Time Spent', '', '',  'Remaining Time Estimate'   ],
                [ 'Date', 'Issue', 'Link', 'Event Type', 'Event Detail', 'Inc.', 'Dec.', 'Sum', 'Inc.', 'Dec.', 'Remaining' ]
            );
            startRows.unshift([ 'Issue', 'Link', 'Estimation' ]);
            addExportButton([
                {
                    name: reportName,
                    from: {
                        array: rows
                    }
                },
                {
                    name: 'Sprint Start',
                    from: {
                        array: startRows
                    }
                }
            ]);
            break;
    }

    function map(a, f) {
        return Array.prototype.map.call(a, f);
    }

    function addExportButton(data) {
        var $button = $('<a id="ys-export-report" class="aui-button" download="' + reportName + '.xlsx" href="#">Export to Excel</a>');
        $button.click(function() {
            try {
                return ExcellentExport.convert(
                    {
                        anchor: this,
                        filename: reportName,
                        format: 'xlsx'
                    },
                    data
                );
            } catch (ex) {
                alert('Unexpected error: ' + ex);
                return false;
            }
        });
        $('#ghx-chart-sprint-actions').prepend($button);
    }
}, 500);
