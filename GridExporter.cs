using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SqlServer.Management.UI.VSIntegration.BrokeredServices;

namespace SsmsExcelExport
{
    static class GridExporter
    {
        const int RowChunk = 5000;
        const int ColChunk = 512;
        const int MaxCell = 32767; // Excel cell limit

        public static async Task<int> ExportAsync(IQueryEditorTabDataServiceBrokered tab, string moniker, string path, Action<string> status, CancellationToken ct)
        {
            var panes = await tab.GetAvailablePanesAsync(moniker, ct);
            if (panes == null || !panes.Any(p => p.PaneType == QueryResultsPane.GridResults)) return 0;

            var first = await tab.GetGridResultsSegmentAsync(moniker, 0, 0, ColChunk, 0, 1, MaxCell, ct);
            int gridCount = first == null ? 0 : first.TotalGridCount;
            if (gridCount == 0) return 0;

            using (var xlsx = new XlsxWriter(path))
            {
                for (int g = 0; g < gridCount; g++)
                {
                    status("Exporting result " + (g + 1) + " of " + gridCount + "...");
                    var head = await tab.GetGridResultsSegmentAsync(moniker, g, 0, ColChunk, 0, 1, MaxCell, ct);
                    int totalCols = head.TotalColumnCount;
                    long totalRows = head.TotalRowCount;

                    var names = new List<string>();
                    for (int c = 0; c < totalCols; c += ColChunk)
                    {
                        var seg = await tab.GetGridResultsSegmentAsync(moniker, g, c, Math.Min(ColChunk, totalCols - c), 0, 1, MaxCell, ct);
                        names.AddRange(seg.ColumnNames);
                    }

                    var sheet = xlsx.AddSheet("Result " + (g + 1), names);
                    for (long r = 0; r < totalRows; r += RowChunk)
                    {
                        int take = (int)Math.Min(RowChunk, totalRows - r);
                        List<List<string>> rows = null;
                        for (int c = 0; c < totalCols; c += ColChunk)
                        {
                            var seg = await tab.GetGridResultsSegmentAsync(moniker, g, c, Math.Min(ColChunk, totalCols - c), r, take, MaxCell, ct);
                            if (rows == null) rows = seg.Rows.Select(x => new List<string>(x)).ToList();
                            else for (int i = 0; i < rows.Count && i < seg.Rows.Count; i++) rows[i].AddRange(seg.Rows[i]);
                        }
                        if (rows == null || rows.Count == 0) break;
                        foreach (var row in rows) sheet.AddRow(row);
                        status("Exporting result " + (g + 1) + " of " + gridCount + ": " + Math.Min(r + take, totalRows) + " / " + totalRows + " rows");
                    }
                    sheet.Finish();
                }
                xlsx.Close();
            }
            return gridCount;
        }
    }
}
