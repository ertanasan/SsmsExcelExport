using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.ServiceHub.Framework;
using Microsoft.SqlServer.Management.UI.VSIntegration.BrokeredServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell.ServiceBroker;
using Task = System.Threading.Tasks.Task;

namespace SsmsExcelExport
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("Export Results to Excel", "Exports all result grids to one xlsx", "1.0")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuid)]
    public sealed class ExportPackage : AsyncPackage
    {
        public const string PackageGuid = "3b9f2c64-8a1d-4f57-b0e2-6d5a97c1e831";
        static readonly Guid CmdSet = new Guid("e4d7a5b0-2c93-4b18-8f6a-1a0c5d3e7f42");

        protected override async Task InitializeAsync(CancellationToken ct, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(ct);
            var mcs = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs == null) return;
            foreach (var id in new[] { 0x0100, 0x0101 })
                mcs.AddCommand(new MenuCommand((s, e) => JoinableTaskFactory.RunAsync(ExportAsync), new CommandID(CmdSet, id)));
        }

        async Task ExportAsync()
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();
            try
            {
                var container = await GetServiceAsync(typeof(SVsBrokeredServiceContainer)) as IBrokeredServiceContainer;
                if (container == null) { Msg("Brokered service container is not available."); return; }
                var broker = container.GetFullAccessServiceBroker();

                string moniker;
                var editorSvc = await broker.GetProxyAsync<ISqlEditorServiceBrokered>(SqlEditorBrokeredServiceDescriptors.SqlEditorService, DisposalToken);
                try
                {
                    if (editorSvc == null) { Msg("SSMS SQL editor service is not available."); return; }
                    var list = (await editorSvc.GetOpenEditorsAsync(DisposalToken)).ToList();
                    var active = list.FirstOrDefault(x => x.IsActive) ?? list.FirstOrDefault();
                    if (active == null) { Msg("No open query window found."); return; }
                    moniker = active.EditorMoniker;
                }
                finally { (editorSvc as IDisposable)?.Dispose(); }

                var dlg = new SaveFileDialog
                {
                    Filter = "Excel workbook (*.xlsx)|*.xlsx",
                    FileName = "Results_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx"
                };
                if (dlg.ShowDialog() != DialogResult.OK) return;

                var statusbar = await GetServiceAsync(typeof(SVsStatusbar)) as IVsStatusbar;
                int sheets;
                var tab = await broker.GetProxyAsync<IQueryEditorTabDataServiceBrokered>(QueryEditorTabDataServiceDescriptors.QueryEditorTabDataService, DisposalToken);
                try
                {
                    if (tab == null) { Msg("SSMS result-grid service is not available."); return; }
                    sheets = await GridExporter.ExportAsync(tab, moniker, dlg.FileName,
                        text => { if (statusbar != null) statusbar.SetText(text); }, DisposalToken);
                }
                finally { (tab as IDisposable)?.Dispose(); }

                if (sheets == 0) { Msg("The active query window has no result grids. Run a query first."); return; }
                if (MessageBox.Show("Saved " + sheets + " sheet(s) to:\n" + dlg.FileName + "\n\nOpen it now?", "Export to Excel",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    Process.Start(new ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Msg("Export failed: " + ex.Message);
            }
        }

        static void Msg(string m)
        {
            MessageBox.Show(m, "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
