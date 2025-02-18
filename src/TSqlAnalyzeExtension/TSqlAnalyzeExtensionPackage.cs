using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Process = System.Diagnostics.Process;
using Task = System.Threading.Tasks.Task;

namespace TSqlAnalyzeExtension
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(TSqlAnalyzeExtensionPackage.PackageGuidString)]
    public sealed class TSqlAnalyzeExtensionPackage : AsyncPackage
    {
         /// <summary>
        /// TSqlAnalyzeExtensionPackage GUID string.
        /// </summary>
        private const string PackageGuidString = "98e76c0f-47b2-42f9-92c6-a54b5097b962";

        private DTE2 _dte;

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            Instance = this;

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _dte = await GetServiceAsync(typeof(DTE)) as DTE2;

            if (await GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService mcs)
            {
                var menuCommandId = new CommandID(new Guid(PackageGuidString), 0x0100);
                var menuItem = new MenuCommand(ExecuteAnalysis, menuCommandId);
                mcs.AddCommand(menuItem);
            }

            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
        }

        public TSqlAnalyzeExtensionPackage Instance { get; set; }

        #endregion

        private void ExecuteAnalysis(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                // Get the active document
                var document = _dte.ActiveDocument;
                if (document == null)
                {
                    MessageBox.Show("No active document found.");
                    return;
                }

                // Verify it's a SQL query window
                if (document.Object("TextDocument") is not TextDocument textDocument)
                {
                    MessageBox.Show("Active document is not a text document.");
                    return;
                }

                // Get the current query text
                var selection = textDocument.Selection;
                var queryText = selection.Text;
                if (string.IsNullOrEmpty(queryText))
                {
                    // If no selection, get entire document
                    var editPoint = textDocument.StartPoint.CreateEditPoint();
                    queryText = editPoint.GetText(textDocument.EndPoint);
                }

                // Save to temporary file
                var tempFile = Path.GetRandomFileName() + ".sql";
                File.WriteAllText(tempFile, queryText);

                // Run the analyzer
                var startInfo = new ProcessStartInfo
                {
                    FileName = "tsqlanalyze",
                    Arguments = $"-i \"{tempFile}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    var output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    // Show results in a new document window
                    var resultWindow = _dte.ItemOperations.NewFile(
                        "General\\Text File",
                        "Analysis Results.txt",
                        EnvDTE.Constants.vsViewKindTextView);

                    var resultDoc = resultWindow.Document.Object("TextDocument") as TextDocument;
                    var editPoint = resultDoc.StartPoint.CreateEditPoint();
                    editPoint.Insert(output);
                }

                // Clean up temp file
                File.Delete(tempFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error executing analysis: {ex.Message}");
            }
        }
    }
}
