using Microsoft.Office.Interop.Outlook;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Outlook= Microsoft.Office.Interop.Outlook;

namespace OutlookSenderStatistics
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }
        Outlook.Application? outlookApp;
        List<InboxSelection> inboxSelections = new List<InboxSelection>();
        Type? outlookType;
        private void toolStripButtonAttachToOutlook_Click(object sender, EventArgs e)
        {
            if (outlookApp != null)
            {
                MessageBox.Show("Outlook already opened, if you closed it by accident, restart the application");
                return;
            }

            try
            {
                outlookApp = Program.GetActiveObject("Outlook.Application") as Outlook.Application;
                if (outlookApp != null)
                {
                    inboxSelections.Clear();
                    var ns = outlookApp.GetNamespace("MAPI");
                    var mailFolders = ns.Folders;
                    foreach (Outlook.MAPIFolder folder in mailFolders)
                    {
                        foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                        {
                            if (subFolder.DefaultItemType == Outlook.OlItemType.olMailItem)
                            {
                                inboxSelections.Add(new InboxSelection() { IsSelected = true, Folder = subFolder.FullFolderPath, EntryId = subFolder.EntryID });
                            }
                            subFolder?.ReleaseComObject();
                        }
                        folder?.ReleaseComObject();
                    }
                    bindingSourceFolders.DataSource = this.inboxSelections;
                    mailFolders?.ReleaseComObject();
                    ns?.ReleaseComObject();
                    MessageBox.Show("Outllook opened successfully");

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(string.Format("Outlook Classic or one of its folders is not accessible, maybe it is running as administrator, having network issues or have not finished loading ({0}).", ex.Message));
            }
        }
        CancellationTokenSource? cancellationTokenSource = new CancellationTokenSource();
        Dictionary<string, SenderStatistics> senderStatistics = new Dictionary<string, SenderStatistics>();
        async Task CountMailBySenderTaskAsync(IProgress<ProgressChangedEventArgs> progress, CancellationToken token)
        {
            var ns = outlookApp.GetNamespace("MAPI");
            var mailFolders = ns.Folders;
            Dictionary<string, Outlook.MAPIFolder> mailFolderIndex = new Dictionary<string, Outlook.MAPIFolder>();

            foreach (Outlook.MAPIFolder mailFolder in mailFolders)
            {
                foreach (Outlook.MAPIFolder subFolder in mailFolder.Folders)
                {
                    if (subFolder.DefaultItemType == Outlook.OlItemType.olMailItem)
                    {
                        mailFolderIndex.Add(subFolder.EntryID, subFolder);
                    }
                }
                mailFolder?.ReleaseComObject();
            }
            var selectedFolders = inboxSelections.Where(i => i.IsSelected).ToList();
            if (selectedFolders.Count == 0) return;

            int totalFolders = selectedFolders.Count;
            int processedFolders = 0;
            foreach (var inboxSelection in selectedFolders)
            {
                if (cancellationTokenSource == null || cancellationTokenSource.Token.IsCancellationRequested)
                    break;
                var folder = mailFolderIndex[inboxSelection.EntryId];
                if (folder != null)
                {
                    var folderItems = folder.Items;
                    int mailCount = folderItems.Count;
                    int processedMails = 0;
                    int batchSize = 100;
                    int currentpositionInBatch = 0;
                    foreach (var folderItem in folderItems)
                    {
                        if (cancellationTokenSource == null || cancellationTokenSource.Token.IsCancellationRequested)
                            break;
                        Outlook.MailItem? mailItem = folderItem as Outlook.MailItem;
                        if (mailItem != null)
                        {
                            try
                            {
                                string sender = mailItem.SenderEmailAddress;
                                //some sender addresses are null
                                if (sender != null)
                                {
                                    if (senderStatistics.ContainsKey(sender))
                                    {
                                        senderStatistics[sender].MailCount += 1;
                                        senderStatistics[sender].TotalMailSize += mailItem.Size;
                                    }
                                    else
                                    {
                                        senderStatistics[sender] = new SenderStatistics
                                        {
                                            Sender = sender,
                                            MailCount = 1,
                                            TotalMailSize = mailItem.Size
                                        };
                                    }
                                }                               
                            }
                            catch (System.Exception ex)
                            {
                                Debug.Write(string.Format("Error processing mail {0}: {1}", processedMails, ex.Message));
                            }
                            finally {
                                processedMails++;
                                currentpositionInBatch++;
                                if (currentpositionInBatch >= batchSize)
                                {
                                    int mailProgress = (int)(((processedFolders + processedMails / mailCount) / (double)totalFolders) * 100);
                                    progress.Report(new ProgressChangedEventArgs(mailProgress, $"Scanned folder: {inboxSelection.Folder}, processed mails: {processedMails}/{mailCount}"));
                                    await Task.Yield(); // Yield to keep UI responsive
                                    currentpositionInBatch = 0;
                                }
                                mailItem.ReleaseComObject();
                            }
                        }
                        folderItem?.ReleaseComObject();
                    }
                    folderItems?.ReleaseComObject();
                    processedFolders++;
                    int progressPercentage = (int)((processedFolders / (double)totalFolders) * 100);
                    progress.Report(new ProgressChangedEventArgs(progressPercentage, $"Scanned folder: {inboxSelection.Folder}"));
                }
            }
            foreach (var item in mailFolderIndex.Values)
            {
                item?.ReleaseComObject();
            }
            mailFolders?.ReleaseComObject();
            ns?.ReleaseComObject();
        }


        private async void buttonStartCounting_Click(object sender, EventArgs e)
        {
            if (outlookApp == null)
            {
                MessageBox.Show("Please open Outlook first.");
                return;
            }
            // Cancel any previous operation
            cancellationTokenSource?.Cancel();
            cancellationTokenSource = new CancellationTokenSource();
            CancellationToken token = cancellationTokenSource.Token;
            senderStatistics = new Dictionary<string, SenderStatistics>();

            // Setup progress reporting
            var progress = new Progress<ProgressChangedEventArgs>(progressChangedEventArgs =>
            {
                // This code runs on the UI thread automatically
                this.toolStripProgressBar1.Value = progressChangedEventArgs.ProgressPercentage;
                this.toolStripStatusLabel1.Text = progressChangedEventArgs.UserState?.ToString() ?? "";
            });
            try
            {
                // Pass the token and progress reporter to the async task
                await Task.Run(()=>CountMailBySenderTaskAsync(progress, token));
                this.toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Operation Completed!";
            }
            catch (OperationCanceledException)
            {
                this.toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Operation Cancelled.";
            }
            catch (System.Exception ex)
            {
                this.toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = $"Error: {ex.Message}";
            }
            finally
            {
                // Re-enable buttons, clean up source
                cancellationTokenSource.Dispose();
                cancellationTokenSource = null;
            }
            if (senderStatistics?.Count > 0)
            {
                saveFileDialog1.FileName = "SenderStatistics.csv";
                bool done=false;
                while (!done)
                { 
                    var result = saveFileDialog1.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        try
                        {

                            CsvHelper.Configuration.CsvConfiguration config = new CsvHelper.Configuration.CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture)
                            {
                                Delimiter = ",",
                            };
                            using (var writer = new StreamWriter(saveFileDialog1.FileName))
                            using (var csv = new CsvHelper.CsvWriter(writer, config))
                            {
                                csv.WriteField("Sender");
                                csv.WriteField("MailCount");
                                csv.WriteField("TotalMailSize");
                                csv.WriteField("AverageMailSize");
                                csv.NextRecord();
                                foreach (var kvp in senderStatistics.OrderByDescending(kvp => kvp.Value.MailCount))
                                {
                                    csv.WriteField(kvp.Key);
                                    csv.WriteField(kvp.Value.MailCount);
                                    csv.WriteField(kvp.Value.TotalMailSize);
                                    csv.WriteField(kvp.Value.TotalMailSize / kvp.Value.MailCount);
                                    csv.NextRecord();
                                }
                            }
                            Process.Start(new ProcessStartInfo()
                            {
                                FileName = saveFileDialog1.FileName,
                                UseShellExecute = true
                            }); 
                            done = true;
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        return;
                    }
                }
            }

        }

        private void buttonStopCounting_Click(object sender, EventArgs e)
        {
            cancellationTokenSource?.Cancel();
        }


        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            cancellationTokenSource?.Cancel();
            if (outlookApp != null)
            {
                outlookApp.ReleaseComObject();
                outlookApp = null;
            }
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            outlookType = Type.GetTypeFromProgID("Outlook.Application", true);
            if (outlookType == null)
            {
                MessageBox.Show("Unable to find Outlook from registry");
                this.Close();
            }
        }

        private async void buttonStartDeletion_Click(object sender, EventArgs e)
        {
            if (outlookApp == null)
            {
                MessageBox.Show("Please open Outlook first.");
                return;
            }
            if (string.IsNullOrEmpty(this.textBoxAddressesToDelete.Text))
            {
                MessageBox.Show("Please enter at least one email address to delete.");
                return;
            }
            int minAgeToDelete = 
                numericUpDownMinAgeToDelete.Value > 0 ? (int)numericUpDownMinAgeToDelete.Value:30;

            int? minSizeToDelete= checkBoxMinSizeToDelete.Checked ?
                numericUpDownMinSizeToDelete.Value > 0 ? (int)numericUpDownMinSizeToDelete.Value*1024 : 1024
                : null;
            bool canMatchPartially=this.checkBoxPartialMatchAddress.Checked;
            var senderAddressesToDelete = this.textBoxAddressesToDelete.Lines
                .Select(line => line.Trim())
                .Where(line => !string.IsNullOrEmpty(line))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            // Cancel any previous operation
            cancellationTokenSource?.Cancel();
            cancellationTokenSource = new CancellationTokenSource();
            CancellationToken token = cancellationTokenSource.Token;
            senderStatistics = new Dictionary<string, SenderStatistics>();

            // Setup progress reporting
            var progress = new Progress<ProgressChangedEventArgs>(progressChangedEventArgs =>
            {
                // This code runs on the UI thread automatically
                this.toolStripProgressBar1.Value = progressChangedEventArgs.ProgressPercentage;
                this.toolStripStatusLabel1.Text = progressChangedEventArgs.UserState?.ToString() ?? "";
            });
            try
            {
                // Pass the token and progress reporter to the async task
                await Task.Run(()=> DeleteMailBySenderTaskAsync(progress, minAgeToDelete, minSizeToDelete, canMatchPartially, senderAddressesToDelete,token));
                this.toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Operation Completed! If you want to undelete, check the Deleted Items folder in Outlook Classic";
            }
            catch (OperationCanceledException)
            {
                this.toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Operation Cancelled.";
            }
            catch (System.Exception ex)
            {
                this.toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = $"Error: {ex.Message}";
            }
            finally
            {
                // Re-enable buttons, clean up source
                cancellationTokenSource.Dispose();
                cancellationTokenSource = null;
            }
        }
        public async Task DeleteMailBySenderTaskAsync(IProgress<ProgressChangedEventArgs> progress, int minAgeToDelete, int? minSizeToDelete, bool canMatchPartially, HashSet<string> senderAddressesToDelete, CancellationToken token)
        {
            var ns = outlookApp.GetNamespace("MAPI");
            var mailFolders = ns.Folders; 
            
            Dictionary<string, Outlook.MAPIFolder> mailFolderIndex = new Dictionary<string, Outlook.MAPIFolder>();

            foreach (Outlook.MAPIFolder mailFolder in mailFolders)
            {
                foreach (Outlook.MAPIFolder subFolder in mailFolder.Folders)
                {
                    if (subFolder.DefaultItemType == Outlook.OlItemType.olMailItem)
                    {
                        mailFolderIndex.Add(subFolder.EntryID, subFolder);
                    }
                }
                mailFolder?.ReleaseComObject();
            }

            var totalFolders = mailFolderIndex.Count;
            int processedFolders = 0;
            var currentDateTime = DateTime.Now;

            foreach (var subFolder in mailFolderIndex.Values)
            {
                var folderItems = subFolder.Items;
                int mailCount = folderItems.Count;
                int processedMails = 0;
                int batchSize = 100;
                int currentpositionInBatch = 0;


                foreach (var folderItem in folderItems)
                {
                    Outlook.MailItem ? mailItem= folderItem as Outlook.MailItem;
                    if (cancellationTokenSource == null || cancellationTokenSource.Token.IsCancellationRequested)
                        break;
                    if (mailItem != null)
                    {                        
                        try
                        {
                            string sender = mailItem.SenderEmailAddress;
                            if (sender != null)
                            {
                                if (canMatchPartially)
                                {
                                    var matchesAny = senderAddressesToDelete.Any(address => sender.IndexOf(address, StringComparison.OrdinalIgnoreCase) >= 0);
                                    if (!matchesAny)
                                    {
                                        processedMails++;
                                        continue;
                                    }
                                }
                                else
                                {
                                    if (!senderAddressesToDelete.Contains(sender))
                                    {
                                        processedMails++;
                                        continue;
                                    }
                                }
                            }
                            else
                                continue;//some sender addresses are null

                            var isOldEnoughToDelete = currentDateTime > mailItem.ReceivedTime.AddDays(minAgeToDelete);
                            bool isLargeEnoughToDelete = false;
                            if (minSizeToDelete.HasValue)
                            {
                                int mailSize = mailItem.Size;
                                isLargeEnoughToDelete = mailSize >= minSizeToDelete.Value;
                            }
                            if (isOldEnoughToDelete && isLargeEnoughToDelete)
                            {

                                Debug.WriteLine(string.Format("Deleting mail from {0}\t\t\t {1} \t\t\t received on {2}", sender, mailItem.Subject, mailItem.ReceivedTime));
                                mailItem.Delete();
                            }             

                        }
                        catch (System.Exception ex)
                        {
                            Debug.Write(string.Format("Error processing mail {0}: {1}", processedMails, ex.Message));
                            int mailProgress = (int)(((processedFolders + processedMails / mailCount) / (double)totalFolders) * 100);

                            progress.Report(new ProgressChangedEventArgs(mailProgress, $"Error deleting Mail: {ex.Message}, processed mails: {processedMails}/{mailCount}"));
                            await Task.Yield(); // Yield to keep UI responsive

                        }
                        finally
                        {
                            processedMails++;
                            currentpositionInBatch++;
                            if (currentpositionInBatch >= batchSize)
                            {
                                int mailProgress = (int)(((processedFolders + processedMails / mailCount) / (double)totalFolders) * 100);
                                progress.Report(new ProgressChangedEventArgs(mailProgress, $"deleting mails from folder: {subFolder.FullFolderPath}, processed mails: {processedMails}/{mailCount}"));
                                await Task.Yield(); // Yield to keep UI responsive
                                currentpositionInBatch = 0;
                            }
                            mailItem.ReleaseComObject();
                        }
                    }
                    folderItem?.ReleaseComObject();
                }

                folderItems?.ReleaseComObject();
                processedFolders++;
                int progressPercentage = (int)((processedFolders / (double)totalFolders) * 100);
                progress.Report(new ProgressChangedEventArgs(progressPercentage, $"Cleaned folder: {subFolder.FullFolderPath}"));
                subFolder?.ReleaseComObject();
            }
            mailFolders?.ReleaseComObject();
            ns?.ReleaseComObject();
        }

        private void buttonStopDeletion_Click(object sender, EventArgs e)
        {
            cancellationTokenSource?.Cancel();
        }
    }
}
