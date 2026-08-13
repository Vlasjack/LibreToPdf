using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace LibreOfficeToPdfGui
{
    public class MainForm : Form
    {
        private readonly TextBox inputBox = new TextBox();
        private readonly TextBox outputBox = new TextBox();
        private readonly TextBox sofficeBox = new TextBox();
        private readonly Button startButton = new Button();
        private readonly Button openOutputButton = new Button();
        private readonly Button showLogButton = new Button();
        private readonly CheckBox overwriteBox = new CheckBox();
        private readonly CheckBox compressBox = new CheckBox();
        private readonly ProgressBar progress = new ProgressBar();
        private readonly Label progressLabel = new Label();
        private readonly Label statusLabel = new Label();
        private readonly Label summaryLabel = new Label();
        private readonly Label watermarkLabel = new Label();
        private readonly TextBox logBox = new TextBox();
        private readonly Panel logPanel = new Panel();

        private bool logVisible = false;
        private int fileCount = 0;
        private string lastOutputRoot = "";
        private string ghostscriptPath = "";
        private readonly Size compactWindowSize = new Size(760, 430);
        private readonly Size logWindowSize = new Size(760, 560);

        private readonly HashSet<string> extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".odt", ".ott", ".ods", ".ots", ".odp", ".otp", ".odg", ".otg", ".odf",
            ".doc", ".docx", ".rtf", ".txt",
            ".xls", ".xlsx", ".csv",
            ".ppt", ".pptx"
        };

        public MainForm()
        {
            Text = T("LibreToPdf");
            Width = compactWindowSize.Width;
            Height = compactWindowSize.Height;
            MinimumSize = compactWindowSize;
            MaximumSize = compactWindowSize;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);

            BuildUi();
            sofficeBox.Text = FindLibreOffice();
            ghostscriptPath = FindGhostscript();
            UpdateStartButtonText();
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.Padding = new Padding(16, 12, 16, 10);
            root.ColumnCount = 1;
            root.RowCount = 11;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.BackColor = SystemColors.Window;

            var title = new Label();
            title.Text = T("LibreToPdf");
            title.Font = new Font("Segoe UI", 15F, FontStyle.Regular);
            title.AutoSize = true;
            title.Margin = new Padding(0, 0, 0, 8);
            root.Controls.Add(title, 0, 0);

            root.Controls.Add(MakePathBlock(T("\u0418\u0441\u0445\u043e\u0434\u043d\u0430\u044f \u043f\u0430\u043f\u043a\u0430"), inputBox, T("\u0412\u044b\u0431\u0440\u0430\u0442\u044c"), ChooseInput), 0, 1);
            root.Controls.Add(MakePathBlock(T("\u041f\u0430\u043f\u043a\u0430 \u0434\u043b\u044f PDF"), outputBox, T("\u0412\u044b\u0431\u0440\u0430\u0442\u044c"), ChooseOutput), 0, 2);

            overwriteBox.Text = T("\u0417\u0430\u043c\u0435\u043d\u044f\u0442\u044c \u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u044e\u0449\u0438\u0435 PDF");
            overwriteBox.AutoSize = true;
            overwriteBox.Margin = new Padding(0, 0, 0, 2);
            root.Controls.Add(overwriteBox, 0, 3);

            var optionsRow = new FlowLayoutPanel();
            optionsRow.FlowDirection = FlowDirection.LeftToRight;
            optionsRow.AutoSize = true;
            optionsRow.Dock = DockStyle.Fill;
            optionsRow.Margin = new Padding(0, 0, 0, 4);

            compressBox.Text = T("\u0421\u0436\u0430\u0442\u044c");
            compressBox.AutoSize = true;
            compressBox.Checked = true;
            compressBox.Margin = new Padding(0, 0, 0, 0);
            optionsRow.Controls.Add(compressBox);

            root.Controls.Add(optionsRow, 0, 4);

            startButton.Text = T("\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u0442\u044c");
            startButton.Height = 36;
            startButton.Width = 210;
            startButton.Anchor = AnchorStyles.None;
            startButton.Font = new Font("Segoe UI", 10F, FontStyle.Regular);
            startButton.Margin = new Padding(0, 4, 0, 4);
            startButton.Click += StartConversion;
            root.Controls.Add(startButton, 0, 5);

            var progressHeader = new TableLayoutPanel();
            progressHeader.Dock = DockStyle.Top;
            progressHeader.Height = 20;
            progressHeader.ColumnCount = 2;
            progressHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            progressHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            progressHeader.Margin = new Padding(0, 0, 0, 0);

            var progressTitle = new Label();
            progressTitle.Text = T("\u041f\u0440\u043e\u0433\u0440\u0435\u0441\u0441");
            progressTitle.AutoSize = true;
            progressTitle.Dock = DockStyle.Top;
            progressHeader.Controls.Add(progressTitle, 0, 0);

            progressLabel.Text = T("0 \u0438\u0437 0 \u2022 0%");
            progressLabel.TextAlign = ContentAlignment.MiddleRight;
            progressLabel.Dock = DockStyle.Top;
            progressLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            progressHeader.Controls.Add(progressLabel, 1, 0);
            root.Controls.Add(progressHeader, 0, 6);

            var progressRow = new Panel();
            progressRow.Dock = DockStyle.Top;
            progressRow.Height = 16;
            progressRow.Margin = new Padding(0, 0, 0, 3);

            progress.Dock = DockStyle.Fill;
            progress.Margin = new Padding(0);
            progressRow.Controls.Add(progress);
            root.Controls.Add(progressRow, 0, 7);

            statusLabel.Text = T("\u25cf \u0413\u043e\u0442\u043e\u0432\u043e \u043a \u0440\u0430\u0431\u043e\u0442\u0435");
            statusLabel.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            statusLabel.ForeColor = SystemColors.GrayText;
            statusLabel.AutoSize = true;
            statusLabel.Margin = new Padding(0, 2, 0, 0);
            root.Controls.Add(statusLabel, 0, 8);

            summaryLabel.Text = T("\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u043d\u043e: 0    \u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u043e: 0    \u041e\u0448\u0438\u0431\u043e\u043a: 0");
            summaryLabel.AutoSize = true;
            summaryLabel.Margin = new Padding(0, 0, 0, 3);
            root.Controls.Add(summaryLabel, 0, 9);

            var actionRow = new TableLayoutPanel();
            actionRow.Dock = DockStyle.Top;
            actionRow.Height = 34;
            actionRow.ColumnCount = 3;
            actionRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            actionRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            actionRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            actionRow.Margin = new Padding(0, 0, 0, 0);

            openOutputButton.Text = T("\u041e\u0442\u043a\u0440\u044b\u0442\u044c \u043f\u0430\u043f\u043a\u0443");
            openOutputButton.Enabled = false;
            openOutputButton.AutoSize = true;
            openOutputButton.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            openOutputButton.Click += OpenOutputFolder;
            actionRow.Controls.Add(openOutputButton, 0, 0);

            showLogButton.Text = T("\u041f\u043e\u043a\u0430\u0437\u0430\u0442\u044c \u0436\u0443\u0440\u043d\u0430\u043b");
            showLogButton.AutoSize = true;
            showLogButton.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            showLogButton.Click += ToggleLog;
            actionRow.Controls.Add(showLogButton, 2, 0);
            root.Controls.Add(actionRow, 0, 10);

            watermarkLabel.Text = T("\u0410\u0431\u043e\u0431\u0430");
            watermarkLabel.AutoSize = true;
            watermarkLabel.ForeColor = Color.FromArgb(255, 255, 250);
            watermarkLabel.Font = new Font("Segoe UI", 6F, FontStyle.Regular);
            watermarkLabel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            watermarkLabel.TextAlign = ContentAlignment.MiddleRight;
            watermarkLabel.Margin = new Padding(0, 0, 4, 0);
            actionRow.Controls.Add(watermarkLabel, 1, 0);

            logPanel.Visible = false;
            logPanel.Dock = DockStyle.Bottom;
            logPanel.Height = 130;
            logPanel.Padding = new Padding(0, 8, 0, 0);
            logBox.Multiline = true;
            logBox.ScrollBars = ScrollBars.Vertical;
            logBox.ReadOnly = true;
            logBox.WordWrap = false;
            logBox.Font = new Font("Consolas", 9F);
            logBox.Dock = DockStyle.Fill;
            logPanel.Controls.Add(logBox);
            root.Controls.Add(logPanel, 0, 10);
            logPanel.BringToFront();

            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            inputBox.TextChanged += delegate { InputChanged(); };
            outputBox.TextChanged += delegate { openOutputButton.Enabled = Directory.Exists(outputBox.Text.Trim()); };

            Controls.Add(root);
        }

        private Control MakePathBlock(string labelText, TextBox box, string buttonText, EventHandler click)
        {
            var block = new TableLayoutPanel();
            block.Dock = DockStyle.Top;
            block.AutoSize = false;
            block.Height = 50;
            block.ColumnCount = 2;
            block.RowCount = 2;
            block.Margin = new Padding(0, 0, 0, 4);
            block.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            block.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            block.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            block.RowStyles.Add(new RowStyle(SizeType.Absolute, 27));

            var label = new Label();
            label.Text = labelText;
            label.AutoSize = true;
            label.Margin = new Padding(0, 0, 0, 3);
            block.Controls.Add(label, 0, 0);
            block.SetColumnSpan(label, 2);

            box.Dock = DockStyle.Fill;
            box.Height = 24;
            box.Margin = new Padding(0, 0, 8, 0);
            block.Controls.Add(box, 0, 1);

            var button = new Button();
            button.Text = buttonText;
            button.Dock = DockStyle.Fill;
            button.Height = 26;
            button.Margin = new Padding(0);
            button.Click += click;
            block.Controls.Add(button, 1, 1);

            return block;
        }

        private void InputChanged()
        {
            var inputRoot = inputBox.Text.Trim();
            if (Directory.Exists(inputRoot))
            {
                if (string.IsNullOrWhiteSpace(outputBox.Text))
                {
                    outputBox.Text = Path.Combine(inputRoot, "_pdf");
                }
                fileCount = CountFiles(inputRoot, outputBox.Text.Trim());
            }
            else
            {
                fileCount = 0;
            }

            UpdateStartButtonText();
        }

        private void UpdateStartButtonText()
        {
            if (fileCount > 0)
            {
                startButton.Text = T("\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u0442\u044c ") + fileCount + T(" \u0444\u0430\u0439\u043b\u043e\u0432");
            }
            else
            {
                startButton.Text = T("\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u0442\u044c");
            }
        }

        private int CountFiles(string inputRoot, string outputRoot)
        {
            try
            {
                return Directory.GetFiles(inputRoot, "*.*", SearchOption.AllDirectories)
                    .Count(file =>
                        extensions.Contains(Path.GetExtension(file)) &&
                        !Path.GetFileName(file).StartsWith("~$") &&
                        (string.IsNullOrWhiteSpace(outputRoot) || !IsInside(file, outputRoot)));
            }
            catch
            {
                return 0;
            }
        }

        private void ChooseInput(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = T("\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043f\u0430\u043f\u043a\u0443 \u0441 \u0438\u0441\u0445\u043e\u0434\u043d\u044b\u043c\u0438 \u0444\u0430\u0439\u043b\u0430\u043c\u0438");
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    inputBox.Text = dialog.SelectedPath;
                    if (string.IsNullOrWhiteSpace(outputBox.Text))
                    {
                        outputBox.Text = Path.Combine(dialog.SelectedPath, "_pdf");
                    }
                    InputChanged();
                }
            }
        }

        private void ChooseOutput(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = T("\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u043f\u0430\u043f\u043a\u0443 \u0434\u043b\u044f PDF");
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    outputBox.Text = dialog.SelectedPath;
                    InputChanged();
                }
            }
        }

        private void ToggleLog(object sender, EventArgs e)
        {
            logVisible = !logVisible;
            logPanel.Visible = logVisible;
            showLogButton.Text = logVisible
                ? T("\u0421\u043a\u0440\u044b\u0442\u044c \u0436\u0443\u0440\u043d\u0430\u043b")
                : T("\u041f\u043e\u043a\u0430\u0437\u0430\u0442\u044c \u0436\u0443\u0440\u043d\u0430\u043b");

            var targetSize = logVisible ? logWindowSize : compactWindowSize;
            MaximumSize = targetSize;
            MinimumSize = targetSize;
            Size = targetSize;
        }

        private void OpenOutputFolder(object sender, EventArgs e)
        {
            var folder = outputBox.Text.Trim();
            if (Directory.Exists(folder))
            {
                Process.Start("explorer.exe", Quote(folder));
            }
        }

        private void StartConversion(object sender, EventArgs e)
        {
            var inputRoot = inputBox.Text.Trim();
            var outputRoot = outputBox.Text.Trim();
            var soffice = sofficeBox.Text.Trim();

            if (!Directory.Exists(inputRoot))
            {
                MessageBox.Show(this, T("\u0412\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u044e\u0449\u0443\u044e \u043f\u0430\u043f\u043a\u0443 \u0441 \u0438\u0441\u0445\u043e\u0434\u043d\u044b\u043c\u0438 \u0444\u0430\u0439\u043b\u0430\u043c\u0438."), T("\u041e\u0448\u0438\u0431\u043a\u0430"));
                return;
            }

            if (string.IsNullOrWhiteSpace(outputRoot))
            {
                outputRoot = Path.Combine(inputRoot, "_pdf");
                outputBox.Text = outputRoot;
            }

            if (!File.Exists(soffice))
            {
                MessageBox.Show(this, T("\u041d\u0435 \u043d\u0430\u0448\u0451\u043b LibreOffice. \u0423\u0441\u0442\u0430\u043d\u043e\u0432\u0438\u0442\u0435 LibreOffice \u0432 \u0441\u0442\u0430\u043d\u0434\u0430\u0440\u0442\u043d\u0443\u044e \u043f\u0430\u043f\u043a\u0443."), T("\u041e\u0448\u0438\u0431\u043a\u0430"));
                return;
            }

            startButton.Enabled = false;
            openOutputButton.Enabled = false;
            logBox.Clear();
            progress.Value = 0;
            progressLabel.Text = T("0 \u0438\u0437 0 \u2022 0%");
            statusLabel.Text = T("\u25cf \u0418\u0434\u0451\u0442 \u043a\u043e\u043d\u0432\u0435\u0440\u0442\u0430\u0446\u0438\u044f...");
            statusLabel.ForeColor = SystemColors.GrayText;
            summaryLabel.Text = T("\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u043d\u043e: 0    \u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u043e: 0    \u041e\u0448\u0438\u0431\u043e\u043a: 0");
            lastOutputRoot = outputRoot;

            var worker = new System.ComponentModel.BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += delegate
            {
                ConvertAll(inputRoot, outputRoot, soffice, overwriteBox.Checked, compressBox.Checked, worker);
            };
            worker.ProgressChanged += WorkerProgressChanged;
            worker.RunWorkerCompleted += delegate
            {
                startButton.Enabled = true;
                openOutputButton.Enabled = Directory.Exists(lastOutputRoot);
                statusLabel.Text = T("\u2713 \u0413\u043e\u0442\u043e\u0432\u043e");
                statusLabel.ForeColor = Color.ForestGreen;
            };
            worker.RunWorkerAsync();
        }

        private void WorkerProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs args)
        {
            var message = args.UserState as ProgressMessage;
            if (message == null)
            {
                return;
            }

            progress.Maximum = Math.Max(message.Total, 1);
            progress.Value = Math.Min(progress.Maximum, Math.Max(0, message.Done));

            var percent = message.Total == 0 ? 0 : (int)Math.Round((message.Done * 100.0) / message.Total);
            progressLabel.Text = message.Done + T(" \u0438\u0437 ") + message.Total + T(" \u2022 ") + percent + "%";
            summaryLabel.Text = T("\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u043d\u043e: ") + message.Ok +
                T("    \u041f\u0440\u043e\u043f\u0443\u0449\u0435\u043d\u043e: ") + message.Skipped +
                T("    \u041e\u0448\u0438\u0431\u043e\u043a: ") + message.Failed;

            if (!string.IsNullOrEmpty(message.Status))
            {
                statusLabel.Text = message.Status;
            }

            if (!string.IsNullOrEmpty(message.Log))
            {
                AppendLog(message.Log);
            }
        }

        private void ConvertAll(string inputRoot, string outputRoot, string soffice, bool overwrite, bool compressPdf, System.ComponentModel.BackgroundWorker worker)
        {
            Directory.CreateDirectory(outputRoot);

            var files = Directory.GetFiles(inputRoot, "*.*", SearchOption.AllDirectories)
                .Where(file => extensions.Contains(Path.GetExtension(file)))
                .Where(file => !Path.GetFileName(file).StartsWith("~$"))
                .Where(file => !IsInside(file, outputRoot))
                .ToList();

            var ok = 0;
            var skipped = 0;
            var failed = 0;

            Report(worker, 0, files.Count, ok, skipped, failed, T("\u041d\u0430\u0439\u0434\u0435\u043d\u043e \u0444\u0430\u0439\u043b\u043e\u0432: ") + files.Count, "");

            for (var i = 0; i < files.Count; i++)
            {
                var source = files[i];
                var relative = GetRelativePath(inputRoot, source);
                var relativeDirectory = Path.GetDirectoryName(relative);
                var targetDirectory = string.IsNullOrEmpty(relativeDirectory)
                    ? outputRoot
                    : Path.Combine(outputRoot, relativeDirectory);
                Directory.CreateDirectory(targetDirectory);

                var targetPdf = Path.Combine(targetDirectory, Path.GetFileNameWithoutExtension(source) + ".pdf");

                if (File.Exists(targetPdf) && !overwrite)
                {
                    skipped++;
                    Report(worker, i + 1, files.Count, ok, skipped, failed, null, "SKIP: " + relative);
                    continue;
                }

                if (File.Exists(targetPdf) && overwrite)
                {
                    File.Delete(targetPdf);
                }

                Report(worker, i, files.Count, ok, skipped, failed, T("\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0430\u0446\u0438\u044f: ") + relative, "CONVERT: " + relative);

                var result = ConvertOne(soffice, source, targetDirectory, targetPdf, compressPdf);
                if (result.Success)
                {
                    ok++;
                    Report(worker, i + 1, files.Count, ok, skipped, failed, null, "OK: " + targetPdf);
                }
                else
                {
                    failed++;
                    Report(worker, i + 1, files.Count, ok, skipped, failed, null, "FAILED: " + relative + Environment.NewLine + result.Message);
                }
            }

            Report(worker, files.Count, files.Count, ok, skipped, failed, T("\u2713 \u0413\u043e\u0442\u043e\u0432\u043e"), "DONE");
        }

        private void Report(System.ComponentModel.BackgroundWorker worker, int done, int total, int ok, int skipped, int failed, string status, string log)
        {
            worker.ReportProgress(done, new ProgressMessage(done, total, ok, skipped, failed, status, log));
        }

        private ConvertResult ConvertOne(string soffice, string source, string targetDirectory, string targetPdf, bool compressPdf)
        {
            var profileDirectory = Path.Combine(Path.GetTempPath(), "lo-pdf-profile-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(profileDirectory);

            var oldPythonHome = Environment.GetEnvironmentVariable("PYTHONHOME");
            var oldPythonPath = Environment.GetEnvironmentVariable("PYTHONPATH");

            try
            {
                Environment.SetEnvironmentVariable("PYTHONHOME", null);
                Environment.SetEnvironmentVariable("PYTHONPATH", null);

                var startInfo = new ProcessStartInfo();
                startInfo.FileName = soffice;
                startInfo.WorkingDirectory = Path.GetDirectoryName(soffice);
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.Arguments =
                    "--headless --nologo --nofirststartwizard --nodefault --nolockcheck --norestore " +
                    Quote("-env:UserInstallation=" + new Uri(profileDirectory).AbsoluteUri) + " " +
                    "--convert-to " + Quote(BuildConvertToArgument(source, compressPdf, false)) + " " +
                    "--outdir " + Quote(targetDirectory) + " " +
                    Quote(source);

                using (var process = Process.Start(startInfo))
                {
                    var stdout = process.StandardOutput.ReadToEnd();
                    var stderr = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (File.Exists(targetPdf))
                    {
                        return new ConvertResult(true, "");
                    }

                    return new ConvertResult(false,
                        "ExitCode: " + process.ExitCode + Environment.NewLine +
                        "STDOUT: " + stdout.Trim() + Environment.NewLine +
                        "STDERR: " + stderr.Trim());
                }
            }
            catch (Exception ex)
            {
                return new ConvertResult(false, ex.Message);
            }
            finally
            {
                Environment.SetEnvironmentVariable("PYTHONHOME", oldPythonHome);
                Environment.SetEnvironmentVariable("PYTHONPATH", oldPythonPath);

                try
                {
                    if (Directory.Exists(profileDirectory))
                    {
                        Directory.Delete(profileDirectory, true);
                    }
                }
                catch
                {
                }
            }
        }

        private static string FindLibreOffice()
        {
            var candidates = new[]
            {
                @"C:\Program Files\LibreOffice\program\soffice.com",
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.com",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };

            foreach (var candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return "";
        }

        private ConvertResult RasterizePdf(string targetPdf)
        {
            if (string.IsNullOrWhiteSpace(ghostscriptPath) || !File.Exists(ghostscriptPath))
            {
                return new ConvertResult(false,
                    T("\u0414\u043b\u044f \u0440\u0435\u0436\u0438\u043c\u0430 \u00ab\u0423\u0431\u0440\u0430\u0442\u044c \u0432\u044b\u0434\u0435\u043b\u0435\u043d\u0438\u0435 \u0442\u0435\u043a\u0441\u0442\u0430\u00bb \u043d\u0443\u0436\u0435\u043d Ghostscript: gswin64c.exe \u0438\u043b\u0438 gswin32c.exe."));
            }

            var temporaryPdf = Path.Combine(
                Path.GetDirectoryName(targetPdf),
                Path.GetFileNameWithoutExtension(targetPdf) + ".image-only.tmp.pdf");

            try
            {
                if (File.Exists(temporaryPdf))
                {
                    File.Delete(temporaryPdf);
                }

                var startInfo = new ProcessStartInfo();
                startInfo.FileName = ghostscriptPath;
                startInfo.WorkingDirectory = Path.GetDirectoryName(ghostscriptPath);
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.Arguments =
                    "-dSAFER -dBATCH -dNOPAUSE " +
                    "-sDEVICE=pdfimage24 " +
                    "-r150 " +
                    "-sCompression=JPEG -dJPEGQ=60 " +
                    "-sOutputFile=" + Quote(temporaryPdf) + " " +
                    Quote(targetPdf);

                using (var process = Process.Start(startInfo))
                {
                    var stdout = process.StandardOutput.ReadToEnd();
                    var stderr = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0 || !File.Exists(temporaryPdf))
                    {
                        return new ConvertResult(false,
                            "Ghostscript ExitCode: " + process.ExitCode + Environment.NewLine +
                            "STDOUT: " + stdout.Trim() + Environment.NewLine +
                            "STDERR: " + stderr.Trim());
                    }
                }

                File.Delete(targetPdf);
                File.Move(temporaryPdf, targetPdf);
                return new ConvertResult(true, "");
            }
            catch (Exception ex)
            {
                return new ConvertResult(false, ex.Message);
            }
            finally
            {
                try
                {
                    if (File.Exists(temporaryPdf))
                    {
                        File.Delete(temporaryPdf);
                    }
                }
                catch
                {
                }
            }
        }

        private static string FindGhostscript()
        {
            var commands = new[] { "gswin64c.exe", "gswin32c.exe", "gs.exe" };

            foreach (var command in commands)
            {
                var found = FindInPath(command);
                if (!string.IsNullOrWhiteSpace(found))
                {
                    return found;
                }
            }

            var roots = new[]
            {
                @"C:\Program Files\gs",
                @"C:\Program Files (x86)\gs"
            };

            foreach (var root in roots)
            {
                if (!Directory.Exists(root))
                {
                    continue;
                }

                foreach (var command in commands)
                {
                    var matches = Directory.GetFiles(root, command, SearchOption.AllDirectories);
                    if (matches.Length > 0)
                    {
                        return matches.OrderByDescending(x => x).First();
                    }
                }
            }

            return "";
        }

        private static string FindInPath(string fileName)
        {
            var pathValue = Environment.GetEnvironmentVariable("PATH") ?? "";
            foreach (var folder in pathValue.Split(Path.PathSeparator))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(folder))
                    {
                        continue;
                    }

                    var candidate = Path.Combine(folder.Trim(), fileName);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
                catch
                {
                }
            }

            return "";
        }

        private static string BuildConvertToArgument(string source, bool compressPdf, bool restrictCopy)
        {
            if (!compressPdf && !restrictCopy)
            {
                return "pdf";
            }

            var filter = GetPdfFilterName(source);
            var parts = new List<string>();

            if (compressPdf)
            {
                parts.Add(JsonOption("UseLosslessCompression", "boolean", "false"));
                parts.Add(JsonOption("Quality", "long", "60"));
                parts.Add(JsonOption("ReduceImageResolution", "boolean", "true"));
                parts.Add(JsonOption("MaxImageResolution", "long", "150"));
                parts.Add(JsonOption("UseTaggedPDF", "boolean", "false"));
            }

            if (restrictCopy)
            {
                parts.Add(JsonOption("RestrictPermissions", "boolean", "true"));
                parts.Add(JsonOption("PermissionPassword", "string", "change-permissions"));
                parts.Add(JsonOption("EnableCopyingOfContent", "boolean", "false"));
                parts.Add(JsonOption("EnableTextAccessForAccessibilityTools", "boolean", "false"));
                parts.Add(JsonOption("Changes", "long", "0"));
            }

            return "pdf:" + filter + ":{" + string.Join(",", parts.ToArray()) + "}";
        }

        private static string GetPdfFilterName(string source)
        {
            var extension = Path.GetExtension(source).ToLowerInvariant();

            if (extension == ".ods" || extension == ".ots" || extension == ".xls" || extension == ".xlsx" || extension == ".csv")
            {
                return "calc_pdf_Export";
            }

            if (extension == ".odp" || extension == ".otp" || extension == ".ppt" || extension == ".pptx")
            {
                return "impress_pdf_Export";
            }

            if (extension == ".odg" || extension == ".otg")
            {
                return "draw_pdf_Export";
            }

            return "writer_pdf_Export";
        }

        private static string JsonOption(string name, string type, string value)
        {
            if (type == "string")
            {
                value = "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
            }

            return "\"" + name + "\":{\"type\":\"" + type + "\",\"value\":" + value + "}";
        }

        private static bool IsInside(string file, string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
            {
                return false;
            }

            var filePath = Path.GetFullPath(file).TrimEnd('\\') + "\\";
            var folderPath = Path.GetFullPath(folder).TrimEnd('\\') + "\\";
            return filePath.StartsWith(folderPath, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetRelativePath(string root, string file)
        {
            var rootUri = new Uri(Path.GetFullPath(root).TrimEnd('\\') + "\\");
            var fileUri = new Uri(Path.GetFullPath(file));
            return Uri.UnescapeDataString(rootUri.MakeRelativeUri(fileUri).ToString()).Replace('/', '\\');
        }

        private static string Quote(string value)
        {
            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        private static string T(string value)
        {
            return value;
        }

        private void AppendLog(string text)
        {
            logBox.AppendText(text + Environment.NewLine);
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        private class ConvertResult
        {
            public bool Success { get; private set; }
            public string Message { get; private set; }

            public ConvertResult(bool success, string message)
            {
                Success = success;
                Message = message;
            }
        }

        private class ProgressMessage
        {
            public int Done { get; private set; }
            public int Total { get; private set; }
            public int Ok { get; private set; }
            public int Skipped { get; private set; }
            public int Failed { get; private set; }
            public string Status { get; private set; }
            public string Log { get; private set; }

            public ProgressMessage(int done, int total, int ok, int skipped, int failed, string status, string log)
            {
                Done = done;
                Total = total;
                Ok = ok;
                Skipped = skipped;
                Failed = failed;
                Status = status;
                Log = log;
            }
        }
    }
}
