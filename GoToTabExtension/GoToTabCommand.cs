using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Task = System.Threading.Tasks.Task;

namespace GoToTabExtension
{
    internal sealed class GoToTabCommand
    {
        private static DTE2 _dte;

        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("33022384-3a67-4fad-8541-f180369113dc");

        private Dictionary<string, DateTime> _docOpenTimes = new Dictionary<string, DateTime>();

        private GoToTabCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            if (package is null) 
                throw new ArgumentNullException(nameof(package));

            if (commandService is null) 
                throw new ArgumentNullException(nameof(commandService));

            CommandID menuCommandId = new CommandID(CommandSet, CommandId);
            MenuCommand menuItem = new MenuCommand(Execute, menuCommandId);
            commandService.AddCommand(menuItem);
            TrackWindowEvents();
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            DTE2 dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            _dte = dte;

            _ = new GoToTabCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            List<string> recentFiles = GetRecentFiles();
            ShowRecentFilesDropdown(recentFiles);
        }

        private void TrackWindowEvents()
        {
            _dte.Events.WindowEvents.WindowActivated += WindowActivated;
        }

        private void WindowActivated(EnvDTE.Window gotFocus, EnvDTE.Window _lostFocus)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var document = gotFocus.Document;
            if (document != null)
            {
                _docOpenTimes[document.FullName] = DateTime.Now;
            }
        }

        private List<string> GetRecentFiles()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (Document doc in _dte.Documents)
            {
                _docOpenTimes[doc.FullName] = _docOpenTimes.ContainsKey(doc.FullName) ? _docOpenTimes[doc.FullName] : DateTime.MinValue;
            }

            List<string> recentFiles = _docOpenTimes.OrderByDescending(pair => pair.Value)
                                                    .Select(pair => pair.Key).ToList();

            return recentFiles;
        }

        private void ShowRecentFilesDropdown(List<string> recentFiles)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            System.Windows.Window dialog = new System.Windows.Window
            {
                Title = "Active Tabs",
                Width = 700,
                Height = 600,
                WindowStartupLocation = WindowStartupLocation.CenterScreen
            };
            dialog.Content = CreateDropdownContent(recentFiles, dialog);

            dialog.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    dialog.Close();
                }
            };

            dialog.ShowDialog();
        }

        private UIElement CreateDropdownContent(List<string> recentFiles, System.Windows.Window dialog)
        {
            StackPanel stackPanel = new StackPanel();

            TextBox searchBox = new TextBox
            {
                Margin = new Thickness(10)
            };

            ListBox listBox = new ListBox
            {
                Margin = new Thickness(10),
                Height = 500,
            };

            var fileDisplayData = recentFiles.Select(filePath =>
            {
                string projectName = GetRelativeDirectoryPath(filePath); 
                return new
                {
                    FileName = System.IO.Path.GetFileName(filePath),
                    FullPath = filePath,
                    ProjectName = projectName
                };
            }).ToList();

            listBox.ItemsSource = fileDisplayData
                .Select(item => $"{item.FileName} ({item.ProjectName})")
                .ToList();

            searchBox.TextChanged += (sender, args) =>
            {
                string searchText = searchBox.Text;
                listBox.ItemsSource = fileDisplayData
                    .Where(item => item.FileName.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                   item.ProjectName.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                    .Select(item => $"{item.FileName} ({item.ProjectName})")
                    .ToList();
            };

            searchBox.PreviewKeyDown += (sender, e) =>
            {
                if (e.Key == Key.Down)
                {
                    listBox.Focus();
                }
            };

            listBox.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    ThreadHelper.ThrowIfNotOnUIThread();
                    if (listBox.SelectedItem is string selectedDisplayText)
                    {
                        var selectedFile = fileDisplayData.FirstOrDefault(item =>
                            $"{item.FileName} ({item.ProjectName})" == selectedDisplayText);

                        if (selectedFile != null)
                        {
                            _dte.ItemOperations.OpenFile(selectedFile.FullPath);
                        }
                    }
                    dialog.Close();
                }
            };

            stackPanel.Children.Add(searchBox);
            stackPanel.Children.Add(listBox);

            searchBox.Focus();

            return stackPanel;
        }

        private string GetRelativeDirectoryPath(string filePath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string solutionDir = System.IO.Path.GetDirectoryName(_dte.Solution.FullName);

            if (!solutionDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            {
                solutionDir += System.IO.Path.DirectorySeparatorChar;
            }

            string fileDir = System.IO.Path.GetDirectoryName(filePath);

            string relativePath = fileDir.Replace(solutionDir, string.Empty);

            return relativePath;
        }
    }
}
