/*
 * OnlyFirmaOutlook
 * Copyright (c) 2026 Danny Perondi. All rights reserved.
 * Author: Danny Perondi
 * Proprietary and confidential.
 * Unauthorized copying, modification, distribution, sublicensing, disclosure,
 * or commercial use is prohibited without prior written permission.
 */

using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using Clipboard = System.Windows.Clipboard;

namespace OnlyFirmaOutlook.Views;

public partial class MainWindow
{
    private void CopyLogButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(LogTextBox.Text);
            _logger.Log("Log copiato negli appunti");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile copiare negli appunti: {ex.Message}");
        }
    }

    private void ClearLogButton_Click(object sender, RoutedEventArgs e)
    {
        _logger.Clear();
        LogTextBox.Clear();
        _logger.Log("Log pulito");
    }

    private void OpenLogFileButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var logFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OnlyFirmaOutlook",
                "Logs");

            var logFile = Path.Combine(logFolder, "app.log");

            if (File.Exists(logFile))
            {
                Process.Start("notepad.exe", logFile);
            }
            else
            {
                Process.Start("explorer.exe", logFolder);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile aprire il file di log: {ex.Message}");
        }
    }

    private void SetBusy(bool isBusy, string? message = null)
    {
        BusyOverlay.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

        if (!string.IsNullOrEmpty(message))
        {
            BusyMessage.Text = message;
        }

        PresetListBox.IsEnabled = !isBusy;
        LoadCustomButton.IsEnabled = !isBusy;
        SignatureNameTextBox.IsEnabled = !isBusy;
        AccountComboBox.IsEnabled = !isBusy;
        IdentifierTextBox.IsEnabled = !isBusy;
        BrowseFolderButton.IsEnabled = !isBusy;
        FilteredHtmlRadio.IsEnabled = !isBusy;
        CompleteHtmlRadio.IsEnabled = !isBusy;

        if (isBusy)
        {
            ConvertButton.IsEnabled = false;
        }
        else
        {
            UpdateConvertButtonState();
        }

        DeleteSignatureButton.IsEnabled = !isBusy && ExistingSignaturesListBox.SelectedItem != null;
        RefreshSignaturesButton.IsEnabled = !isBusy;
        BrowseSignaturesButton.IsEnabled = !isBusy;
        RestoreBackupButton.IsEnabled = !isBusy && BackupsListBox.SelectedItem != null;
        DeleteBackupButton.IsEnabled = !isBusy && BackupsListBox.SelectedItem != null;
        RefreshBackupsButton.IsEnabled = !isBusy;
        BrowseBackupsButton.IsEnabled = !isBusy;
    }

    private void OnLogAdded(object? sender, string message)
    {
        Dispatcher.InvokeAsync(() =>
        {
            LogTextBox.AppendText(message + Environment.NewLine);
            ScrollLogToEnd();
        });
    }

    private void ScrollLogToEnd()
    {
        LogTextBox.ScrollToEnd();
    }

    private void ResetUiForNewSignature()
    {
        _selectedFilePath = null;
        _currentEditorState = null;
        _isWordOpen = false;

        PresetListBox.SelectedItem = null;
        SelectedFileText.Text = "Nessun file selezionato";
        SignatureNameTextBox.Text = string.Empty;
        IdentifierTextBox.Text = string.Empty;
        AccountComboBox.SelectedIndex = _accounts.Count > 0 ? 0 : -1;
        FilteredHtmlRadio.IsChecked = false;
        CompleteHtmlRadio.IsChecked = true;
        ExistingSignaturesListBox.SelectedItem = null;
        BackupsListBox.SelectedItem = null;

        UpdateWordOpenIndicator();
        RefreshExistingSignatures();
        RefreshBackups();
        UpdateConvertButtonState();
    }

    private void GuideToggleButton_Checked(object sender, RoutedEventArgs e)
    {
        if (_guideWindow == null)
        {
            _guideWindow = new GuideWindow
            {
                Owner = this
            };
            _guideWindow.Closed += GuideWindow_Closed;
        }

        _guideWindow.Show();
        _guideWindow.Activate();
    }

    private void GuideToggleButton_Unchecked(object sender, RoutedEventArgs e)
    {
        if (_guideWindow != null)
        {
            _guideWindow.Closed -= GuideWindow_Closed;
            _guideWindow.Close();
            _guideWindow = null;
        }
    }

    private void GuideWindow_Closed(object? sender, EventArgs e)
    {
        if (_guideWindow != null)
        {
            _guideWindow.Closed -= GuideWindow_Closed;
            _guideWindow = null;
        }

        GuideToggleButton.IsChecked = false;
    }
}
