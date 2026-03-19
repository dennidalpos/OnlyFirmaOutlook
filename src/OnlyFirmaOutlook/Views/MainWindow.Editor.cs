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
using System.Windows.Controls;
using System.Windows.Threading;
using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace OnlyFirmaOutlook.Views;

public partial class MainWindow
{
    private void OpenWordDocument(string filePath)
    {
        try
        {
            _logger.Log($"Apertura documento in Word: {filePath}");

            var startInfo = new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            Process.Start(startInfo);

            if (_currentEditorState != null)
            {
                EditorStateTransitions.MarkDocumentOpened(_currentEditorState);
            }

            StartFileWatcher(filePath);
            StartWordCheckTimer();

            _isWordOpen = true;
            UpdateWordOpenIndicator();

            _logger.Log("Word avviato - in attesa di modifiche e chiusura");
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante l'apertura di Word", ex);
            MessageBox.Show(
                $"Impossibile aprire Word:\n\n{ex.Message}",
                "Errore Word",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void StartFileWatcher(string filePath)
    {
        StopFileWatcher();

        try
        {
            var directory = Path.GetDirectoryName(filePath);
            var fileName = Path.GetFileName(filePath);

            if (string.IsNullOrEmpty(directory))
            {
                return;
            }

            _fileWatcher = new FileSystemWatcher(directory, fileName)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
                EnableRaisingEvents = true
            };

            _fileWatcher.Changed += OnFileChanged;
            _fileWatcher.Created += OnFileChanged;
            _fileWatcher.Renamed += OnFileChanged;

            _logger.Log($"FileWatcher avviato per: {fileName}");
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile avviare FileWatcher: {ex.Message}");
        }
    }

    private void StopFileWatcher()
    {
        if (_fileWatcher != null)
        {
            _fileWatcher.EnableRaisingEvents = false;
            _fileWatcher.Changed -= OnFileChanged;
            _fileWatcher.Created -= OnFileChanged;
            _fileWatcher.Renamed -= OnFileChanged;
            _fileWatcher.Dispose();
            _fileWatcher = null;
        }
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        Dispatcher.InvokeAsync(() =>
        {
            try
            {
                if (_currentEditorState == null)
                {
                    return;
                }

                if (!File.Exists(_currentEditorState.LocalFilePath))
                {
                    _logger.LogWarning("File temporaneo non trovato durante il controllo modifica.");
                    return;
                }

                var currentModTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);
                if (EditorStateTransitions.TryMarkDocumentSaved(
                        _currentEditorState,
                        currentModTime,
                        ref _lastFileModifiedTime))
                {
                    _logger.Log("Documento salvato in Word - rilevata modifica file");
                    UpdateConvertButtonState();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Errore durante il controllo modifica file: {ex.Message}");
            }
        });
    }

    private void StartWordCheckTimer()
    {
        StopWordCheckTimer();

        _wordCheckTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _wordCheckTimer.Tick += OnWordCheckTimerTick;
        _wordCheckTimer.Start();
    }

    private void StopWordCheckTimer()
    {
        if (_wordCheckTimer != null)
        {
            _wordCheckTimer.Stop();
            _wordCheckTimer.Tick -= OnWordCheckTimerTick;
            _wordCheckTimer = null;
        }
    }

    private void OnWordCheckTimerTick(object? sender, EventArgs e)
    {
        if (_currentEditorState == null)
        {
            return;
        }

        var isWordStillOpen = IsFileLockedByWord(_currentEditorState.LocalFilePath);

        if (_isWordOpen && !isWordStillOpen)
        {
            _isWordOpen = false;
            _logger.Log("Word chiuso - documento non più in editing");

            StopWordCheckTimer();
            StopFileWatcher();

            CheckFinalFileState();

            UpdateWordOpenIndicator();
            UpdateConvertButtonState();
        }
    }

    private bool IsFileLockedByWord(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore verifica lock file: {ex.Message}");
            return false;
        }
    }

    private void CheckFinalFileState()
    {
        if (_currentEditorState == null)
        {
            return;
        }

        try
        {
            if (!File.Exists(_currentEditorState.LocalFilePath))
            {
                _logger.LogWarning("File temporaneo non trovato durante verifica finale.");
                return;
            }

            var currentModTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);
            if (EditorStateTransitions.TryMarkDocumentSaved(
                    _currentEditorState,
                    currentModTime,
                    ref _lastFileModifiedTime))
            {
                _logger.Log("Verifica finale: documento risulta salvato");
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante verifica finale: {ex.Message}");
        }
    }

    private void UpdateWordOpenIndicator()
    {
        WordOpenIndicator.Visibility = _isWordOpen ? Visibility.Visible : Visibility.Collapsed;
    }

    private void EditSignatureButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentEditorState == null || string.IsNullOrEmpty(_selectedFilePath))
        {
            MessageBox.Show(
                "Nessuna firma selezionata da modificare.",
                "Attenzione",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        if (!_wordEditorService.ValidateEditorState(_currentEditorState))
        {
            MessageBox.Show(
                "Il file temporaneo non esiste più. Seleziona nuovamente la firma.",
                "File non trovato",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            _currentEditorState = null;
            UpdateConvertButtonState();
            return;
        }

        OpenWordDocument(_currentEditorState.LocalFilePath);
    }

    private void PresetListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PresetListBox.SelectedItem is not PresetFile preset)
        {
            return;
        }

        LoadDocumentForEditing(
            preset.FullPath,
            preset.FileName,
            preset.DisplayName,
            "Preset");
    }

    private void LoadCustomButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Seleziona documento",
            Filter = "Documenti supportati (*.docx;*.doc;*.rtf)|*.docx;*.doc;*.rtf|Documenti Word (*.docx;*.doc)|*.docx;*.doc|File RTF (*.rtf)|*.rtf",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            PresetListBox.SelectedItem = null;
            var fileName = Path.GetFileName(dialog.FileName);
            var proposedName = Path.GetFileNameWithoutExtension(fileName);

            LoadDocumentForEditing(
                dialog.FileName,
                fileName,
                proposedName,
                "File personalizzato");
        }
    }

    private void LoadDocumentForEditing(string sourceFilePath, string displayFileName, string proposedSignatureName, string sourceLabel)
    {
        try
        {
            if (!File.Exists(sourceFilePath))
            {
                MessageBox.Show(
                    "Il file selezionato non esiste più. Riprova.",
                    "File non trovato",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!IsSupportedDocument(sourceFilePath))
            {
                MessageBox.Show(
                    "Il file selezionato non è un documento supportato (.doc/.docx/.rtf).",
                    "File non valido",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var normalizedName = WordConversionService.SanitizeFileName(proposedSignatureName);
            if (!string.Equals(proposedSignatureName, normalizedName, StringComparison.Ordinal))
            {
                _logger.LogWarning($"Nome firma normalizzato in import: '{proposedSignatureName}' → '{normalizedName}'");
            }

            var editableSource = PrepareSourceForEditing(sourceFilePath);

            _currentEditorState = _wordEditorService.PrepareFileForEditing(editableSource, normalizedName);
            _selectedFilePath = _currentEditorState.LocalFilePath;
            _lastFileModifiedTime = File.GetLastWriteTime(_currentEditorState.LocalFilePath);

            SelectedFileText.Text = displayFileName;
            SignatureNameTextBox.Text = normalizedName;

            _logger.Log($"{sourceLabel} importato: {displayFileName}");

            UpdateConvertButtonState();
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante l'import del documento", ex);
            MessageBox.Show(
                $"Errore durante l'import del documento:\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private string PrepareSourceForEditing(string sourceFilePath)
    {
        if (TempFileManager.IsUncPath(sourceFilePath))
        {
            _logger.Log("File su rete: copia in temporanea locale.");
            return _tempFileManager.CopyToLocalTemp(sourceFilePath);
        }

        return sourceFilePath;
    }

    private static bool IsSupportedDocument(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        return extension.Equals(".doc", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".rtf", StringComparison.OrdinalIgnoreCase);
    }
}
