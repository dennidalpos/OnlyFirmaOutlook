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
using OnlyFirmaOutlook.Models;
using OnlyFirmaOutlook.Services;
using MessageBox = System.Windows.MessageBox;

namespace OnlyFirmaOutlook.Views;

public partial class MainWindow
{
    private void RefreshExistingSignatures()
    {
        var folderPath = DestinationFolderTextBox.Text;

        if (string.IsNullOrEmpty(folderPath))
        {
            _existingSignatures = new List<SignatureInfo>();
        }
        else
        {
            _existingSignatures = _signatureRepository.GetSignatures(folderPath);
        }

        if (_existingSignatures.Count > 0)
        {
            ExistingSignaturesListBox.ItemsSource = _existingSignatures;
            ExistingSignaturesListBox.DisplayMemberPath = "DisplayInfo";
            NoSignaturesText.Visibility = Visibility.Collapsed;
        }
        else
        {
            ExistingSignaturesListBox.ItemsSource = null;
            NoSignaturesText.Visibility = Visibility.Visible;
        }

        DeleteSignatureButton.IsEnabled = false;
        CheckOverwriteWarning();
    }

    private void RefreshBackups()
    {
        var backupsFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
        BackupFolderText.Text = $"Cartella backup: {backupsFolder}";

        _existingBackups = _signatureRepository.GetBackups(backupsFolder);

        if (_existingBackups.Count > 0)
        {
            BackupsListBox.ItemsSource = _existingBackups;
            BackupsListBox.DisplayMemberPath = "DisplayInfo";
            NoBackupsText.Visibility = Visibility.Collapsed;
        }
        else
        {
            BackupsListBox.ItemsSource = null;
            NoBackupsText.Visibility = Visibility.Visible;
        }

        UpdateBackupButtons();
    }

    private void UpdateBackupButtons()
    {
        var hasSelection = BackupsListBox.SelectedItem != null;
        RestoreBackupButton.IsEnabled = hasSelection;
        DeleteBackupButton.IsEnabled = hasSelection;
    }

    private void CheckOverwriteWarning()
    {
        var baseName = SignatureNameTextBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(baseName))
        {
            OverwriteWarningBorder.Visibility = Visibility.Collapsed;
            return;
        }

        var finalName = _signatureWorkflowService.BuildFinalSignatureName(baseName, GetCurrentSignatureIdentifier());
        var destinationFolder = DestinationFolderTextBox.Text;

        if (!string.IsNullOrEmpty(destinationFolder) &&
            _signatureWorkflowService.SignatureExists(destinationFolder, finalName))
        {
            OverwriteWarningText.Text = $"La firma '{finalName}' esiste già e verrà sovrascritta!";
            OverwriteWarningBorder.Visibility = Visibility.Visible;
        }
        else
        {
            OverwriteWarningBorder.Visibility = Visibility.Collapsed;
        }
    }

    private void UpdateFinalSignatureName()
    {
        var baseName = SignatureNameTextBox.Text?.Trim() ?? string.Empty;

        if (string.IsNullOrEmpty(baseName))
        {
            FinalNameBorder.Visibility = Visibility.Collapsed;
            return;
        }

        var finalName = _signatureWorkflowService.BuildFinalSignatureName(baseName, GetCurrentSignatureIdentifier());

        FinalSignatureNameText.Text = finalName;
        FinalNameBorder.Visibility = Visibility.Visible;
    }

    private string? GetCurrentSignatureIdentifier()
    {
        if (_isOutlookAvailable && AccountComboBox.SelectedItem is OutlookAccount account)
        {
            return account.DisplayText;
        }

        if (!string.IsNullOrWhiteSpace(IdentifierTextBox.Text))
        {
            return IdentifierTextBox.Text.Trim();
        }

        return null;
    }

    private void SignatureNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateConvertButtonState();
    }

    private void IdentifierTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateConvertButtonState();
    }

    private void AccountComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateConvertButtonState();
    }

    private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Seleziona cartella di destinazione per le firme",
            ShowNewFolderButton = true
        };

        if (!string.IsNullOrEmpty(DestinationFolderTextBox.Text) &&
            Directory.Exists(DestinationFolderTextBox.Text))
        {
            dialog.SelectedPath = DestinationFolderTextBox.Text;
        }

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            DestinationFolderTextBox.Text = dialog.SelectedPath;
            ValidateDestinationFolder(dialog.SelectedPath);
            RefreshExistingSignatures();
        }
    }

    private void ExistingSignaturesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        DeleteSignatureButton.IsEnabled = ExistingSignaturesListBox.SelectedItem != null;
    }

    private void BackupsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateBackupButtons();
    }

    private void RestoreBackupButton_Click(object sender, RoutedEventArgs e)
    {
        if (BackupsListBox.SelectedItem is not BackupInfo backup)
        {
            return;
        }

        var result = MessageBox.Show(
            $"Ripristinare il backup '{backup.FileName}'?\n\n" +
            "I file nella cartella firme verranno sovrascritti.",
            "Conferma ripristino",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        SetBusy(true, "Ripristino backup in corso...");

        try
        {
            var destinationFolder = SignatureRepository.GetDefaultOutlookSignaturesFolder();
            var restored = _signatureRepository.RestoreBackup(backup, destinationFolder);

            if (restored)
            {
                _logger.Log($"Backup ripristinato: {backup.FileName}");
                MessageBox.Show("Backup ripristinato con successo.", "Ripristino completato", MessageBoxButton.OK,
                    MessageBoxImage.Information);
                RefreshExistingSignatures();
            }
            else
            {
                MessageBox.Show("Impossibile ripristinare il backup selezionato.", "Errore ripristino",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void DeleteBackupButton_Click(object sender, RoutedEventArgs e)
    {
        if (BackupsListBox.SelectedItem is not BackupInfo backup)
        {
            return;
        }

        var result = MessageBox.Show(
            $"Eliminare il backup '{backup.FileName}'?",
            "Conferma eliminazione",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            if (_signatureRepository.DeleteBackup(backup))
            {
                _logger.Log($"Backup eliminato: {backup.FileName}");
            }
            else
            {
                MessageBox.Show("Impossibile eliminare il backup selezionato.", "Errore eliminazione",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            RefreshBackups();
        }
    }

    private void RefreshBackupsButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshBackups();
    }

    private void DeleteSignatureButton_Click(object sender, RoutedEventArgs e)
    {
        if (ExistingSignaturesListBox.SelectedItem is not SignatureInfo signature)
        {
            return;
        }

        var result = MessageBox.Show(
            $"Eliminare la firma '{signature.Name}'?\n\n" +
            "Verranno eliminati tutti i file associati (.htm, .rtf, .txt e cartelle assets).",
            "Conferma eliminazione",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            _signatureRepository.DeleteSignature(signature);
            RefreshExistingSignatures();
            _logger.Log($"Firma '{signature.Name}' eliminata");
        }
        catch (Exception ex)
        {
            _logger.LogError("Errore durante l'eliminazione della firma", ex);
            MessageBox.Show(
                $"Errore durante l'eliminazione:\n{ex.Message}",
                "Errore",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void RefreshSignaturesButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshExistingSignatures();
    }

    private void BrowseSignaturesButton_Click(object sender, RoutedEventArgs e)
    {
        OpenSelectedDestinationFolder();
    }

    private void BrowseBackupsButton_Click(object sender, RoutedEventArgs e)
    {
        OpenBackupsFolder();
    }

    private void OpenSelectedDestinationFolder()
    {
        var destinationFolder = DestinationFolderTextBox.Text?.Trim();

        if (string.IsNullOrWhiteSpace(destinationFolder))
        {
            MessageBox.Show("Seleziona prima una cartella di destinazione valida.", "Cartella non disponibile",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        OpenFolderInExplorer(
            destinationFolder,
            "La cartella di destinazione selezionata non è disponibile.",
            "destinazione selezionata");
    }

    private void OpenBackupsFolder()
    {
        OpenFolderInExplorer(
            SignatureRepository.GetDefaultOutlookSignaturesFolder(),
            "La cartella backup non è disponibile.",
            "backup firme");
    }

    private void OpenFolderInExplorer(string folderPath, string missingMessage, string logContext)
    {
        try
        {
            if (Directory.Exists(folderPath))
            {
                Process.Start("explorer.exe", folderPath);
            }
            else
            {
                MessageBox.Show(missingMessage, "Cartella non trovata",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile aprire la cartella {logContext}: {ex.Message}");
        }
    }
}
