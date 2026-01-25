using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OnlyFirmaOutlook.Services;
using WpfUserControl = System.Windows.Controls.UserControl;
using WpfTextBox = System.Windows.Controls.TextBox;
using WpfLabel = System.Windows.Controls.Label;

namespace OnlyFirmaOutlook.Controls;

/// <summary>
/// Toolbar personalizzata per controllare Word via COM.
/// Non usa Ribbon Office, solo comandi COM diretti.
/// </summary>
public partial class WordToolbar : WpfUserControl
{
    private readonly LoggingService _logger;
    private dynamic? _wordDocument;
    private dynamic? _wordApplication;

    // Costanti Word per formattazione
    private const int WdAlignParagraphLeft = 0;
    private const int WdAlignParagraphCenter = 1;
    private const int WdAlignParagraphRight = 2;

    public event EventHandler? DocumentModified;

    public WordToolbar()
    {
        InitializeComponent();
        _logger = LoggingService.Instance;

        // Carica font di sistema
        LoadSystemFonts();
    }

    private void LoadSystemFonts()
    {
        try
        {
            var fonts = new System.Drawing.Text.InstalledFontCollection();
            var fontNames = fonts.Families
                .Select(f => f.Name)
                .OrderBy(n => n)
                .ToList();

            FontComboBox.ItemsSource = fontNames;

            // Seleziona Calibri o Arial come default
            var defaultFont = fontNames.FirstOrDefault(f => f == "Calibri")
                           ?? fontNames.FirstOrDefault(f => f == "Arial")
                           ?? fontNames.FirstOrDefault();

            if (defaultFont != null)
            {
                FontComboBox.SelectedItem = defaultFont;
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Errore durante il caricamento dei font: {ex.Message}");
        }
    }

    public void SetWordDocument(dynamic? document, dynamic? application)
    {
        _wordDocument = document;
        _wordApplication = application;
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        var hasDocument = _wordDocument != null;

        FontComboBox.IsEnabled = hasDocument;
        FontSizeComboBox.IsEnabled = hasDocument;
        BoldButton.IsEnabled = hasDocument;
        ItalicButton.IsEnabled = hasDocument;
        UnderlineButton.IsEnabled = hasDocument;
        FontColorButton.IsEnabled = hasDocument;
        AlignLeftButton.IsEnabled = hasDocument;
        AlignCenterButton.IsEnabled = hasDocument;
        AlignRightButton.IsEnabled = hasDocument;
        InsertLinkButton.IsEnabled = hasDocument;
        InsertImageButton.IsEnabled = hasDocument;
        UndoButton.IsEnabled = hasDocument;
        RedoButton.IsEnabled = hasDocument;
        ZoomInButton.IsEnabled = hasDocument;
        ZoomOutButton.IsEnabled = hasDocument;
    }

    private void NotifyDocumentModified()
    {
        DocumentModified?.Invoke(this, EventArgs.Empty);
    }

    #region Font Handlers

    private void FontComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_wordDocument == null || FontComboBox.SelectedItem == null) return;

        try
        {
            var fontName = FontComboBox.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(fontName)) return;

            var wordDoc = _wordDocument;
            if (wordDoc == null) return;

            var selection = wordDoc.Application.Selection;
            selection.Font.Name = fontName;
            NotifyDocumentModified();
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante il cambio font: {ex.Message}");
        }
    }

    private void FontSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            var sizeText = (FontSizeComboBox.SelectedItem as ComboBoxItem)?.Content?.ToString()
                        ?? FontSizeComboBox.Text;

            if (int.TryParse(sizeText, out var size))
            {
                var selection = _wordDocument.Application.Selection;
                selection.Font.Size = size;
                NotifyDocumentModified();
            }
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante il cambio dimensione font: {ex.Message}");
        }
    }

    #endregion

    #region Formattazione Handlers

    private void BoldButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            var selection = _wordDocument.Application.Selection;
            selection.Font.Bold = BoldButton.IsChecked == true ? -1 : 0;
            NotifyDocumentModified();
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante l'applicazione grassetto: {ex.Message}");
        }
    }

    private void ItalicButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            var selection = _wordDocument.Application.Selection;
            selection.Font.Italic = ItalicButton.IsChecked == true ? -1 : 0;
            NotifyDocumentModified();
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante l'applicazione corsivo: {ex.Message}");
        }
    }

    private void UnderlineButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            var selection = _wordDocument.Application.Selection;
            // 1 = wdUnderlineSingle, 0 = wdUnderlineNone
            selection.Font.Underline = UnderlineButton.IsChecked == true ? 1 : 0;
            NotifyDocumentModified();
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante l'applicazione sottolineato: {ex.Message}");
        }
    }

    private void FontColorButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            var colorDialog = new System.Windows.Forms.ColorDialog();

            if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var color = colorDialog.Color;
                // Word usa RGB in formato inverso: BGR
                var rgbValue = (color.B << 16) | (color.G << 8) | color.R;

                var selection = _wordDocument.Application.Selection;
                selection.Font.Color = rgbValue;
                NotifyDocumentModified();
            }
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante il cambio colore: {ex.Message}");
        }
    }

    #endregion

    #region Allineamento Handlers

    private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyAlignment(WdAlignParagraphLeft);
    }

    private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyAlignment(WdAlignParagraphCenter);
    }

    private void AlignRightButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyAlignment(WdAlignParagraphRight);
    }

    private void ApplyAlignment(int alignment)
    {
        if (_wordDocument == null) return;

        try
        {
            var selection = _wordDocument.Application.Selection;
            selection.ParagraphFormat.Alignment = alignment;
            NotifyDocumentModified();

            // Aggiorna stato pulsanti
            AlignLeftButton.IsChecked = alignment == WdAlignParagraphLeft;
            AlignCenterButton.IsChecked = alignment == WdAlignParagraphCenter;
            AlignRightButton.IsChecked = alignment == WdAlignParagraphRight;
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante l'allineamento: {ex.Message}");
        }
    }

    #endregion

    #region Inserimenti Handlers

    private void InsertLinkButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            var selection = _wordDocument.Application.Selection;
            var selectedText = selection.Text?.ToString()?.Trim() ?? string.Empty;

            var dialog = new InsertLinkDialog(selectedText);
            if (dialog.ShowDialog() == true)
            {
                var hyperlinks = _wordDocument.Hyperlinks;
                hyperlinks.Add(
                    Anchor: selection.Range,
                    Address: dialog.Url,
                    TextToDisplay: dialog.DisplayText);

                NotifyDocumentModified();
            }
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante l'inserimento link: {ex.Message}");
        }
    }

    private void InsertImageButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Seleziona immagine",
                Filter = "Immagini (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var selection = _wordDocument.Application.Selection;
                var inlineShapes = selection.InlineShapes;
                inlineShapes.AddPicture(
                    FileName: openFileDialog.FileName,
                    LinkToFile: false,
                    SaveWithDocument: true);

                NotifyDocumentModified();
            }
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante l'inserimento immagine: {ex.Message}");
        }
    }

    #endregion

    #region Modifica Handlers

    private void UndoButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            _wordDocument.Undo();
            NotifyDocumentModified();
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante undo: {ex.Message}");
        }
    }

    private void RedoButton_Click(object sender, RoutedEventArgs e)
    {
        if (_wordDocument == null) return;

        try
        {
            _wordDocument.Redo();
            NotifyDocumentModified();
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante redo: {ex.Message}");
        }
    }

    #endregion

    #region Zoom Handlers

    private void ZoomInButton_Click(object sender, RoutedEventArgs e)
    {
        AdjustZoom(10);
    }

    private void ZoomOutButton_Click(object sender, RoutedEventArgs e)
    {
        AdjustZoom(-10);
    }

    private void AdjustZoom(int delta)
    {
        if (_wordApplication == null) return;

        try
        {
            var currentZoom = (int)_wordApplication.ActiveWindow.View.Zoom.Percentage;
            var newZoom = Math.Max(10, Math.Min(500, currentZoom + delta));
            _wordApplication.ActiveWindow.View.Zoom.Percentage = newZoom;
            ZoomTextBlock.Text = $"{newZoom}%";
        }
        catch (COMException ex)
        {
            _logger.LogWarning($"Errore durante il cambio zoom: {ex.Message}");
        }
    }

    #endregion
}

/// <summary>
/// Dialog semplice per inserimento link.
/// </summary>
internal class InsertLinkDialog : Window
{
    private readonly WpfTextBox _urlTextBox;
    private readonly WpfTextBox _displayTextBox;

    public string Url => _urlTextBox.Text;
    public string DisplayText => _displayTextBox.Text;

    public InsertLinkDialog(string selectedText)
    {
        Title = "Inserisci collegamento";
        Width = 400;
        Height = 180;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;

        var grid = new Grid { Margin = new Thickness(10) };
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var urlLabel = new WpfLabel { Content = "URL:" };
        Grid.SetRow(urlLabel, 0);

        _urlTextBox = new WpfTextBox { Margin = new Thickness(0, 5, 0, 0), Text = "https://" };
        Grid.SetRow(_urlTextBox, 1);

        var displayLabel = new WpfLabel { Content = "Testo da visualizzare:", Margin = new Thickness(0, 10, 0, 0) };
        Grid.SetRow(displayLabel, 2);

        _displayTextBox = new WpfTextBox { Margin = new Thickness(0, 5, 0, 0), Text = selectedText };
        Grid.SetRow(_displayTextBox, 3);

        var buttonPanel = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Margin = new Thickness(0, 10, 0, 0)
        };
        Grid.SetRow(buttonPanel, 5);

        var okButton = new System.Windows.Controls.Button
        {
            Content = "OK",
            Width = 80,
            Height = 28,
            Margin = new Thickness(0, 0, 10, 0),
            IsDefault = true
        };
        okButton.Click += (s, e) => { DialogResult = true; Close(); };

        var cancelButton = new System.Windows.Controls.Button
        {
            Content = "Annulla",
            Width = 80,
            Height = 28,
            IsCancel = true
        };

        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        grid.Children.Add(urlLabel);
        grid.Children.Add(_urlTextBox);
        grid.Children.Add(displayLabel);
        grid.Children.Add(_displayTextBox);
        grid.Children.Add(buttonPanel);

        Content = grid;
    }
}
