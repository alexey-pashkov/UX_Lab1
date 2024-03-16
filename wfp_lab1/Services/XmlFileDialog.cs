using Microsoft.Win32;

public enum XmlFileDialogMode
{
    OpenFile,
    SaveFile
}

class XmlFileDialogService
{
    private string filePath;

    public string FilePath { get; private set; }

    public XmlFileDialogService()
    {
        FilePath = String.Empty;
    }

    public bool Show(XmlFileDialogMode mode)
    {
        FileDialog dialog = mode == XmlFileDialogMode.OpenFile ? dialog = new OpenFileDialog() : dialog = new SaveFileDialog();

        string errorMessage = mode == XmlFileDialogMode.OpenFile ? "Невозможно открыть файл!" : "Невозможно сохранить файл!";

        dialog.DefaultExt = ".xml";
        dialog.Filter = "XML Files|*.xml";

        if (dialog.ShowDialog() == true)
        {
            FilePath = dialog.FileName;
            return true;
        }
        return false;
    }

}