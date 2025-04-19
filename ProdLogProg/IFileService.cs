namespace ProdLogProg
{
    public interface IFileService
    {
        // Method to save a file
        void SaveFile(string fileName, byte[] fileData);

        // Method to get the file path of the last saved file
        string GetSavedFilePath(string fileName);
    }
}