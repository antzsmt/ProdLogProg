namespace ProdLogProg
{
    public interface IFileService
    {
        void SaveFile(string fileName, byte[] fileData);
    }
}