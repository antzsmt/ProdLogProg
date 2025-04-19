using Android.OS;
using System.IO;
using Android.Widget;
using Xamarin.Forms;
using ProdLogProg.Droid;

[assembly: Dependency(typeof(FileService))]
namespace ProdLogProg.Droid
{
    public class FileService : IFileService
    {
        // Implementation of SaveFile method
        public void SaveFile(string fileName, byte[] fileData)
        {
            string downloadsPath = Environment.GetExternalStoragePublicDirectory(Environment.DirectoryDownloads).AbsolutePath;
            string filePath = Path.Combine(downloadsPath, fileName);

            File.WriteAllBytes(filePath, fileData);

            var context = Android.App.Application.Context;
            Toast.MakeText(context, $"File saved to {filePath}", ToastLength.Long).Show();
        }

        // Implementation of GetSavedFilePath method
        public string GetSavedFilePath(string fileName)
        {
            string downloadsPath = Environment.GetExternalStoragePublicDirectory(Environment.DirectoryDownloads).AbsolutePath;
            return Path.Combine(downloadsPath, fileName);
        }
    }
}
