using Android.Content;
using Android.OS;
using Xamarin.Forms;
using ProdLogProg.Droid;
using System.IO;
using Android.Widget;

[assembly: Dependency(typeof(FileService))]
namespace ProdLogProg.Droid
{
    public class FileService : IFileService
    {
        public void SaveFile(string fileName, byte[] fileData)
        {
            string downloadsPath = Environment.GetExternalStoragePublicDirectory(Environment.DirectoryDownloads).AbsolutePath;

            string filePath = Path.Combine(downloadsPath, fileName);

            File.WriteAllBytes(filePath, fileData);

            var context = Android.App.Application.Context;
            Toast.MakeText(context, $"File saved to {filePath}", ToastLength.Long).Show();
        }
    }
}
