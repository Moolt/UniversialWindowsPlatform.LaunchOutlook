using MsgKit;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using Windows.Storage;
using Windows.Storage.AccessCache;
using Windows.Storage.Pickers;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace UniversialWindowsPlatform.LaunchOutlook
{
    public sealed partial class MainPage : Page
    {
        private readonly List<string> _fileTokens = new List<string>();

        public MainPage()
        {
            this.InitializeComponent();
        }

        private async void OnOpenOutlook(object sender, RoutedEventArgs e)
        {
            var mailSender = new Sender(string.Empty, string.Empty, senderIsCreator: true);
            var mail = new Email(mailSender, Subject.Text, true);
            mail.Recipients.AddTo(To.Text);
            mail.BodyText = Body.Text;
            var local = ApplicationData.Current.LocalCacheFolder.Path;
            var filepath = Path.Combine(local, "foo.msg");

            foreach(var attachment in Attachments)
            {
                mail.Attachments.Add(attachment.Path);
            }

            mail.Save(filepath);
            mail.Dispose();
            OpenOutlook(filepath);
        }

        private async void OpenOutlook(string filepath)
        {
            var msgFile = await StorageFile.GetFileFromPathAsync(filepath);
            await Windows.System.Launcher.LaunchFileAsync(msgFile);
        }

        private async void OnAddAttachment(object sender, RoutedEventArgs e)
        {
            var openPicker = new FileOpenPicker();
            openPicker.ViewMode = PickerViewMode.Thumbnail;
            openPicker.SuggestedStartLocation = PickerLocationId.PicturesLibrary;
            openPicker.FileTypeFilter.Add(".jpg");
            openPicker.FileTypeFilter.Add(".jpeg");
            openPicker.FileTypeFilter.Add(".png");

            var file = await openPicker.PickSingleFileAsync();
            if (file == null)
            {
                return;
            }

            var copiedFile = await file.CopyAsync(ApplicationData.Current.LocalCacheFolder, file.Name, NameCollisionOption.ReplaceExisting);
            Attachments.Add(copiedFile);
        }

        public ObservableCollection<StorageFile> Attachments { get; } = new ObservableCollection<StorageFile>();
    }
}
