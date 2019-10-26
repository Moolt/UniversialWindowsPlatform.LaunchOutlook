using MsgKit;
using OpenMcdf;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace UniversialWindowsPlatform.LaunchOutlook
{
    public sealed partial class MainPage : Page
    {
        private const string MsgFilename = "tmp.msg";
        private readonly string _msgFilepath;

        public MainPage()
        {
            this.InitializeComponent();
            var local = ApplicationData.Current.LocalCacheFolder.Path;
            _msgFilepath = Path.Combine(local, MsgFilename);
        }

        /// <summary>
        /// Assembles an *.msg File, launches Outlook with that file and clears all attachments afterwards.
        /// </summary>
        private async void OnOpenOutlook(object sender, RoutedEventArgs e)
        {
            if (!await AssembleMail())
            {
                return;
            }

            OpenOutlookAsync();
            ClearAttachmentsAsync();
        }

        /// <summary>
        /// Opens a file picker dialog and copies the selected files to the local cache folder.
        /// </summary>
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

            var copiedFile = await file.CopyAsync(
                ApplicationData.Current.LocalCacheFolder,
                file.Name,
                NameCollisionOption.ReplaceExisting);
            Attachments.Add(copiedFile);
        }

        /// <summary>
        /// Assembles an *.msg file and saves it to the local cache folder.
        /// Thanks to Sicos1977 for providing MsgKit: https://github.com/Sicos1977/MsgKit
        /// </summary>
        private async Task<bool> AssembleMail()
        {
            var mailSender = new Sender(string.Empty, string.Empty);
            var mail = new Email(mailSender, Subject.Text, true);
            mail.Recipients.AddTo(To.Text);
            mail.BodyText = Body.Text;
            var local = ApplicationData.Current.LocalCacheFolder.Path;
            var filepath = Path.Combine(local, MsgFilename);

            foreach (var attachment in Attachments)
            {
                mail.Attachments.Add(attachment.Path);
            }

            try
            {
                mail.Save(filepath);
            }
            catch (CFException)
            {
                await new MessageDialog("An Outlook instance is already opened.").ShowAsync();
                return false;
            }

            mail.Dispose();
            return true;
        }

        /// <summary>
        /// Launches the locally saved tmp.msg file.
        /// *.msg files are primarily associated with outlook, so opening an *.msg file will also open Outlook.
        /// </summary>
        private async void OpenOutlookAsync()
        {
            var msgFile = await StorageFile.GetFileFromPathAsync(_msgFilepath);
            var result = await Windows.System.Launcher.LaunchFileAsync(msgFile);

            if (!result)
            {
                await new MessageDialog("Outlook can not be found.").ShowAsync();
            }
        }

        /// <summary>
        /// Clears all cached attachments.
        /// </summary>
        private async void ClearAttachmentsAsync()
        {
            foreach (var cachedFile in Attachments)
            {
                await cachedFile.DeleteAsync();
            }
            Attachments.Clear();
        }

        public ObservableCollection<StorageFile> Attachments { get; } = new ObservableCollection<StorageFile>();
    }
}
