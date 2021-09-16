using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

using Microsoft.Win32;
using OfficeTools.Types;
using Prism.Commands;
using Prism.Mvvm;

namespace OfficeTools.ViewModels
{
    public class CompareStylesViewModel : BindableBase
    {


        #region Commands

        private ICommand selectFilesCommand;

        public ICommand SelectFilesCommand => selectFilesCommand ??= new DelegateCommand(OnSelectFilesCommand);

        private void OnSelectFilesCommand()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Title = "Dateien auswählen";
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Alle Dateien (*.*)|*.*|Word-Dokumente (*.docx)|*.docx";
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() != true)
                return;

            Files.Clear();

            foreach (string file in openFileDialog.FileNames)
            {
                Files.Add(new WordFileInfo() {DisplayName = Path.GetFileName(file), Name = file});
            }
        }


        #endregion


        #region Eigenschaften

        private ObservableCollection<WordFileInfo> files;

        public ObservableCollection<WordFileInfo> Files
        {
            get => files ??= new ();
            set => SetProperty(ref files, value);
        }


        #endregion



    }
}
