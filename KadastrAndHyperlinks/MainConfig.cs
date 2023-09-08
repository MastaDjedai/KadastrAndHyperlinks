using CommunityToolkit.Mvvm.Input;
using DevExpress.Mvvm;
using DevExpress.Mvvm.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Windows;
using System.IO;
using System.Drawing;

namespace KadastrAndHyperlinks
{
    public class MainConfig : ViewModelBase
    {
        
        private string _pathToXlsx;
        public string pathToXlsx
        {
            get
            {
                return _pathToXlsx;
            }
            set
            {
                _pathToXlsx = value;
                RaisePropertyChanged(()=> pathToXlsx);
            }
        }

        private string _pathToFolder;
        public string pathToFolder
        {
            get
            {
                return _pathToFolder;
            }
            set
            {
                _pathToFolder = value;
                RaisePropertyChanged(()=> pathToFolder);
            }
        }
        private void SelectXlsxFile()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.ShowDialog();
            pathToXlsx = openFile.FileName;
        }
        private void GetPathToMainFolder()
        {
            var folderDialog = new FolderBrowserDialog();
            folderDialog.ShowDialog();
            pathToFolder = folderDialog.SelectedPath;

        }
        public ICommand ChoseXlsxfile
        {
            get
            {
                return new RelayCommand(SelectXlsxFile);
            }
        }
        public ICommand ChoseFolderPath
        {
            get
            {
                return new RelayCommand(GetPathToMainFolder);
            }
        }

        public void WorkWithFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string xlsxPath = pathToXlsx;
            string rootFolder = pathToFolder;
            var package = new ExcelPackage(xlsxPath);
            var worksheet = package.Workbook.Worksheets[0];
            int column = worksheet.Dimension.Rows;
            List<string>fullFolderPath = new List<string>();
            for (int i = 1; i <= column; i++)
            {
                string temporaryStorage = worksheet.Cells[i, 1].Text.Replace(":", "");
                fullFolderPath.Add(Path.Combine(rootFolder, temporaryStorage));
            }

            string[] folders = Directory.GetDirectories(rootFolder);
            for(int i = 1; i <= column; i++)
            {
                string temporaryStorage = worksheet.Cells[i, 1].Text.Replace(":","");
                bool folderExist = folders.Any(folder => Path.GetFileName(folder) == temporaryStorage);
                if (folderExist)
                {
                    worksheet.Cells[i, 1].Hyperlink = new ExcelHyperLink(fullFolderPath[i-1]);
                    worksheet.Cells[i, 1].Style.Font.UnderLine = true;
                    worksheet.Cells[i, 1].Style.Font.Color.SetColor(Color.Blue);
                }
            }

            package.Save();
        }

        public ICommand GetLinks
        {
            get
            {
                return new RelayCommand(WorkWithFile);
            }
        }

    }
}
