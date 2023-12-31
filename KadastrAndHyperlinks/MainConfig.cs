﻿using CommunityToolkit.Mvvm.Input;
using DevExpress.Mvvm;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Animation;

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
                RaisePropertyChanged(() => pathToXlsx);
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
                RaisePropertyChanged(() => pathToFolder);
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

        public void WorkWithFileNew()
        {
            bool emptyFolderPath = !string.IsNullOrEmpty(pathToFolder);
            bool emptyXlsxPath = !string.IsNullOrEmpty(pathToXlsx);
            if (emptyFolderPath&& emptyXlsxPath)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string xlsxPath = pathToXlsx;
                string rootFolder = pathToFolder;
                var package = new ExcelPackage(xlsxPath);
                var worksheet = package.Workbook.Worksheets[0];
                int row = worksheet.Dimension.End.Row;
                int column = worksheet.Dimension.End.Column;
                string[] foldersName = Directory.GetDirectories(rootFolder);

                for (int columns = 1; columns <= column; columns++)
                {
                    for (int rows = 1; rows <= row; rows++)
                    {
                        string temp = worksheet.Cells[rows, columns].Text.Replace(":", "");
                        for (int i = 0; i < foldersName.Length; i++)
                        {
                            if (temp == Path.GetFileName(foldersName[i]))
                            {
                                string fullFolderPath = Path.Combine(rootFolder, foldersName[i]);
                                worksheet.Cells[rows, columns].Hyperlink = new ExcelHyperLink(fullFolderPath);
                                worksheet.Cells[rows, columns].Style.Font.UnderLine = true;
                                worksheet.Cells[rows, columns].Style.Font.Color.SetColor(Color.Blue);
                            }
                        }
                    }
                }
                StartProgressBar();
                package.Save();
            }
            else
            {
                EmptyPathException();
            }
        }

        public void MakeFoldersFromExcel()
        {
            
            bool emptyFolderPath = !string.IsNullOrEmpty(pathToFolder);
            bool emptyXlsxPath = !string.IsNullOrEmpty(pathToXlsx);
            if (emptyFolderPath&&emptyXlsxPath)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string rootFolder = pathToFolder;
                string xlsxPath = pathToXlsx;
                ExcelPackage package = new ExcelPackage(xlsxPath);
                var worksheet = package.Workbook.Worksheets[0];
                string kadastrForm = @"^\d{10}:\d{2}:\d{3}:\d{4}$";
                int lastNonEmptyColumn = worksheet.Dimension.End.Column;
                int lastNonEmptyRow = worksheet.Dimension.End.Row;
                for (int columns = 1; columns <= lastNonEmptyColumn; columns++)
                {
                    for (int rows = 1; rows <= lastNonEmptyRow; rows++)
                    {
                        string temp = worksheet.Cells[rows, columns].Text;
                        if (!string.IsNullOrEmpty(temp))
                        {
                            bool isMatch = Regex.IsMatch(temp, kadastrForm);
                            if(isMatch)
                            {
                                temp = temp.Replace(":", "");
                                string fullPath = Path.Combine(rootFolder, temp);
                                if(!Directory.Exists(fullPath))
                                {
                                    Directory.CreateDirectory(fullPath);
                                }
                            }
                        }
                    }
                }
                StartProgressBar();
            }
            else
            {
                EmptyPathException();
            }
        }
        private Window checkPathWindow;
        public MainConfig()
        {
            
        }
        private void EmptyPathException()
        {
            checkPathWindow = new Views.EmptyFoldersOrXlsxTextBlock();
            checkPathWindow.ShowDialog();
        }
        
        private int _progress;
        public int Progress
        {
            get
            {
                return _progress;
            }
            set
            {
                _progress = value;
                RaisePropertyChanged(() => Progress);
            }
        }
        private async Task ProgressBarAnimationAsync()
        {
            Progress = 0;
            for (int i = 0; i < 100; i++)
            {
                Progress++;
                await Task.Delay(3);
            }
        }
        private async void StartProgressBar()
        {
            await ProgressBarAnimationAsync();
        }
        
        public ICommand GetLinks => new RelayCommand(WorkWithFileNew);
        public ICommand CreateFolder => new RelayCommand(MakeFoldersFromExcel);
    }
}
