using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

using NPOI.SS.UserModel;

namespace ExcelToEnum
{
    class Converter
    {
        const string _FIELD_ENUM_NAME = "[*N]";
        const string _FIELD_ENUM_VALUE = "[*V]";
        const string _INVALID_REPLACEMENT = "_";
        readonly string[] _INVALID_TEXT = new string[] { " ", "　", "[", "]", "【", "】" };

        const int _START_ROW = 1;

        string m_directory;
        string m_outputDirectory;
        string m_namespaceName;
        OutputEnumData m_outputData;

        public Converter(string directory, string outputDirectory, string settingsDirectory, string namespaceName)
        {
            m_directory = directory;
            m_outputDirectory = outputDirectory;
            m_namespaceName = namespaceName;
            m_outputData = new OutputEnumData(outputDirectory, settingsDirectory);
        }

        public void Convert(string[] originalFiles)
        {
            try
            {
                m_outputData.Clear();
                int fileCount = originalFiles.Length;
                Logger.WriteLine("Start Processing...{0} files", fileCount);
                for (int i = 0; i < fileCount; ++i)
                {
                    Logger.WriteLine("...processing file [{0}/{1}]", i + 1, fileCount);
                    ConvertFile(originalFiles[i]);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLine(ex.Message);
            }
        }

        private void ConvertFile(string path)
        {
            var book = WorkbookFactory.Create(path);
            if (book == null)
            {
                Logger.WriteLine("Failed to load file: " + path);
                return;
            }

            int sheetCount = book.NumberOfSheets;
            for (int i = 0; i < sheetCount; ++i)
            {
                Logger.WriteLine("......processing sheet [{0}/{1}]", i + 1, sheetCount);
                var sheet = book.GetSheetAt(i);
                if (sheet == null)
                {
                    Logger.WriteLine("Failed to get sheet! sheet No = " + i + ", fileName = " + path);
                    continue;
                }

                ConvertSheet(sheet);
            }           
        }

        private void ConvertSheet(ISheet sheet)
        {
            IRow row = sheet.GetRow(sheet.FirstRowNum);

            int nameColumnIndex = Defines._INVALID_VALUE;
            int valueColumnIndex = Defines._INVALID_VALUE;
            if (!FindColumndIndex(row, ref nameColumnIndex, ref valueColumnIndex))
            {
                Logger.WriteLine("Column with mark NOT FOUND!");
                return;
            }

            FillEnumData(sheet, nameColumnIndex, valueColumnIndex);
            m_outputData.Write(sheet.SheetName, m_namespaceName);
        }

        private bool FindColumndIndex(IRow row, ref int nameColumnIndex, ref int valueColumnIndex)
        {
            int columnCount = row.LastCellNum;
            for (int i = 0; i < columnCount; ++i)
            {
                ICell cell = row.GetCell(i);
                if (cell.StringCellValue.Contains(_FIELD_ENUM_NAME))
                {
                    nameColumnIndex = i;
                    continue;
                }

                if (cell.StringCellValue.Contains(_FIELD_ENUM_VALUE))
                {
                    valueColumnIndex = i;
                    continue;
                }
            }

            return (nameColumnIndex != Defines._INVALID_VALUE && valueColumnIndex != Defines._INVALID_VALUE);
        }

        private void FillEnumData(ISheet sheet, int nameColumnIndex, int valueColumnIndex)
        {
            int rowCount = sheet.LastRowNum;
            for (int ii = _START_ROW; ii < rowCount; ++ii)
            {
                ICell nameCell = NPOIUtility.GetCell(sheet, nameColumnIndex, ii);
                ICell valueCell = NPOIUtility.GetCell(sheet, valueColumnIndex, ii);
                if (nameCell == null || valueCell == null)
                {
                    continue;
                }

                string name = nameCell.StringCellValue;
                if (nameCell.CellType == CellType.Blank
                    || valueCell.CellType == CellType.Blank
                    || string.IsNullOrEmpty(name)
                    || string.IsNullOrWhiteSpace(name)
                    )
                {
                    continue;
                }

                // スペースはいらないから何かに差し替える
                int length = _INVALID_TEXT.Length;
                for (int kk = 0; kk < length; ++kk)
                {
                    name = name.Replace(_INVALID_TEXT[kk], _INVALID_REPLACEMENT);
                }

                m_outputData.AddData((int)valueCell.NumericCellValue, name);
            }
        }
    }
}
