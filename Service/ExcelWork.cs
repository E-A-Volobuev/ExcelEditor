using ExcelEditor.Model;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ExcelEditor
{
    public class ExcelWork
    {
        /// <summary>
        /// книга откуда считываем данные
        /// </summary>
        public IWorkbook workBook = null;

        /// <summary>
        /// книга,куда записываем данные
        /// </summary>
        private IWorkbook _templateWorkbook = null;

        /// <summary>
        /// словарь собранных данных из ячеек файла excel
        /// </summary>
        private List<Dictionary<int, string>> _rawData { get; set; }

        /// <summary>
        /// считывание из файла ексель объектов бурения
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public List<Info> GetRawData(string filePath)
        {
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            workBook = new XSSFWorkbook(fs);
            var cols = new List<int>();


            for (int i = 0; i < 25; i++)
                cols.Add(i);

            var list = GetRawDataHelper(cols);

            return list;

        }

        private List<Info> GetRawDataHelper(List<int> cols)
        {
            List<Info> list = new List<Info>();

            ExcelHelperModel helper = new ExcelHelperModel() { IndexSheet = 0, IndexRowStart = 1, IndexRowEnd = 0, IndexColumns = cols };

            _rawData = GetDataFromSheet(helper);

            foreach (var rawItem in _rawData)
            {
                var info = CreateInfo(rawItem);

                if (info != null)
                    list.Add(info);
            }

            return list;
        }

        private Info CreateInfo(Dictionary<int, string> rawItem)
        {
            Info info = new Info();
            if (!string.IsNullOrEmpty(rawItem[0]))
            {
                info.CodeSia = rawItem[0];
                info.Croz = rawItem[4];
                info.Kolvo = rawItem[7];
                info.Kpr = rawItem[8];
                info.NameKpr = rawItem[9];
                info.NameMed = rawItem[1];
                info.NameMnn = rawItem[16];
                info.Series = rawItem[5];
                info.Srgodn = rawItem[6];
                info.Sum = SumHelper(info.Croz, info.Kolvo);

                return info;
            }
            else
                return null;
        }
        private double SumHelper(string croz, string col)
        {
            try
            {
                double price = Convert.ToDouble(croz);
                double colvo = Convert.ToDouble(col);
                double result = price * colvo;

                return result;
            }
            catch
            {
                return 0;
            }
        }
        /// <summary>
        /// Считывание данных с листа
        /// </summary>
        /// <param name="indexSheet"></param>
        /// <param name="indexRowStart"></param>
        /// <param name="indexColumns"></param>
        /// <param name="indexRowEnd"></param>
        /// <returns></returns>
        public List<Dictionary<int, string>> GetDataFromSheet(ExcelHelperModel helper)
        {
            var listData = new List<Dictionary<int, string>>();

            if (helper.IndexSheet == -1) return listData;

            var sheet = workBook.GetSheetAt(helper.IndexSheet);
            var rows = sheet.GetRowEnumerator();

            rows.Reset();

            while (rows.MoveNext())
            {
                try
                {
                    var row = (IRow)rows.Current;

                    GetDataFormSheetHelper(row, helper, listData);

                    //если указан индекс последний строки, проверяем и выходим из цикла
                    if (row != null && helper.IndexRowEnd > 0 && row.RowNum == helper.IndexRowEnd)
                    {
                        break;
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

            }

            return listData;
        }

        private void GetDataFormSheetHelper(IRow row, ExcelHelperModel helper, List<Dictionary<int, string>> listData)
        {
            if ((row != null) && (row.RowNum >= helper.IndexRowStart))
            {
                var dicRow = new Dictionary<int, string>();

                SwitchCellType(row, helper, dicRow);

                listData.Add(dicRow);
            }
        }

        private void SwitchCellType(IRow row, ExcelHelperModel helper, Dictionary<int, string> dicRow)
        {
            foreach (int indexCol in helper.IndexColumns)
            {
                var cell = row.GetCell(indexCol);
                var cellValue = "";

                if (cell != null)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                DateTime date = cell.DateCellValue;
                                cellValue = date.ToString().Substring(0, 10);
                            }
                            else
                                cellValue = cell.NumericCellValue.ToString();
                            break;
                        case CellType.String:
                            cellValue = cell.StringCellValue;
                            break;
                        case CellType.Formula:
                            if (cell.CachedFormulaResultType == CellType.String)
                            {
                                cellValue = cell.StringCellValue;
                            }
                            else if (cell.CachedFormulaResultType == CellType.Numeric)
                            {
                                cellValue = cell.NumericCellValue.ToString();
                            }
                            break;
                    }
                }

                dicRow.Add(indexCol, cellValue);
            }
        }


        public void WriteBook(string pathSource, List<Info> list, string pathResult, string fileName)
        {
            using (FileStream fs = new FileStream(pathSource, FileMode.Open, FileAccess.Read))
            {
                _templateWorkbook = new XSSFWorkbook(fs);
            }



            Pharmacy(_templateWorkbook, list);
            Storage(_templateWorkbook, list);

            using (FileStream fs = new FileStream(pathSource, FileMode.Create, FileAccess.Write))
            {
                _templateWorkbook.Write(fs);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            CreateExcelResult(pathSource, pathResult, fileName);

        }

        private void Pharmacy(IWorkbook templateWorkbook, List<Info> list)
        {
            string firstSheetName = "Аптеки";
            ISheet sheetFirst = templateWorkbook.GetSheet(firstSheetName) ?? templateWorkbook.CreateSheet(firstSheetName);

            var listWithoutStorage = list.Where(x => x.NameKpr.Contains("Самарафармация") == false).ToList();

            if (listWithoutStorage.Count() > 0)
            {
                WriteHelper(listWithoutStorage, sheetFirst);
                AllSumStorageHelper(templateWorkbook, sheetFirst, listWithoutStorage);
            }
        }

        private void Storage(IWorkbook templateWorkbook, List<Info> list)
        {
            string secondSheetName = "Склад";
            ISheet sheetSecond = templateWorkbook.GetSheet(secondSheetName) ?? templateWorkbook.CreateSheet(secondSheetName);

            var storageList = list.Where(x => x.NameKpr.Contains("Самарафармация")).ToList();

            if (storageList.Count() > 0)
            {
                WriteHelper(storageList, sheetSecond);
                AllSumStorageHelper(templateWorkbook, sheetSecond, storageList);
            }
        }
        private void AllSumStorageHelper(IWorkbook templateWorkbook, ISheet sheetSecond, List<Info> list)
        {
            IRow dataRow = sheetSecond.GetRow(list.Count() + 4) ?? sheetSecond.CreateRow(list.Count() + 4);

            var styleByText = GetStyleForCell(TypyCell.TEXT, true);
            var styleBySum = GetStyleForCell(TypyCell.NUMERIC, true);

            for (int i = 0; i < 10; i++)
            {
                ICell cell = dataRow.GetCell(i) ?? dataRow.CreateCell(i);
                if (i == 1)
                {
                    cell.SetCellValue("ИТОГО:");
                    cell.CellStyle = styleByText;
                }
                else if (i == 9)
                    SetterSum(cell, styleBySum, list);
                else
                    cell.CellStyle = styleByText;
            }
        }

        private void SetterSum(ICell cell, XSSFCellStyle style, List<Info> list)
        {
            double allSum = list.Select(x => x.Sum).Sum();

            //var nfi = new NumberFormatInfo();
            //nfi.NumberGroupSeparator = " ";
            //nfi.NumberDecimalSeparator = ",";

            //cell.SetCellValue(allSum.ToString("N2", nfi));
            cell.CellStyle = style;
            cell.SetCellValue(allSum);
        }

        private void WriteHelper(List<Info> list, ISheet sheet)
        {
            var orderlyList = list.OrderBy(x => x.NameMnn).ToList();
            var styleByNumeric = GetStyleForCell(TypyCell.NUMERIC,false);
            var styleByText = GetStyleForCell(TypyCell.TEXT, false);
            var styleByDateTime = GetStyleForCell(TypyCell.DATETIME, false);

            for (int i = 4; i < orderlyList.Count() + 4; i++)
            {
                IRow dataRow = sheet.GetRow(i) ?? sheet.CreateRow(i);

                ICell cellByCodeSia = dataRow.GetCell(0) ?? dataRow.CreateCell(0);
                SelectTypeDoubleCellValue(cellByCodeSia, orderlyList[i - 4].CodeSia, true, styleByText);

                ICell cellByKpr = dataRow.GetCell(3) ?? dataRow.CreateCell(3);
                SelectTypeDoubleCellValue(cellByKpr, orderlyList[i - 4].Kpr, true, styleByText);

                ICell cellByCroz = dataRow.GetCell(8) ?? dataRow.CreateCell(8);
                SelectTypeDoubleCellValue(cellByCroz, orderlyList[i - 4].Croz, true, styleByNumeric);

                ICell cellBySum = dataRow.GetCell(9) ?? dataRow.CreateCell(9);
                SelectTypeDoubleCellValue(cellBySum, orderlyList[i - 4].Sum.ToString(), true, styleByNumeric);

                ICell cellByNameMnn = dataRow.GetCell(1) ?? dataRow.CreateCell(1);
                SelectTypeDoubleCellValue(cellByNameMnn, orderlyList[i - 4].NameMnn, false, styleByText);

                ICell cellByNameMed = dataRow.GetCell(2) ?? dataRow.CreateCell(2);
                SelectTypeDoubleCellValue(cellByNameMed, orderlyList[i - 4].NameMed, false, styleByText);

                ICell cellByNameKpr = dataRow.GetCell(4) ?? dataRow.CreateCell(4);
                SelectTypeDoubleCellValue(cellByNameKpr, orderlyList[i - 4].NameKpr, false, styleByText);

                ICell cellBySeries = dataRow.GetCell(5) ?? dataRow.CreateCell(5);
                SelectTypeDoubleCellValue(cellBySeries, orderlyList[i - 4].Series, false, styleByText);

                ICell cellByKolvo = dataRow.GetCell(7) ?? dataRow.CreateCell(7);
                SelectTypeDoubleCellValue(cellByKolvo, orderlyList[i - 4].Kolvo, true, styleByText);

                ICell cellBySrgodn = dataRow.GetCell(6) ?? dataRow.CreateCell(6);
                if (!string.IsNullOrEmpty(orderlyList[i - 4].Srgodn))
                    SelectTypeDoubleCellValue(cellBySrgodn, orderlyList[i - 4].Srgodn, false, styleByDateTime);

            }

        }

        private void SelectTypeDoubleCellValue(ICell cell, string value, bool flag, XSSFCellStyle style)
        {
            cell.CellStyle = style;

            if (flag)
            {
                double number = Convert.ToDouble(value);
                cell.SetCellValue(number);
            }
            else
                cell.SetCellValue(value);

        }

        private XSSFCellStyle GetStyleForCell(TypyCell type, bool isBold)
        {
            XSSFFont myFont = (XSSFFont)_templateWorkbook.CreateFont();
            myFont.FontHeightInPoints = 9;
            myFont.FontName = "Calibri";
            if (isBold)
                myFont.IsBold = true;
            else
                myFont.IsBold = false;

            ///рамка ячейки
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)_templateWorkbook.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            borderedCellStyle.BorderLeft = BorderStyle.Thin;
            borderedCellStyle.BorderTop = BorderStyle.Thin;
            borderedCellStyle.BorderRight = BorderStyle.Thin;
            borderedCellStyle.BorderBottom = BorderStyle.Thin;

            if (type == TypyCell.NUMERIC)
                borderedCellStyle.DataFormat = _templateWorkbook.CreateDataFormat().GetFormat("#,##0.00");
            if (type == TypyCell.DATETIME)
                borderedCellStyle.DataFormat = _templateWorkbook.CreateDataFormat().GetFormat("yyyyMMdd");

            return borderedCellStyle;
        }
        private void CreateExcelResult(string path, string resultPath, string fileName)
        {
            string newPath = resultPath + fileName;
            File.Copy(path, newPath, true);
        }

        private enum TypyCell
        {
            NUMERIC,
            TEXT,
            DATETIME
        }
    }
}
