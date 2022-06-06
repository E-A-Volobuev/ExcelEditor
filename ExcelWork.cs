using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelEditor
{
    public class ExcelWork
    {
        public IWorkbook workBook = null;
        /// <summary>
        /// словарь собранных данных из ячеек файла excel
        /// </summary>
        public List<Dictionary<int, string>> _rawData { get; set; }
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
            List<Info> list = new List<Info>();

            for (int i = 0; i < 25; i++)
                cols.Add(i);

            _rawData = GetDataFromSheet(0, 1, cols);

            foreach (var rawItem in _rawData)
            {
                Info info = new Info();
                info.CODE_SIA = rawItem[0];
                info.CROZ = rawItem[4];
                info.KOLVO = rawItem[7];
                info.KPR = rawItem[8];
                info.NAME_KPR = rawItem[9];
                info.NAME_MED = rawItem[1];
                info.NAME_MNN = rawItem[16];
                info.SERIES = rawItem[5];
                info.SRGODN = rawItem[6];

                info.SUM = SumHelper(info.CROZ, info.KOLVO);

                if (info != null)
                    list.Add(info);

            }

            return list;

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
        public List<Dictionary<int, string>> GetDataFromSheet(int indexSheet, int indexRowStart, List<int> indexColumns, int indexRowEnd = 0)
        {
            var listData = new List<Dictionary<int, string>>();

            if (indexSheet == -1) return listData;

            var sheet = workBook.GetSheetAt(indexSheet);
            var rows = sheet.GetRowEnumerator();

            rows.Reset(); // необязательно, но для страховки оставил

            while (rows.MoveNext())
            {
                try
                {
                    var row = (IRow)rows.Current;

                    if (row != null && row.RowNum >= indexRowStart)
                    {
                        var dicRow = new Dictionary<int, string>();

                        foreach (int indexCol in indexColumns)
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
                                        {
                                            cellValue = cell.NumericCellValue.ToString();
                                        }
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

                        listData.Add(dicRow);

                    } //end if

                    //если указан индекс последний строки, проверяем и выходим из цикла
                    if (row != null && indexRowEnd > 0 && row.RowNum == indexRowEnd)
                    {
                        break;
                    }

                } //try
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

            } //while

            return listData;
        }


        public void WriteBook(string pathSource, List<Info> list,string pathResult, string fileName)
        {
            IWorkbook templateWorkbook;
            using (FileStream fs = new FileStream(pathSource, FileMode.Open, FileAccess.Read))
            {
                templateWorkbook = new XSSFWorkbook(fs);
            }

            XSSFFont myFont = (XSSFFont)templateWorkbook.CreateFont();
            myFont.FontHeightInPoints = (short)9;
            myFont.FontName = "Calibre";
            myFont.IsBold = false;

            ///рамка ячейки
            XSSFCellStyle borderedCellStyle = (XSSFCellStyle)templateWorkbook.CreateCellStyle();
            borderedCellStyle.SetFont(myFont);
            borderedCellStyle.BorderLeft = BorderStyle.Thin;
            borderedCellStyle.BorderTop = BorderStyle.Thin;
            borderedCellStyle.BorderRight = BorderStyle.Thin;
            borderedCellStyle.BorderBottom = BorderStyle.Thin;

            string firstSheetName = "Аптека";
            ISheet sheetFirst = templateWorkbook.GetSheet(firstSheetName) ?? templateWorkbook.CreateSheet(firstSheetName);

            string secondSheetName = "Склад";
            ISheet sheetSecond = templateWorkbook.GetSheet(secondSheetName) ?? templateWorkbook.CreateSheet(secondSheetName);

            short doubleFormat = templateWorkbook.CreateDataFormat().GetFormat("#,##0.###");

            var listWithoutStorage = list.Where(x => x.NAME_KPR.Contains("Самарафармация")==false).ToList();

            if (listWithoutStorage.Count() > 0)
            {
                WriteHelper(listWithoutStorage, sheetFirst, borderedCellStyle, doubleFormat);
                AllSumStorageHelper(templateWorkbook,sheetFirst,listWithoutStorage);
            }

            var storageList = list.Where(x => x.NAME_KPR.Contains("Самарафармация")).ToList();

            if (storageList.Count() > 0)
            {
                WriteHelper(storageList, sheetSecond, borderedCellStyle, doubleFormat);
                AllSumStorageHelper(templateWorkbook,sheetSecond, storageList);
            }
            using (FileStream fs = new FileStream(pathSource, FileMode.Create, FileAccess.Write))
            {
                templateWorkbook.Write(fs);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            CreateExcelResult(pathSource,pathResult, fileName);

        }
        private void AllSumStorageHelper(IWorkbook templateWorkbook,ISheet sheetSecond, List<Info> list)
        {
            IRow dataRow = sheetSecond.GetRow(list.Count()+4) ?? sheetSecond.CreateRow(list.Count()+4);

            XSSFFont myFontSum = (XSSFFont)templateWorkbook.CreateFont();
            myFontSum.FontHeightInPoints = (short)9;
            myFontSum.FontName = "Calibre";
            myFontSum.IsBold = true;

            ///рамка ячейки
            XSSFCellStyle newFontCellStyle = (XSSFCellStyle)templateWorkbook.CreateCellStyle();
            newFontCellStyle.SetFont(myFontSum);
            newFontCellStyle.BorderLeft = BorderStyle.Thin;
            newFontCellStyle.BorderTop = BorderStyle.Thin;
            newFontCellStyle.BorderRight = BorderStyle.Thin;
            newFontCellStyle.BorderBottom = BorderStyle.Thin;

            for (int i = 0; i < 10; i++)
            {
                ICell cell = dataRow.GetCell(i) ?? dataRow.CreateCell(i);
                if (i == 1)
                {
                    cell.SetCellValue("ИТОГО:");
                    cell.CellStyle = newFontCellStyle;
                }                           
                else if (i == 9)
                {
                    double allSum = list.Select(x => x.SUM).Sum();

                    var nfi = new NumberFormatInfo();
                    nfi.NumberGroupSeparator = " "; 
                    nfi.NumberDecimalSeparator = ","; 


                    cell.SetCellValue(allSum.ToString("N2", nfi));
                    cell.CellStyle = newFontCellStyle;
                }
                else
                    cell.CellStyle = newFontCellStyle;
            }         
        }

        private void WriteHelper(List<Info> list, ISheet sheet, XSSFCellStyle borderedCellStyle, short doubleFormat)
        {
            var orderlyList = list.OrderBy(x => x.NAME_MNN).ToList();

            for (int i = 4; i < orderlyList.Count()+4; i++)
            {
                IRow dataRow = sheet.GetRow(i) ?? sheet.CreateRow(i);
                ICell cellByCODE_SIA = dataRow.GetCell(0) ?? dataRow.CreateCell(0);
                cellByCODE_SIA.CellStyle = borderedCellStyle;
                cellByCODE_SIA.SetCellValue(orderlyList[i-4].CODE_SIA);

                ICell cellByNAME_MNN = dataRow.GetCell(1) ?? dataRow.CreateCell(1);
                cellByNAME_MNN.CellStyle = borderedCellStyle;
                cellByNAME_MNN.SetCellValue(orderlyList[i-4].NAME_MNN);

                ICell cellByNAME_MED = dataRow.GetCell(2) ?? dataRow.CreateCell(2);
                cellByNAME_MED.CellStyle = borderedCellStyle;
                cellByNAME_MED.SetCellValue(orderlyList[i-4].NAME_MED);

                ICell cellByKPR = dataRow.GetCell(3) ?? dataRow.CreateCell(3);
                cellByKPR.CellStyle = borderedCellStyle;
                cellByKPR.SetCellValue(orderlyList[i-4].KPR);

                ICell cellByNAME_KPR = dataRow.GetCell(4) ?? dataRow.CreateCell(4);
                cellByNAME_KPR.CellStyle = borderedCellStyle;
                cellByNAME_KPR.SetCellValue(orderlyList[i-4].NAME_KPR);

                ICell cellBySERIES = dataRow.GetCell(5) ?? dataRow.CreateCell(5);
                cellBySERIES.CellStyle = borderedCellStyle;
                cellBySERIES.SetCellValue(orderlyList[i-4].SERIES);

                ICell cellBySRGODN = dataRow.GetCell(6) ?? dataRow.CreateCell(6);
                if (!string.IsNullOrEmpty(orderlyList[i-4].SRGODN))
                {
                    cellBySRGODN.CellStyle = borderedCellStyle;
                    cellBySRGODN.SetCellValue(orderlyList[i - 4].SRGODN);
                }

                ICell cellByKOLVO = dataRow.GetCell(7) ?? dataRow.CreateCell(7);
                cellByKOLVO.CellStyle = borderedCellStyle;
                cellByKOLVO.SetCellValue(Convert.ToDouble(orderlyList[i-4].KOLVO));

                ICell cellByCROZ = dataRow.GetCell(8) ?? dataRow.CreateCell(8);
                cellByCROZ.CellStyle = borderedCellStyle;
                cellByCROZ.SetCellValue(Convert.ToDouble(orderlyList[i-4].CROZ));

                ICell cellBySum = dataRow.GetCell(9) ?? dataRow.CreateCell(9);
                cellBySum.CellStyle = borderedCellStyle;
                cellBySum.SetCellValue(Convert.ToDouble(orderlyList[i-4].SUM));
            }


        }

        private void CreateExcelResult(string path, string resultPath, string fileName)
        {
            string newPath = resultPath + fileName;
            File.Copy(path, newPath, true);
        }
    }
}
