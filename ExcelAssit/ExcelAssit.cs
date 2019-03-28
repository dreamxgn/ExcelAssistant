using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.OPC;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAssit
{
    public static class ExcelAssit
    {
        public static bool WriteExcel(string path, string sheetName,List<Object> datas)
        {
            if (datas==null || datas.Count <= 0)
            {
                return false;
            }

            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet(sheetName);
            ExcelAssit.CreateTitle(sheet1, datas[0]);

            int rowIndex = 1;
            datas.ForEach(a => {
                ExcelAssit.WriteRow(a, workbook, sheet1, rowIndex);
                rowIndex++;
            });
            return WriteToDisk(path,workbook);
        }

        private static void WriteRow(Object item, IWorkbook book,ISheet sheet,int rowIndex)
        {
            Type type = item.GetType();
            IRow row = sheet.CreateRow(rowIndex);

            int cellIndex = 0;
            type.GetProperties().ToList().ForEach(a => {
                AssitCellAttribute attr = ExcelAssit.GetCellAttribute(a);
                if (attr == null)
                {
                    return;
                }

                ICell cell= ExcelAssit.CreateCell(attr, row, cellIndex);
                cellIndex++;

                if (attr.CellType == CellType.Image)
                {
                    if (a.GetValue(item) == null)
                    {
                        return;
                    }

                    byte[] bytes = null;

                    if (a.PropertyType == typeof(string))
                    {
                        bytes = System.IO.File.ReadAllBytes(a.GetValue(item).ToString());
                        
                    }

                    if (a.PropertyType == typeof(byte[]))
                    {
                        bytes = (byte[])a.GetValue(item);
                    }


                    int pictureIdx = book.AddPicture(bytes, PictureType.JPEG);
                    HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                    HSSFClientAnchor anchor = new HSSFClientAnchor(70, 10, 0, 0, cellIndex-1, rowIndex, cellIndex, rowIndex + 1);
                    HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                    return;
                }

                if (attr.CellType == CellType.Int)
                {
                    if (a.GetValue(item) == null)
                    {
                        return;
                    }
                    cell.SetCellValue(int.Parse(a.GetValue(item).ToString()));
                    return;
                }

                if (attr.CellType == CellType.Float)
                {
                    if (a.GetValue(item) == null)
                    {
                        return;
                    }
                    cell.SetCellValue(float.Parse(a.GetValue(item).ToString()));
                    return;
                }

                if (attr.CellType == CellType.DateTime)
                {
                    if (a.GetValue(item) == null)
                    {
                        return;
                    }
                    if(a.PropertyType==typeof(string))
                    {
                        cell.SetCellValue(int.Parse(a.GetValue(item).ToString()));
                    }

                    if (a.PropertyType == typeof(DateTime))
                    {
                        cell.SetCellValue(DateTime.Parse(a.GetValue(item).ToString()).ToString("yyyy-MM-dd hh:mm:ss"));
                    }
                }

                if (attr.CellType == CellType.String)
                {
                    if (a.GetValue(item) == null)
                    {
                        return;
                    }
                    cell.SetCellValue(a.GetValue(item).ToString());
                    return;
                }
            });
        }


        private static bool WriteToDisk(string path, IWorkbook book)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (FileStream fs = new FileStream(path, FileMode.Create))
            {
                book.Write(fs);
                return true;
            }
        }

        private static IWorkbook NPOIOpenExcel(string filename)
        {
            IWorkbook myworkBook;
            Stream excelStream = OpenResource(filename);
            if (POIFSFileSystem.HasPOIFSHeader(excelStream))
                return new HSSFWorkbook(excelStream);
            if (POIXMLDocument.HasOOXMLHeader(excelStream))
            {
                return new XSSFWorkbook(OPCPackage.Open(excelStream));
            }
            if (filename.EndsWith(".xlsx"))
            {
                return new XSSFWorkbook(excelStream);
            }
            if (filename.EndsWith(".xls"))
            {
                new HSSFWorkbook(excelStream);
            }
            throw new Exception("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
        }

        private static Stream OpenResource(string filename)
        {
            FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read);
            return fs;
        }



        public static bool AppendExcel(string path, string sheetName, List<Object> datas)
        {
            if (datas == null || datas.Count <= 0)
            {
                return false;
            }

            IWorkbook workbook = NPOIOpenExcel(path);

            ISheet sheet= workbook.GetSheet(sheetName);
            int rowIndex = sheet.LastRowNum + 1;

            datas.ForEach(a => {
                ExcelAssit.WriteRow(a, workbook, sheet, rowIndex);
                rowIndex++;
            });
            return WriteToDisk(path, workbook);
        }

        public static List<T> ReadExcel<T>(string path, string sheetName)
        {
            Type type = typeof(T);
            List<T> list = new List<T>();

            IWorkbook workbook = NPOIOpenExcel(path);
            ISheet sheet= workbook.GetSheet(sheetName);

            IRow row = sheet.GetRow(0);
            Dictionary<string, PropertyInfo> propertys = new Dictionary<string, PropertyInfo>();

            var propertiess = type.GetProperties().ToList();

            int cellIndex = 0;
            row.Cells.ForEach(a => {
                propertiess.ForEach(a1 => {
                    AssitCellAttribute attr = a1.GetCustomAttribute<AssitCellAttribute>();
                    if (attr == null) { return; }

                    if (attr.CellTitle == a.StringCellValue.Trim())
                    {
                        propertys.Add(cellIndex.ToString(), a1);
                    }
                    cellIndex++;
                });
            });

            for (int i = 1; i < sheet.LastRowNum; i++)
            {
                IRow irow = sheet.GetRow(i);

                var createfn = typeof(T).GetConstructors()[0];
                Object t = createfn.Invoke(null);

                for (int j = 0; j < irow.Cells.Count; j++)
                {
                    PropertyInfo pi = propertys[j.ToString()];
                    if (pi == null)
                    {
                        continue;
                    }

                    AssitCellAttribute attr = GetCellAttribute(pi);
                    if (attr == null)
                    {
                        continue;
                    }

                    if (attr.CellType == CellType.Image)
                    {
                        continue;
                    }

                    ICell a = irow.GetCell(j);

                    if (attr.CellType == CellType.String)
                    {
                        if (a != null && a.StringCellValue != null)
                        {
                            pi.SetValue(t, a.StringCellValue);
                        }



                    }

                    if (attr.CellType == CellType.Int)
                    {
                        if (a != null && a.StringCellValue != null)
                        {
                            pi.SetValue(t, int.Parse(a.StringCellValue));
                        }
                    }

                    if (attr.CellType == CellType.Float)
                    {
                        if (a != null && a.StringCellValue != null)
                        {
                            pi.SetValue(t, float.Parse(a.StringCellValue));
                        }
                    }

                    if (attr.CellType == CellType.DateTime)
                    {
                        if (a != null && a.StringCellValue != null)
                        {
                            if (pi.PropertyType == typeof(string))
                            {
                                pi.SetValue(t, a.StringCellValue);
                            }

                            if (pi.PropertyType == typeof(DateTime))
                            {
                                pi.SetValue(t, DateTime.Parse(a.StringCellValue));
                            }
                        }
                    }
                }

                



                list.Add((T)t);
            }




            return null;
        }

        private static void CreateTitle(ISheet sheet,Object item)
        {
            Type type= item.GetType();
            List<AssitCellAttribute> titles = ExcelAssit.GetCellAttributes(type);

            IRow row = sheet.CreateRow(0);
            int cellIndex = 0;
            titles.ForEach(a => {
                ExcelAssit.CreateCellTitle(a, row, cellIndex);
                cellIndex++;
            });
        }

        private static List<AssitCellAttribute> GetCellAttributes(Type type)
        {
            List<AssitCellAttribute> list = new List<AssitCellAttribute>();
            type.GetProperties().ToList().ForEach(a => {

                AssitCellAttribute[] attr = (AssitCellAttribute[])a.GetCustomAttributes(typeof(AssitCellAttribute), false);
                if (attr != null && attr.Length == 1)
                {
                    list.Add(attr[0]);
                }

            });
            return list;
        }

        private static AssitCellAttribute GetCellAttribute(PropertyInfo p)
        {
            AssitCellAttribute[] attr = (AssitCellAttribute[])p.GetCustomAttributes(typeof(AssitCellAttribute), false);
            if (attr != null && attr.Length == 1)
            {
                return attr[0];
            }
            return null;
        }

        private static ICell CreateCellTitle(AssitCellAttribute attr,IRow row,int index)
        {
            ICell cell = ExcelAssit.CreateCell(attr, row, index);
            cell.SetCellValue(attr.CellTitle);
            return cell;
        }

        private static ICell CreateCell(AssitCellAttribute attr, IRow row, int index)
        {
            ICell cell = row.CreateCell(index);

            if (attr.CellType == CellType.String)
            {
                cell.SetCellType(NPOI.SS.UserModel.CellType.String);
            }

            if (attr.CellType == CellType.Int)
            {
                cell.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
            }

            if (attr.CellType == CellType.Float)
            {
                cell.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
            }

            if (attr.CellType == CellType.DateTime)
            {
                cell.SetCellType(NPOI.SS.UserModel.CellType.String);
            }

            if (attr.CellType == CellType.Image)
            {
                cell.SetCellType(NPOI.SS.UserModel.CellType.String);
            }
            return cell;
        }
    }
}
