using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Reflection;

namespace ExcelImportImage
{
    /*
     * Author：zhoukaikai
     * 注意：本项目中使用的NPOI版本为 V2.4.1.0；需要在NuGet上添加NPOI包
     * 如果发现图片位置错误 或 图片不显示，请先确认NPOI版本是否与本项目一致
     */
    class Program
    {
        static void Main(string[] args)
        {
            string excelPath = @"D:\利用NPOI向Excel指定位置中加入图片\excel.xlsx";
            //string excelPath = @"D:\利用NPOI向Excel指定位置中加入图片\excelXLS.xls";
            string imgPath = @"D:\利用NPOI向Excel指定位置中加入图片\image.png";
            string fileExtensionName = Path.GetExtension(excelPath);
            if (fileExtensionName.ToLower() == ".xlsx")
            {
                InsertImageToXLSXExcel(excelPath, imgPath);
            }
            if (fileExtensionName.ToLower() == ".xls")
            {
                InsertImageToXLSExcel(excelPath, imgPath);
            }
        }

        /// <summary>
        /// .xlsx后缀的Excel文件添加图片
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="imgPath"></param>
        public static void InsertImageToXLSXExcel(string excelPath, string imgPath)
        {
            try
            {
                using (FileStream fs = new FileStream(excelPath, FileMode.Open))//获取指定Excel文件流
                {
                    //创建工作簿
                    XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
                    //获取第一个工作表（下标从0起）
                    XSSFSheet sheet = (XSSFSheet)xssfworkbook.GetSheet(xssfworkbook.GetSheetName(0));

                    //获取指定图片的字节流
                    byte[] bytes = System.IO.File.ReadAllBytes(imgPath);
                    //将图片添加到工作簿中，返回值为该图片在工作表中的索引（从0开始）
                    //图片所在工作簿索引理解：如果原Excel中没有图片，那执行下面的语句后，该图片为Excel中的第1张图片，其索引为0；
                    //同理，如果原Excel中已经有1张图片，执行下面的语句后，该图片为Excel中的第2张图片，其索引为1；
                    int pictureIdx = xssfworkbook.AddPicture(bytes, PictureType.JPEG);

                    //创建画布
                    XSSFDrawing patriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();
                    //设置图片坐标与大小
                    //函数原型：XSSFClientAnchor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2)；
                    //坐标(col1,row1)表示图片左上角所在单元格的位置，均从0开始，比如(5,2)表示(第五列，第三行),即F3；注意：图片左上角坐标与(col1,row1)单元格左上角坐标重合
                    //坐标(col2,row2)表示图片右下角所在单元格的位置，均从0开始，比如(10,3)表示(第十一列，第四行),即K4；注意：图片右下角坐标与(col2,row2)单元格左上角坐标重合
                    //坐标(dx1,dy1)表示图片左上角在单元格(col1,row1)基础上的偏移量(往右下方偏移)；(dx1，dy1)的最大值为(1023, 255),为一个单元格的大小
                    //坐标(dx2,dy2)表示图片右下角在单元格(col2,row2)基础上的偏移量(往右下方偏移)；(dx2,dy2)的最大值为(1023, 255),为一个单元格的大小
                    //注意：目前测试发现，对于.xlsx后缀的Excel文件，偏移量设置(dx1,dy1)(dx2,dy2)无效；只会对.xls生效
                    XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 5, 2, 10, 3);
                    //正式在指定位置插入图片
                    XSSFPicture pict = (XSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

                    //创建一个新的Excel文件流，可以和原文件名不一样，
                    //如果不一样，则会创建一个新的Excel文件；如果一样，则会覆盖原文件
                    FileStream file = new FileStream(excelPath, FileMode.Create);
                    //将已插入图片的Excel流写入新创建的Excel中
                    xssfworkbook.Write(file);
                    
                    //关闭工作簿
                    xssfworkbook.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// .xls后缀的Excel文件添加图片
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="imgPath"></param>
        public static void InsertImageToXLSExcel(string excelPath, string imgPath)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(imgPath);
            try
            {
                using (FileStream fs = new FileStream(excelPath, FileMode.Open))
                {
                    HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);
                    HSSFSheet sheet = (HSSFSheet)hssfworkbook.GetSheet(hssfworkbook.GetSheetName(0));

                    int pictureIdx = hssfworkbook.AddPicture(bytes, PictureType.JPEG);

                    HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                    HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255, 5, 2, 10, 3);//(255, 125, 1023, 150, 5, 2, 10, 3);//(0, 0, 0, 0, 5, 2, 10, 3);
                    HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

                    FileStream file = new FileStream(excelPath, FileMode.Create);
                    hssfworkbook.Write(file);
                    hssfworkbook.Close();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
