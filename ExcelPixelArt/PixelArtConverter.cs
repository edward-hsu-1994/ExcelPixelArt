using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace ExcelPixelArt {
    public class PixelArtConverter {
        /// <summary>
        /// 取得Excel像素藝術中，一個像素的長寬大小
        /// </summary>
        public short PixelSize { get; private set; }

        /// <summary>
        /// 取得清晰度(合併像素)，預設為1，表示最高清晰度
        /// </summary>
        public int Clarity { get; private set; }

        /// <summary>
        /// 初始化Excel像素藝術轉換器實例
        /// </summary>
        /// <param name="pixelSize">像素長寬</param>
        /// <param name="clarity">清晰度</param>
        public PixelArtConverter(short pixelSize = 8, int clarity = 1) {
            if (pixelSize <= 0) throw new ArgumentException($"{nameof(pixelSize)}不應該為0或負數");
            if (clarity <= 0) throw new ArgumentException($"{nameof(clarity)}不應該為0或負數");
            this.PixelSize = pixelSize;
            this.Clarity = clarity;
        }

        private Bitmap BitmapColorConvert(Bitmap image) {
            return ImageConvertor.Convertor1.ConvertTo8bppFormat(image);
        }

        
        private Bitmap Zoom(Bitmap image) {
            //ref : http://stackoverflow.com/questions/23879178/zoom-bitmap-image 
            Bitmap result = new Bitmap(image.Width / Clarity, image.Height / Clarity);
            Graphics g = Graphics.FromImage(result);
            Rectangle srcRect = new Rectangle(0, 0, image.Width, image.Height);
            Rectangle dstRect = new Rectangle(0, 0, result.Width, result.Height);
            g.DrawImage(image, dstRect, srcRect, GraphicsUnit.Pixel);

            return result;
        }

        private XSSFColor[][] BitmapToExcelColorArray(Bitmap image) {
            XSSFColor[][] result = new XSSFColor[image.Width][];
            for (int x = 0; x < image.Width; x++) {
                result[x] = new XSSFColor[image.Height];
                for (int y = 0; y < image.Height; y++) {
                    result[x][y] = new XSSFColor(image.GetPixel(x, y));
                }
            }
            return result;
        }

        public IWorkbook Convert(string filePath) => Convert(new Bitmap(filePath));

        public IWorkbook Convert(Image image) => Convert(new Bitmap(image));

        public IWorkbook Convert(Bitmap image) {
            Bitmap zoomImage = Zoom(BitmapColorConvert(image));//縮放
            var data = BitmapToExcelColorArray(zoomImage);

            IWorkbook result = new XSSFWorkbook();//Excel 2007+
            ISheet sheet = result.CreateSheet("pixelArt");//產生工作表
            double temp = zoomImage.Width * zoomImage.Height;

            //欄寬設定
            for (int col = 0; col < zoomImage.Width; col++) {
                sheet.SetColumnWidth(col, 37 * PixelSize);//1000 = 27px 1px
            }
            for (int row = 0; row < zoomImage.Height; row++) {
                sheet.CreateRow(row);
            }

            double okRow = 0;
            //Parallel.For(0, 2, x => {
                for (int row = 0/*x==0 ? 0 : data[0].Length - 1*/;
                     /*x== 0 ?*/ row < data[0].Length/* / 2 : row >= data[0].Length / 2*/;
                     row++/*row+= x == 0 ? 1 : -1*/) {
                    var Row = sheet.CreateRow(row);
                    Row.Height = (short)(15 * PixelSize);
                    for (int col = 0; col < data.Length; col++) {

                        var Cell = Row.CreateCell(col);

                        XSSFCellStyle style = (XSSFCellStyle)result.CreateCellStyle();
                        style.SetFillForegroundColor(data[col][row]);
                        style.FillPattern = FillPattern.SolidForeground;

                        Cell.CellStyle = style;
                    }
                    okRow++;
                    Console.WriteLine($"執行進度: {Math.Round((okRow / data[0].Length) * 100)}%");
                }
            //});
            return result;
        }
    }
}
