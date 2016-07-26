using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

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
        /// 圖片品質，如果true則表示自動轉換圖片為256色
        /// </summary>
        public bool LowQuality { get; private set; }
        
        /// <summary>
        /// 初始化Excel像素藝術轉換器實例
        /// </summary>
        /// <param name="pixelSize">像素長寬</param>
        /// <param name="clarity">清晰度</param>
        public PixelArtConverter(short pixelSize = 8, int clarity = 1, bool lowQuality = false) {
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
            Bitmap result = new Bitmap(image.Width / Clarity, image.Height / Clarity,PixelFormat.Format32bppArgb);
            Graphics g = Graphics.FromImage(result);
            Rectangle srcRect = new Rectangle(0, 0, image.Width, image.Height);
            Rectangle dstRect = new Rectangle(0, 0, result.Width, result.Height);
            g.DrawImage(image, dstRect, srcRect, GraphicsUnit.Pixel);

            return result;
        }

        private Color[][] BitmapToColorArray(Bitmap image) {
            Color[][] result = new Color[image.Width][];
            for (int x = 0; x < image.Width; x++) {
                result[x] = new Color[image.Height];
                for (int y = 0; y < image.Height; y++) {
                    result[x][y] = image.GetPixel(x, y);
                }
            }
            return result;
        }

        public void Convert(string filePath, string savePath) => Convert(new Bitmap(filePath), savePath);

        public void Convert(Image image, string savePath) => Convert(new Bitmap(image), savePath);

        public void Convert(Bitmap image, string savePath) {
            FileInfo newFile = new FileInfo(savePath);

            Bitmap zoomImage = Zoom(image);//縮放

            if (LowQuality) {//低色彩品質
                zoomImage = BitmapColorConvert(zoomImage);
            }

            var data = BitmapToColorArray(zoomImage);
            using (ExcelPackage excelPackage = new ExcelPackage(newFile)) {
                ExcelWorkbook workbook = excelPackage.Workbook;
                ExcelWorksheet worksheet = workbook.Worksheets.Add("pixelArt");

                worksheet.DefaultColWidth = 10.0 / 70.0 * PixelSize;//10 = 70px

                for (int row = 0; row < data[0].Length; row++) {
                    worksheet.Row(row + 1).Height = 10.0 / 13.0 * PixelSize;//10 = 13px
                    for (int col = 0; col < data.Length; col++) {
                        if(data[col][row].A == 0)continue;//透明背景

                        worksheet.Cells[row + 1, col + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row + 1, col + 1].Style.Fill.BackgroundColor.SetColor(data[col][row]);
                    }

                    Console.WriteLine($"完成進度: {Math.Round(row / (double)data[0].Length * 100)}%");
                }
                worksheet.Cells[data[0].Length + 1, 1].Value= "由ExcelPixelArt繪製，https://github.com/XuPeiYao/ExcelPixelArt/";
                worksheet.Row(data[0].Length + 1).Height = 20;

                Console.WriteLine("檔案儲存中");
                excelPackage.Save();
            }
        }
    }
}
