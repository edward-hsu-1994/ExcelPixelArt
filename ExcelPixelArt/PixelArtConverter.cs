using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Drawing;

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

        private Bitmap Zoom(Bitmap image) {
            //ref : http://stackoverflow.com/questions/23879178/zoom-bitmap-image 
            Bitmap result = new Bitmap(image.Width / Clarity, image.Height / Clarity);
            Graphics g = Graphics.FromImage(result);
            Rectangle srcRect = new Rectangle(0,0,image.Width,image.Height);
            Rectangle dstRect = new Rectangle(0,0,result.Width,result.Height);
            g.DrawImage(image, dstRect, srcRect, GraphicsUnit.Pixel);

            return result;
        }


        public IWorkbook Convert(string filePath) => Convert(new Bitmap(filePath));

        public IWorkbook Convert(Image image) => Convert(new Bitmap(image));

        public IWorkbook Convert(Bitmap image) {
            Bitmap zoomImage = Zoom(image);//縮放處裡
            IWorkbook result = new XSSFWorkbook();//Excel 2007+
            ISheet sheet = result.CreateSheet("pixelArt");//產生工作表
            double temp = zoomImage.Width * zoomImage.Height;
            for(int col = 0; col < zoomImage.Width; col++) {
                sheet.SetColumnWidth(col, 37 * PixelSize);//1000 = 27px 1px
            }

            for (int row = 0; row < zoomImage.Height; row++) {
                var Row = sheet.CreateRow(row);
                var temp2 = row * zoomImage.Width;
                Row.Height = (short)(15* PixelSize);//60 = 4px  1px = 15
                for (int col = 0; col < zoomImage.Width; col++) {
                    Console.WriteLine((temp2 + col) / temp);
                    var Cell = Row.CreateCell(col);
                    
                    //Cell.SetCellValue("GG");
                    var color = new XSSFColor(zoomImage.GetPixel(col, row));


                    XSSFCellStyle style = (XSSFCellStyle)result.CreateCellStyle();
                    style.SetFillForegroundColor(color);
                    style.FillPattern = FillPattern.SolidForeground;

                    Cell.CellStyle = style;
                }
            }

            
            
            return result;
        }
    }
}
