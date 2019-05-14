using System;
using System.Drawing;
using System.Windows.Forms;
using Xceed.Words.NET;
using Microsoft.Office.Interop.Excel;

namespace PaintLibrary
{
    public class Painter
    {
        private int X1;
        private int Y1;
        private int X2;
        private int Y2;
  
        public PictureBox pictureBox { get; set; }
        public int complexity { get; set; }
        public Color color;

        private static string path = @"D:\Study\4sem\ДЕЛФИ курсовая\Saved\";

        public Painter()
        {
            x1 = 150;
            y1 = 100;
            x2 = 250;
            y2 = 200;
        }

        public Painter(PictureBox pictureBox, int complexity, Color color) : this()
        {
            this.pictureBox = pictureBox;
            this.complexity = complexity;
            this.color = color;
        }

        public void Draw()
        {
            DragonFunction(x1, y1, x2, y2, complexity, pictureBox, color);
        }

        void DragonFunction(int x1, int y1, int x2, int y2, int n, PictureBox pictureBox, Color color)
        {
            int xn, yn;
            Graphics g = Graphics.FromHwnd(pictureBox.Handle);
            var drawingPen = new Pen(color, 1);

            if (n > 0)
            {
                xn = (x1 + x2) / 2 + (y2 - y1) / 2;
                yn = (y1 + y2) / 2 - (x2 - x1) / 2;

                DragonFunction(x2, y2, xn, yn, n - 1, pictureBox, color);
                DragonFunction(x1, y1, xn, yn, n - 1, pictureBox, color);
            }

            var point1 = new System.Drawing.Point(x1, y1);
            var point2 = new System.Drawing.Point(x2, y2);
            g.DrawLine(drawingPen, point1, point2);
        }

        public static void SaveAllData(PictureBox pictureBox1, int complexity)
        {
            //Сохранение картинки
            System.Drawing.Rectangle r = pictureBox1.RectangleToScreen(pictureBox1.ClientRectangle);
            Bitmap b = new Bitmap(r.Width, r.Height);
            Graphics g = Graphics.FromImage(b);
            g.CopyFromScreen(r.Location, new System.Drawing.Point(0, 0), r.Size);
            b.Save(path + "Image.jpg");


            //Сохранение в WORD
            string pathDocument = path + "DragonCurveInfo.docx";
            string pathImage = path + "Image.jpg";
            DocX document = DocX.Create(pathDocument);

            // вставляем параграф и передаём текст
            document.InsertParagraph("Complexity: " + complexity).
                     // устанавливаем шрифт
                     Font("Calibri").
                     // устанавливаем размер шрифта
                     FontSize(15).
                     // устанавливаем цвет
                     Color(Color.Black).
                     // делаем текст жирным
                     Bold().
                     // устанавливаем интервал между символами
                     Spacing(5).
                     // выравниваем текст по центру
                     Alignment = Alignment.left;
            Painter painter = new Painter();
            document.InsertParagraph("X1: " + painter.x1.ToString() +
                                    "\nX2: " + painter.x2.ToString() +
                                    "\nY1: " + painter.y1.ToString() +
                                    "\nY2: " + painter.y2.ToString()).
                     Font("Calibri").
                     FontSize(15).
                     Color(Color.Black).
                     Bold().
                     Spacing(5).
                     Alignment = Alignment.left;

            Xceed.Words.NET.Image image = document.AddImage(pathImage);
            // создание параграфа
            Paragraph paragraph = document.InsertParagraph();
            // вставка изображения в параграф
            paragraph.AppendPicture(image.CreatePicture());
            // выравнивание параграфа по центру
            paragraph.Alignment = Alignment.center;
            // сохраняем документ
            document.Save();

            //Сохранение в EXCEL
            string pathExcel = path + "DragonCurveInfo.xls";

            // Создаём экземпляр нашего приложения
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Создаём экземпляр рабочий книги Excel
            Workbook workBook;
            // Создаём экземпляр листа Excel
            Worksheet workSheet;

            workBook = excelApp.Workbooks.Add();
            workSheet = (Worksheet)workBook.Worksheets.get_Item(1);

            workSheet.Cells[1, 1] = "Complexity";
            workSheet.Cells[1, 2] = "X1";
            workSheet.Cells[1, 3] = "X2";
            workSheet.Cells[1, 4] = "Y1";
            workSheet.Cells[1, 5] = "Y2";
            workSheet.Cells[2, 1] = complexity.ToString();
            workSheet.Cells[2, 2] = painter.x1.ToString();
            workSheet.Cells[2, 3] = painter.x2.ToString();
            workSheet.Cells[2, 4] = painter.y1.ToString();
            workSheet.Cells[2, 5] = painter.y2.ToString();
            workSheet.Shapes.AddPicture(path + "Image.jpg",
                Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 30, 350, 250);
            //excelApp.Visible = true;
            //excelApp.UserControl = true;

            //excelApp.DefaultSaveFormat = Excel.XlFileFormat.xlExcel9795;
            excelApp.DefaultFilePath = pathExcel;
            excelApp.DisplayAlerts = false;
            workBook.SaveAs(pathExcel, XlFileFormat.xlWorkbookNormal,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                XlSaveAsAccessMode.xlExclusive,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing);
            excelApp.Quit();
        }

        public int x1 {
            get
            {
                return X1;
            }
            set {
                if (value > 1000)
                {
                    X1 = 1000;
                }
                else
                {
                    X1 = value;
                }   
            }
        }
        public int y1
        {
            get
            {
                return Y1;
            }
            set
            {
                if (value > 1000)
                {
                    Y1 = 1000;
                }
                else
                {
                    Y1 = value;
                }
            }
        }
        public int x2
        {
            get
            {
                return X2;
            }
            set
            {
                if (value > 1000)
                {
                    X2 = 1000;
                }
                else
                {
                    X2 = value;
                }
            }
        }
        public int y2
        {
            get
            {
                return Y2;
            }
            set
            {
                if (value > 1000)
                {
                    Y2 = 1000;
                }
                else
                {
                    Y2 = value;
                }
            }
        }
    }
}
