using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using ExcelReader;
using Excel = Microsoft.Office.Interop.Excel;


namespace QRCodeGenerator
{
    public partial class QRCode : Form
    {
        Reader reader=new Reader();
        List<Account> accounts=new List<Account>();
        List<Account> newAccounts=new List<Account>();
        Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        
        //Spreadsheet spreadsheet = new Spreadsheet();
       
        public QRCode()
        {
            InitializeComponent();
        }

        private void excelFileGetButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel 2003 (*.xls)|*.xls|Excel 2007 (*.xlsx)|*.xlsx";
            dialog.InitialDirectory = @"d:\";
            dialog.Title = "Please select an image file to encrypt.";
            if (dialog.ShowDialog()==DialogResult.OK)
            {
                string path = dialog.FileName;
                DataTable dt = reader.ExcelFile(path, "Sheet1");
                accounts = dt.AsEnumerable().Select(dataRow => new Account()
                {
                    Bank = dataRow[0].ToString(),
                    ReqId = dataRow[1].ToString(),
                    AccountNo = dataRow[2].ToString(),
                    Name = dataRow[3].ToString(),
                    Qty = Convert.ToInt32(dataRow[4]),
                    Branch = (dataRow[5].ToString()),
                    Currency = dataRow[6].ToString(),

                    StartStNo = (dataRow[7].ToString()),
                    EndStNo = (dataRow[8].ToString()),
                    Routing = dataRow[9].ToString(),
                    MICRAc = dataRow[10].ToString(),

                    Type = dataRow[11].ToString(),
                    
                        
                }).ToList();
                //accounts.Insert(0, new Account() { ReqId = dt.Columns[0].ColumnName, AccountNo = dt.Columns[1].ColumnName, Name = dt.Columns[2].ColumnName, Qty = Convert.ToInt32(dt.Columns[3].ColumnName), StartStNo = (dt.Columns[4].ColumnName), EndStNo = (dt.Columns[5].ColumnName), Routing = dt.Columns[6].ColumnName, Type = dt.Columns[7].ColumnName, Branch = dt.Columns[8].ColumnName });
                foreach (Account account in accounts)
                {
                    string s = "";
                    string a = "";
                    
                    int serialNo = Convert.ToInt32(account.StartStNo);
                    for (int i = 0; i < account.Qty; i++)
                    {

                        if (serialNo.ToString().Count() < 7)
                        {
                            a = "0" + serialNo;
                        }
                            s = null;
                            s += account.Name + " " + account.AccountNo + " " + account.Routing + " " + a + " " +
                                 account.Branch;
                            QRCoder.QRCodeGenerator qrGenerator = new QRCoder.QRCodeGenerator();
                            //QRCoder.QRCodeGenerator.QRCode qrCode = qrGenerator.CreateQrCode(s, QRCodeGenerator.ECCLevel.l);
                        QRCoder.QRCodeGenerator.QRCode qrCode = qrGenerator.CreateQrCode(s,
                            QRCoder.QRCodeGenerator.ECCLevel.H);
                            Bitmap imgOutput = qrCode.GetGraphic(1);
                        newAccounts.Add(new Account()
                        {
                            Bank = account.Bank, AccountNo = account.AccountNo, ReqId = account.ReqId, MICRAc = account.MICRAc, Currency = account.Currency, Branch = account.Branch, EndStNo = account.EndStNo, Name = account.Name, Qty = 1, Routing = account.Routing, Type = account.Type, StartStNo = a, Image = imgOutput
                        });
                        //imgOutput = ResizeImage(imgOutput, imgOutput.Width, imgOutput.Height);
                            ////Bitmap imgOutput = new Bitmap(@"D:\image.jpg");
                            //Graphics outputGraphics = Graphics.FromImage(imgOutput);

                            //EncoderParameters myEncoderParameters = new EncoderParameters(3);
                            //myEncoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L);
                            //myEncoderParameters.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.ScanMethod, (int)EncoderValue.ScanMethodInterlaced);
                            //myEncoderParameters.Param[2] = new EncoderParameter(System.Drawing.Imaging.Encoder.RenderMethod, (int)EncoderValue.RenderProgressive);

                            //ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
                            //ImageCodecInfo ici = null;
                            //foreach (ImageCodecInfo codec in codecs)
                            //{
                            //    if (codec.MimeType == "image/jpeg")
                            //        ici = codec;
                            //}

                            //imgOutput.Save(@"D:\Image\" + a +".jpg", ici, myEncoderParameters);


                       //imgOutput.Save(@"D:\Image\" + a + ".jpg", ImageFormat.Jpeg);
                        serialNo++;

                    }
                }
                CreateExcel(newAccounts);
                dataGridView1.DataSource = newAccounts;
                MessageBox.Show("Successfully Done", "Message", MessageBoxButtons.OK);
                // spreadsheet.ImportFromList(newAccounts);

            }
        }



        private void Expotr()
        {
            int rowsTotal = 0;
            int colsTotal = 0;
            Excel.Application xlApp = new Excel.Application();
            try
            {
                Excel.Workbook excelBook = xlApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet) excelBook.Worksheets[1];
                xlApp.Visible = true;
                rowsTotal = dataGridView1.RowCount;
                colsTotal = dataGridView1.Columns.Count;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (int i = 0; i < colsTotal; i++)
                {
                    //_with1.Cells[1, iC + 1].Value = dataGridView1.Columns[iC].HeaderText;
                }
            }
            catch ( Exception exception)
            {
                
            }
        }

        private void CreateExcel(List<Account> accounts )
        {
            

            
            
            excelApp.Visible = false;
            excelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = excelApp.ActiveSheet;
            
            worksheet.Cells[1, "A"] ="Bank";
            worksheet.Cells[1, "B"] = "Req Id";
            worksheet.Cells[1, "C"] = "Account No";
            worksheet.Cells[1, "D"] = "Name";
            worksheet.Cells[1, "E"] ="Leaves Qty";
            worksheet.Cells[1, "F"] ="Delevery Branch";
            worksheet.Cells[1, "G"] = "Currency";
            worksheet.Cells[1, "H"] = "St No";
            worksheet.Cells[1, "I"] = "End No";
            worksheet.Cells[1, "J"] = "Routing";
            worksheet.Cells[1, "K"] = "MICR  A/c";
            worksheet.Cells[1, "L"] = "TC";
            worksheet.Cells[1, "M"] = "Image";
            int row = 1;
            foreach (Account account in accounts)
            {

                row++;
                worksheet.Cells[row, "A"] = account.Bank;
                worksheet.Cells[row, "B"].NumberFormat = "@";
                worksheet.Cells[row, "B"].Value = account.ReqId;
                worksheet.Cells[row, "C"].NumberFormat = "@";
                worksheet.Cells[row, "C"] = account.AccountNo;
                worksheet.Cells[row, "D"] = account.Name;
                worksheet.Cells[row, "E"] = 1.ToString();
                worksheet.Cells[row, "F"] = account.Branch;
                worksheet.Cells[row, "G"] = account.Currency;
                worksheet.Cells[row, "H"].NumberFormat = "@";
                worksheet.Cells[row, "H"] = account.StartStNo;
                worksheet.Cells[row, "I"].NumberFormat = "@";
                worksheet.Cells[row, "I"] = account.EndStNo;
                worksheet.Cells[row, "J"].NumberFormat = "@";
                worksheet.Cells[row, "J"] = account.Routing;
                worksheet.Cells[row, "K"].NumberFormat = "@";
                worksheet.Cells[row, "K"] = account.MICRAc;
                worksheet.Cells[row, "L"].NumberFormat = "@";
                worksheet.Cells[row, "L"] = account.Type;
                //worksheet.Cells[row, "M"] = account.Image;
                //worksheet.Cells[row, "M"].Shapes.AddPicture(@"C:\Users\Public\Pictures\Sample Pictures\Chrysanthemum - Copy.jpg",Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 49, 49, 49,49);



               // var commnetImage = worksheet.Shapes.AddPicture(@"C:\Users\Public\Pictures\Sample Pictures\Chrysanthemum - Copy.jpg", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse,
           // 0, 0, 100, 200);

                Excel.Range oRange = (Excel.Range)worksheet.Cells[row, "M"];
                oRange.set_Item(1, 1, account.Image);
                worksheet.Paste(oRange, account.Image);
                


            }
            //worksheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);
            if (File.Exists(@"D:Data\BankAsia.xlsx"))
            {
                File.Delete(@"D:Data\BankAsia.xlsx");
            }
            worksheet.SaveAs(@"D:Data\BankAsia.xlsx");
            excelApp.Quit();
        }







        public DataTable ConvertToDataTable<T>(List<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;

        }




        public  System.Drawing.Bitmap ResizeImage(System.Drawing.Image image, int width, int height)
        {
            //a holder for the result
            Bitmap result = new Bitmap(width, height);
            //set the resolutions the same to avoid cropping due to resolution differences
            result.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            //use a graphics object to draw the resized image into the bitmap
            using (Graphics graphics = Graphics.FromImage(result))
            {
                //set the resize quality modes to high quality
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                //draw the image into the target bitmap
                graphics.DrawImage(image, 0, 0, result.Width, result.Height);
            }

            //return the resulting bitmap
            return result;
        }

    }
}
