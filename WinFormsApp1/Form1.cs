
using ClosedXML.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Spire.Xls;
using System.Diagnostics;
using WinFormsApp1.Operation;
using ExcelHorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment;
using ExcelVerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment;
using Workbook = Spire.Xls.Workbook;


namespace WinFormsApp1
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static class PropertyHelper
        {
            public static Func<object, string, string> GetPropertyValue = (obj, propertyName) =>
            {
                if (obj == null || string.IsNullOrEmpty(propertyName))
                    return null;

                var property = obj.GetType().GetProperty(propertyName);
                if (property == null)
                    return null;

                var value = property.GetValue(obj);

                if (propertyName == "sepetNumaralari" && value is IEnumerable<object> Values)
                    value = string.Join("/", Values.Select(v => v.ToString()));
                if (propertyName == "kontrolSonucu" && value is IEnumerable<object> ControlStepValue)
                    value = string.Join("/", ControlStepValue.Select(v => v.ToString()));
                if (propertyName == "izinVerilenUygulamaSuresi")
                    value = $"maksimum {value} sn / maximum {value} sec";
                if (propertyName == "minUygulamaSuresi")
                {
                   string maxDesiredExposureTime= GetPropertyValue(obj, "maxUygulamaSüresi");
                    value = $"{value}-{maxDesiredExposureTime} sn / {value}-{maxDesiredExposureTime} sec";
                }
                if (propertyName == "istenilenHavuzSicakligi")
                    value = $"{value} ± 5°C";
                if (propertyName == "uygulamaSuresi")
                    value = $"{value} sn / {value} sec";
                if (propertyName == "havuzSicakligi")
                    value = $"{value}°C";
                return value?.ToString(); // Geriye string döner
            };
        }

        
       
        void CreateHeaderTable(ExcelWorksheet worksheet, ProsesBilgileri prosesBilgileri) {
            #region Hücre Birleþtirme ve Düzenleme
            // A1:C3 alanýný birleþtir
            worksheet.Cells["A1:B3"].Merge = true;
            worksheet.Cells["A1:B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A1:B3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["A1:B3"].Style.WrapText = true; // Metni sar

            // C1:H3 alanýný birleþtir
            worksheet.Cells["C1:H3"].Merge = true;
            worksheet.Cells["C1:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C1:H3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["C1:H3"].Style.WrapText = true; // Metni sar

            // I1:J3 alanýný birleþtir
            worksheet.Cells["I1:J3"].Merge = true;
            worksheet.Cells["I1:J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["I1:J3"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            worksheet.Cells["I1:J3"].Style.WrapText = true; // Metni sar
            #endregion

            #region Logo Eklenmesi
            var picture = worksheet.Drawings.AddPicture("Logo", new FileInfo("C:\\Users\\mehme\\Desktop\\test\\Picture1.png"));
            picture.SetPosition(0, 0, 0, 0); // (Satýr 0, Offset 0, Sütun 0, Offset 0)
            picture.SetSize(110, 63); // Geniþlik ve yükseklik (pixel cinsinden)
            #endregion

            #region RichText Ýçerik Düzenleme
            // C1:H3 içerik
            var richTextA1_C3 = worksheet.Cells["C1:H3"].RichText;
            var percinText = richTextA1_C3.Add("PERÇÝN ÝÞLEME FORMU");
            percinText.Size = 12;
            percinText.Color = System.Drawing.Color.Black;

            var rivetText = richTextA1_C3.Add("\nRivet Treatment Form");
            rivetText.Size = 10;
            rivetText.Color = System.Drawing.Color.Black;
            // I1:J3 içerik
            var richTextN1_O3 = worksheet.Cells["I1:J3"].RichText;
            var sarjNoText = richTextN1_O3.Add("Þarj No");
            sarjNoText.Size = 12;
            sarjNoText.Color = System.Drawing.Color.Black;

            var chargeNrText = richTextN1_O3.Add("\nCharge Nr.");
            chargeNrText.Size = 10;
            chargeNrText.Color = System.Drawing.Color.Black;

            var chargeNrValue = richTextN1_O3.Add($"\n     {prosesBilgileri.sarjNo}");
            chargeNrValue.Size = 11;
            chargeNrValue.Color = System.Drawing.Color.Black;
            #endregion

            #region Kenarlýk Ekleme
            // A1:B3 kenarlýk
            worksheet.Cells["A1:B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

            // C1:H3 kenarlýk
            worksheet.Cells["C1:H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

            // I1:J3 kenarlýk
            worksheet.Cells["I1:J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            #endregion
        }
        void CreateChemicalDynamicTable(ExcelWorksheet worksheet, List<ExcelKDynamiclTablo> data, AdimBilgisi adimBilgisiData, string headerText, ref int refstartRow)
        {
            int startHeaderRow = refstartRow + 1;
            int nextRow = 0;
            // Baþlýk ekle (A:D birleþtirilmiþ, sola yaslý)
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Merge = true;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Value = headerText;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.Font.Bold = true;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.Font.Size = 12;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            // Aktif satýrýn altýna 1 boþluk býrak
            int startRow = refstartRow + 2;

            for (int i = 0; i < data.Count; i++)
            {
                ExcelKDynamiclTablo excelKimyasalTablo = data[i];
                int currentRow = startRow + (i * 2); // Her veri için 2 satýr kullanýlýr
                  nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ayýr
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar hücresi
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Merge = true;               
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Font.Bold = true;

                var richText = worksheet.Cells[$"A{currentRow}:D{nextRow}"].RichText;
                var topText = richText.Add(tabloNameParts[0]);
                topText.Size = 12;
                topText.Color = System.Drawing.Color.Black;

                var bottomText = richText.Add($"\n{tabloNameParts[1]}");
                bottomText.Size = 10;
                bottomText.Color = System.Drawing.Color.Black;

                // F : hücresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O Deðer hücresi
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Merge = true;               
                string value = PropertyHelper.GetPropertyValue(adimBilgisiData, data[i].TargetProp);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Value = value;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Font.Bold = true;
                 
            }
            refstartRow = nextRow + 1;
        }
        void CreateControlStepTable(ExcelWorksheet worksheet, List<ExcelKDynamiclTablo> data, AdimBilgisi adimBilgisiData, ref int refstartRow)
        {
            int startHeaderRow = refstartRow + 1;
            int nextRow = 0;
            // Baþlýk ekle (A:O birleþtirilmiþ, sola yaslý)
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Merge = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Bold = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Size = 12;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


            var richTextHeader=worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].RichText;
            var headerTopTextRich=richTextHeader.Add("Kontrol Adýmý");
            headerTopTextRich.Size=12;
            headerTopTextRich.Color = System.Drawing.Color.Black;

            var headerBottomTextRich = richTextHeader.Add("Control Step");
            headerBottomTextRich.Size = 12;
            headerBottomTextRich.Color = System.Drawing.Color.Black;

            // Aktif satýrýn altýna 1 boþluk býrak
            int startRow = refstartRow + 2;

            for (int i = 0; i < data.Count; i++)
            {
                ExcelKDynamiclTablo excelKimyasalTablo = data[i];
                int currentRow = startRow + (i * 2); // Her veri için 2 satýr kullanýlýr
                nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ayýr
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar hücresi
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Merge = true;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Font.Bold = true;

                var richText = worksheet.Cells[$"A{currentRow}:D{nextRow}"].RichText;
                var topText = richText.Add(tabloNameParts[0]);
                topText.Size = 12;
                topText.Color = System.Drawing.Color.Black;

                var bottomText = richText.Add($"\n{tabloNameParts[1]}");
                bottomText.Size = 10;
                bottomText.Color = System.Drawing.Color.Black;

                // F : hücresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O Deðer hücresi
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Merge = true;
                
                string value = PropertyHelper.GetPropertyValue(adimBilgisiData, data[i].TargetProp);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Value = value;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Font.Bold = true;

            }
            refstartRow = nextRow + 1;
        }
        void CreateWaterDynamicTable(ExcelWorksheet worksheet, List<ExcelKDynamiclTablo> data, AdimBilgisi adimBilgisiData, string headerText, ref int refstartRow)
        {
            int startHeaderRow = refstartRow + 1;
            int nextRow = 0;
            // Baþlýk ekle (A:J birleþtirilmiþ, sola yaslý)
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Merge = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Value = headerText;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Bold = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Size = 12;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            // Aktif satýrýn altýna 1 boþluk býrak
            int startRow = refstartRow + 2;

            for (int i = 0; i < data.Count; i++)
            {
                ExcelKDynamiclTablo excelKimyasalTablo = data[i];
                int currentRow = startRow + (i * 2); // Her veri için 2 satýr kullanýlýr
                nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ayýr
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar hücresi
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Merge = true;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Font.Bold = true;

                var richText = worksheet.Cells[$"A{currentRow}:D{nextRow}"].RichText;
                var topText = richText.Add(tabloNameParts[0]);
                topText.Size = 12;
                topText.Color = System.Drawing.Color.Black;

                var bottomText = richText.Add($"\n{tabloNameParts[1]}");
                bottomText.Size = 10;
                bottomText.Color = System.Drawing.Color.Black;

                // F : hücresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O Deðer hücresi
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Merge = true;
                
                string value = PropertyHelper.GetPropertyValue(adimBilgisiData, data[i].TargetProp);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Value = value;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Font.Bold = true;

            }
            refstartRow = nextRow + 1;
        }

        void CreateReportStaticTable(ExcelWorksheet worksheet, List<ExcelKDynamiclTablo> data, ProsesBilgileri prosesBilgileri, ref int refstartRow)
        {
            
            int startRow = refstartRow + 2;
            int nextRow = 0;
            for (int i = 0; i < data.Count; i++)
            {
                ExcelKDynamiclTablo excelKimyasalTablo = data[i];
                int currentRow = startRow + (i * 2); // Her veri için 2 satýr kullanýlýr
                nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ayýr
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar hücresi
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Merge = true;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"A{currentRow}:D{nextRow}"].Style.Font.Bold = true;

                var richText = worksheet.Cells[$"A{currentRow}:D{nextRow}"].RichText;
                var topText = richText.Add(tabloNameParts[0]);
                topText.Size = 12;
                topText.Color = System.Drawing.Color.Black;

                var bottomText = richText.Add($"\n{tabloNameParts[1]}");
                bottomText.Size = 10;
                bottomText.Color = System.Drawing.Color.Black;

                // F : hücresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O Deðer hücresi
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Merge = true;
                string value = PropertyHelper.GetPropertyValue(prosesBilgileri, data[i].TargetProp);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Value = value;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.WrapText = true;
                worksheet.Cells[$"F{currentRow}:J{nextRow}"].Style.Font.Bold = true;
                
            }
            refstartRow = nextRow + 1;
        }
        
        async Task<bool> CreateExcelAsync(ProsesBilgileri prosesBilgileri, List<AdimBilgisi> listAdimBilgisi, string fullPath)
        {
            try
            {
                
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // ...TableData Dinamik olarak sahalarý oluþtupundan sahalarý burdan alýr 
                var ExcelChemicaTableData = new List<ExcelKDynamiclTablo>
            {
                new() { TabloName = "Proses Adým No-Process Step No",TargetProp="adimNumarasi" },
                new() { TabloName = "Ýzin Verilen Uygulama Süresi-Allowed Exposure Time", TargetProp="izinVerilenUygulamaSuresi"},
                new() { TabloName = "Ýstenilen Uygulama Süresi-Desired Exposure Time", TargetProp="minUygulamaSuresi"},
                new() { TabloName = "Ýstenilen Havuz Sýcaklýðý-Desired Bath Temperature", TargetProp="istenilenHavuzSicakligi"},
                new() { TabloName = "Havuza Giriþ (Tarih, Saat)-Bath Entrance(Date,Time)", TargetProp="havuzaGirisTarihi"},
                new() { TabloName = "Havuzdan Çýkýþ (Tarih, Saat)-Bath Exit (Date,Time)", TargetProp="havuzdanCikisTarihi"},
                new() { TabloName = "Uygulama Süresi-Exposure Time", TargetProp="uygulamaSuresi"  },
                new() { TabloName = "Havuz Sýcaklýðý- Bath Temperature", TargetProp="havuzSicakligi"},
                new() { TabloName = "Personel Kimlik No-Personnel Identification No", TargetProp="personelKimlikNo"}
                };
                var ExcelKWaterTableData = new List<ExcelKDynamiclTablo>
            {
            new() { TabloName = "Proses Adým No-Process Step No",TargetProp="adimNumarasi" },
            new() { TabloName = "Ýstenilen Uygulama Süresi-Desired Exposure Time", TargetProp="minUygulamaSuresi"},
            new() { TabloName = "Ýstenilen Havuz Sýcaklýðý-Desired Bath Temperature", TargetProp="istenilenHavuzSicakligi"},
            new() { TabloName = "Havuza Giriþ (Tarih, Saat)-Bath Entrance(Date,Time)", TargetProp="havuzaGirisTarihi"},
            new() { TabloName = "Havuzdan Çýkýþ (Tarih, Saat)-Bath Exit (Date,Time)", TargetProp="havuzdanCikisTarihi"},
            new() { TabloName = "Uygulama Süresi-Exposure Time", TargetProp="uygulamaSuresi"  },
            new() { TabloName = "Havuz Sýcaklýðý- Bath Temperature", TargetProp="havuzSicakligi"},
            new() { TabloName = "Personel Kimlik No-Personnel Identification No", TargetProp="personelKimlikNo"}
                };
                var ExcelProcInfolTabloData = new List<ExcelKDynamiclTablo>
            {
            new() { TabloName = "Reçete Adý-Recipe Name",TargetProp="receteAdi" },
            new() { TabloName = "Seper Numaralarý-Basket Numbers", TargetProp="sepetNumaralari"},
            new() { TabloName = "Baþlangýç Tarihi-Start Date", TargetProp="baslangicTarihi"},

                };
                var ExcelControlStepTabloData = new List<ExcelKDynamiclTablo>
        {
        new() { TabloName = "Kontrol Sonucu-Control Result",TargetProp="kontrolSonucu" },
        new() { TabloName = "Personel Kimlik No-Personnel Identification No", TargetProp="kontrolPersonelKimlikNo"},
        new() { TabloName = "Tarih,Saat-Date,Time", TargetProp="kontrolTarihi"},

            };

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Perçin Ýþleme Formu");
                    int refStartRow = 3;
                    #region Sayfa Ayarlarý (A4 Çýktýsý Ýçin)
                    // Sayfa ayarlarý
                    worksheet.PrinterSettings.Orientation = eOrientation.Portrait;
                    worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
                    // Sütun geniþlikleri kaldýrýldý, sadece birleþtirilen alanlar ayarlanacak
                    #endregion
                    worksheet.Column(5).Width = 5; // E sütununun geniþliðini 20 birim yap
                                                   // Headr Tablo
                    CreateHeaderTable(worksheet, prosesBilgileri);
                    //Sabit 1. Tablo
                    CreateReportStaticTable(worksheet, ExcelProcInfolTabloData, prosesBilgileri, ref refStartRow);
                    // Dinamik tablo oluþtur 

                    foreach (AdimBilgisi adimData in listAdimBilgisi)
                    {

                        if (adimData.kimyasallikBilgisi)
                        {
                            //Kimyasal true ise eklenecek tablo 
                            CreateChemicalDynamicTable(worksheet, ExcelChemicaTableData, adimData, adimData.tabloAdi, ref refStartRow);
                            if (adimData.kontrolAdimiVarlikBilgisi)
                                CreateControlStepTable(worksheet, ExcelControlStepTabloData, adimData, ref refStartRow);
                        }

                        else
                        {  //Kimyasal false ise eklenecek tablo 
                            CreateWaterDynamicTable(worksheet, ExcelKWaterTableData, adimData, adimData.tabloAdi, ref refStartRow);
                            if (adimData.kontrolAdimiVarlikBilgisi)
                                CreateControlStepTable(worksheet, ExcelControlStepTabloData, adimData, ref refStartRow);
                        }

                    }


                    // Dosyayý kaydet
                    await File.WriteAllBytesAsync(fullPath, package.GetAsByteArray());
                    //hata alýnmadýgýnda excel baþarý bir þekilde kaydetilecektir
                    return true;
                }
            }
            catch (Exception  ex)
            {
                return false;
                //todo log
            }
        }
        private async void button2_Click(object sender, EventArgs e)
        {

            var ProsesBilgileri = new ProsesBilgileri
            {
                sarjNo = 12345,
                receteAdi = "Kapalama - Sökmekkkkkkkkk",
                sepetNumaralari = new List<string> { "25L", "38S", "56M" },
                baslangicTarihi = DateTime.Parse("16.11.2024 17:18:30"),
                adimSayisi = 3,
                enBuyukSepetNo = "56M"
            };
            var AdimBilgileri = new List<AdimBilgisi>
                {
                new AdimBilgisi
                {
                    tabloAdi = "1. Alkali Aþýndýrma / Alkaline Pickling * (80-T-35-0110)",
                    adimNumarasi = 1,
                    kimyasallikBilgisi = true,
                    izinVerilenUygulamaSuresi = 60,
                    minUygulamaSuresi = 45,
                    maxUygulamaSüresi = 60,
                    istenilenHavuzSicakligi = 60,
                    istenilenHavuzSicaklikToleransi = 2,
                    havuzaGirisTarihi = DateTime.Parse("16.11.2024 17:15:30"),
                    havuzdanCikisTarihi = DateTime.Parse("16.11.2024 17:18:30"),
                    uygulamaSuresi = 180,
                    havuzSicakligi = 57,
                    personelKimlikNo = "0123456789",
                    kontrolAdimiVarlikBilgisi = false,
                    kontrolSonucu = "",
                    kontrolPersonelKimlikNo = "",
                    kontrolTarihi = null
                },
                new AdimBilgisi
                {
                    tabloAdi = "2. Yýkama Havuzu / Rinsing Bath * (80-T-35-0110/80-T-35-0090)",
                    adimNumarasi = 2,
                    kimyasallikBilgisi = false,
                    izinVerilenUygulamaSuresi = 0,
                    minUygulamaSuresi = 120,
                    maxUygulamaSüresi = 180,
                    istenilenHavuzSicakligi = 25,
                    istenilenHavuzSicaklikToleransi = 5,
                    havuzaGirisTarihi = DateTime.Parse("16.11.2024 17:35:45"),
                    havuzdanCikisTarihi = DateTime.Parse("16.11.2024 17:35:50"),
                    uygulamaSuresi = 185,
                    havuzSicakligi = 28,
                    personelKimlikNo = "0123456789",
                    kontrolAdimiVarlikBilgisi = true,
                    kontrolSonucu = "OK",
                    kontrolPersonelKimlikNo = "aaaaaaaaaa",
                    kontrolTarihi = DateTime.Parse("16.11.2024 17:36:02")
                },
                new AdimBilgisi
                {
                    tabloAdi = "3. Durulama Havuzu / Flow Rinsing Bath * (80-T-35-0090)",
                    adimNumarasi = 3,
                    kimyasallikBilgisi = false,
                    izinVerilenUygulamaSuresi = 0,
                    minUygulamaSuresi = 120,
                    maxUygulamaSüresi = 180,
                    istenilenHavuzSicakligi = 25,
                    istenilenHavuzSicaklikToleransi = 5,
                    havuzaGirisTarihi = DateTime.Parse("16.11.2024 17:35:52"),
                    havuzdanCikisTarihi = DateTime.Parse("16.11.2024 17:35:58"),
                    uygulamaSuresi = 187,
                    havuzSicakligi = 26,
                    personelKimlikNo = "0123456789",
                    kontrolAdimiVarlikBilgisi = true,
                    kontrolSonucu = "OK",
                    kontrolPersonelKimlikNo = "0123456789",
                    kontrolTarihi = DateTime.Parse("16.11.2024 17:36:02")
                }
                };

            await GenerateExcel(ProsesBilgileri, AdimBilgileri);



        }

       
        /// <summary>
        ///  /// <param name="excelFilePath">Giriþ Excel dosyasýnýn tam yolu</param>
        /// <param name="csvFilePath"> Oluþturulacak CSV dosyasýnýn tam yolu.</param>
        /// </summary>
        public async Task ConvertExcelToCsvAsync(string excelFilePath, string csvFilePath, int sheetNumber = 1)
        {
            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    // Belirtilen sayfayý al
                    var worksheet = workbook.Worksheet(sheetNumber);

                    // CSV olarak kaydetmek için StreamWriter kullanarak her satýrý yaz
                    using (var writer = new StreamWriter(csvFilePath))
                    {
                        // Sayfanýn her satýrýný dolaþ
                        foreach (var row in worksheet.RowsUsed())
                        {
                            var rowValues = new List<string>();

                            // Hücreleri gez ve deðerlerini al
                            foreach (var cell in row.Cells())
                            {
                                // Hücre deðerlerini kontrol et, boþsa varsayýlan deðer ata
                                var cellValue = cell.IsEmpty() ? "" : cell.GetValue<string>();
                                // Hücre deðerlerini CSV'ye uygun hale getir
                                rowValues.Add(cellValue.Replace(",", "")); // Virgül kaçýþý
                            }

                            // Satýrý CSV formatýnda yaz
                           await  writer.WriteLineAsync(string.Join(",", rowValues));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hata: {ex.Message}");
            }
        }
        /// <summary>
        ///  /// <param name="excelFilePath">Giriþ Excel dosyasýnýn tam yolu</param>
        /// <param name="pdfFilePath"> Oluþturulacak PDF dosyasýnýn tam yolu.</param>
        /// </summary>
        public void ConvertExcelToPdfAsync(string excelFilePath, string pdfFilePath)
        {
            try
            {
                // Workbook nesnesi oluþtur ve Excel dosyasýný yükle
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(excelFilePath);

                // Sayfa ayarlarýný uygula
                foreach (Spire.Xls.Worksheet sheet in workbook.Worksheets)
                {
                    // Tüm içeriði sayfaya sýðdýr
                    sheet.PageSetup.FitToPagesWide = 1;
                    sheet.PageSetup.FitToPagesTall = 1;

                    sheet.PageSetup.Zoom = 98;
                    // Yönlendirme ayarla (gerekiyorsa)
                    sheet.PageSetup.Orientation = PageOrientationType.Portrait; // veya Portrait
                }

                // PDF olarak kaydet
                workbook.SaveToFile(pdfFilePath, FileFormat.PDF);


            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata oluþtu: " + ex.Message);
            }
        }

        /// <summary>
        ///  /// <param name="prosesBilgileri">params</param>
        /// <param name="adimBilgisis"> params</param>
        /// <param > PROGRAM START VERILDIGINDE TETIKLENECEK FUNC </param>
        /// </summary>
        public async Task  GenerateExcel(ProsesBilgileri prosesBilgileri, List<AdimBilgisi> adimBilgisis)
        {
            // Parametre kontrolleri
            if (prosesBilgileri == null || adimBilgisis == null || !adimBilgisis.Any())
            {
                Console.WriteLine("Geçersiz parametreler. Ýþlem yapýlmadý.");
                return;
            }
            string fileName = @$"{NameOperation.CharacterRegulatory(prosesBilgileri.receteAdi)}_"
                            + $"{NameOperation.CharacterRegulatory(prosesBilgileri.baslangicTarihi.ToString())}_"
                            + $"{NameOperation.CharacterRegulatory(prosesBilgileri.enBuyukSepetNo)}_"
                            + $"{NameOperation.CharacterRegulatory(prosesBilgileri.sarjNo.ToString())}";

            string pathOrContainerName = Configuration.Path;
            string ifExistsPathDateTime = Path.Combine(pathOrContainerName, DateTime.Now.ToString("yyy-MM-dd"));
            string ifExistsPath = Path.Combine(ifExistsPathDateTime, fileName);
            // Klosor Kontrol
            if (!Directory.Exists(ifExistsPathDateTime))
                Directory.CreateDirectory(ifExistsPathDateTime);

            if (!Directory.Exists(ifExistsPath))
                Directory.CreateDirectory(ifExistsPath);

            string excelFileName = fileName + ".xlsx"; // Excel dosya adý
            string csvFileName = fileName + ".csv";   // CSV dosya adý          
            string pdfFileName = fileName + ".pdf";   // PDF dosya adý          
            // Dosya yollarýný oluþtur
            string excelFilePath = Path.Combine(ifExistsPath, excelFileName);
            string csvFilePath = Path.Combine(ifExistsPath, csvFileName);
            string pdfFilePath = Path.Combine(ifExistsPath, pdfFileName);
            if (await CreateExcelAsync(prosesBilgileri, adimBilgisis, excelFilePath))
            {
               
                await ConvertExcelToCsvAsync(excelFilePath, csvFilePath);
                ConvertExcelToPdfAsync(excelFilePath, pdfFilePath);
                PrintFile(excelFilePath);// default yazýcýya excel için istek atýlýyor 
            }

        }

         void PrintFile(string filePath)
        {
            try
            {
                Process printProcess = new Process();
                printProcess.StartInfo.FileName = filePath;// yazdýrýlacak dosyanýn tam yolu 
                printProcess.StartInfo.Verb = "Print"; //"Print" deðeri, dosyayý varsayýlan yazýcýya yazdýrmak için iþletim sistemine talimat veriyor.
                printProcess.StartInfo.CreateNoWindow = true;
                printProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                printProcess.Start();

                //Yazdýrma iþlemi gönderildi
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }
    }
}
