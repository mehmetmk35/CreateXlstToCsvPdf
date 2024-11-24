
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
                   string maxDesiredExposureTime= GetPropertyValue(obj, "maxUygulamaS�resi");
                    value = $"{value}-{maxDesiredExposureTime} sn / {value}-{maxDesiredExposureTime} sec";
                }
                if (propertyName == "istenilenHavuzSicakligi")
                    value = $"{value} � 5�C";
                if (propertyName == "uygulamaSuresi")
                    value = $"{value} sn / {value} sec";
                if (propertyName == "havuzSicakligi")
                    value = $"{value}�C";
                return value?.ToString(); // Geriye string d�ner
            };
        }

        
       
        void CreateHeaderTable(ExcelWorksheet worksheet, ProsesBilgileri prosesBilgileri) {
            #region H�cre Birle�tirme ve D�zenleme
            // A1:C3 alan�n� birle�tir
            worksheet.Cells["A1:B3"].Merge = true;
            worksheet.Cells["A1:B3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A1:B3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["A1:B3"].Style.WrapText = true; // Metni sar

            // C1:H3 alan�n� birle�tir
            worksheet.Cells["C1:H3"].Merge = true;
            worksheet.Cells["C1:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["C1:H3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["C1:H3"].Style.WrapText = true; // Metni sar

            // I1:J3 alan�n� birle�tir
            worksheet.Cells["I1:J3"].Merge = true;
            worksheet.Cells["I1:J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells["I1:J3"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            worksheet.Cells["I1:J3"].Style.WrapText = true; // Metni sar
            #endregion

            #region Logo Eklenmesi
            var picture = worksheet.Drawings.AddPicture("Logo", new FileInfo("C:\\Users\\mehme\\Desktop\\test\\Picture1.png"));
            picture.SetPosition(0, 0, 0, 0); // (Sat�r 0, Offset 0, S�tun 0, Offset 0)
            picture.SetSize(110, 63); // Geni�lik ve y�kseklik (pixel cinsinden)
            #endregion

            #region RichText ��erik D�zenleme
            // C1:H3 i�erik
            var richTextA1_C3 = worksheet.Cells["C1:H3"].RichText;
            var percinText = richTextA1_C3.Add("PER��N ��LEME FORMU");
            percinText.Size = 12;
            percinText.Color = System.Drawing.Color.Black;

            var rivetText = richTextA1_C3.Add("\nRivet Treatment Form");
            rivetText.Size = 10;
            rivetText.Color = System.Drawing.Color.Black;
            // I1:J3 i�erik
            var richTextN1_O3 = worksheet.Cells["I1:J3"].RichText;
            var sarjNoText = richTextN1_O3.Add("�arj No");
            sarjNoText.Size = 12;
            sarjNoText.Color = System.Drawing.Color.Black;

            var chargeNrText = richTextN1_O3.Add("\nCharge Nr.");
            chargeNrText.Size = 10;
            chargeNrText.Color = System.Drawing.Color.Black;

            var chargeNrValue = richTextN1_O3.Add($"\n     {prosesBilgileri.sarjNo}");
            chargeNrValue.Size = 11;
            chargeNrValue.Color = System.Drawing.Color.Black;
            #endregion

            #region Kenarl�k Ekleme
            // A1:B3 kenarl�k
            worksheet.Cells["A1:B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

            // C1:H3 kenarl�k
            worksheet.Cells["C1:H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

            // I1:J3 kenarl�k
            worksheet.Cells["I1:J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            #endregion
        }
        void CreateChemicalDynamicTable(ExcelWorksheet worksheet, List<ExcelKDynamiclTablo> data, AdimBilgisi adimBilgisiData, string headerText, ref int refstartRow)
        {
            int startHeaderRow = refstartRow + 1;
            int nextRow = 0;
            // Ba�l�k ekle (A:D birle�tirilmi�, sola yasl�)
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Merge = true;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Value = headerText;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.Font.Bold = true;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.Font.Size = 12;
            worksheet.Cells[$"A{startHeaderRow}:j{startHeaderRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            // Aktif sat�r�n alt�na 1 bo�luk b�rak
            int startRow = refstartRow + 2;

            for (int i = 0; i < data.Count; i++)
            {
                ExcelKDynamiclTablo excelKimyasalTablo = data[i];
                int currentRow = startRow + (i * 2); // Her veri i�in 2 sat�r kullan�l�r
                  nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ay�r
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar h�cresi
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

                // F : h�cresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O De�er h�cresi
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
            // Ba�l�k ekle (A:O birle�tirilmi�, sola yasl�)
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Merge = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Bold = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Size = 12;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);


            var richTextHeader=worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].RichText;
            var headerTopTextRich=richTextHeader.Add("Kontrol Ad�m�");
            headerTopTextRich.Size=12;
            headerTopTextRich.Color = System.Drawing.Color.Black;

            var headerBottomTextRich = richTextHeader.Add("Control Step");
            headerBottomTextRich.Size = 12;
            headerBottomTextRich.Color = System.Drawing.Color.Black;

            // Aktif sat�r�n alt�na 1 bo�luk b�rak
            int startRow = refstartRow + 2;

            for (int i = 0; i < data.Count; i++)
            {
                ExcelKDynamiclTablo excelKimyasalTablo = data[i];
                int currentRow = startRow + (i * 2); // Her veri i�in 2 sat�r kullan�l�r
                nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ay�r
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar h�cresi
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

                // F : h�cresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O De�er h�cresi
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
            // Ba�l�k ekle (A:J birle�tirilmi�, sola yasl�)
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Merge = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Value = headerText;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Bold = true;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Font.Size = 12;
            worksheet.Cells[$"A{startHeaderRow}:J{startHeaderRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            // Aktif sat�r�n alt�na 1 bo�luk b�rak
            int startRow = refstartRow + 2;

            for (int i = 0; i < data.Count; i++)
            {
                ExcelKDynamiclTablo excelKimyasalTablo = data[i];
                int currentRow = startRow + (i * 2); // Her veri i�in 2 sat�r kullan�l�r
                nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ay�r
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar h�cresi
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

                // F : h�cresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O De�er h�cresi
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
                int currentRow = startRow + (i * 2); // Her veri i�in 2 sat�r kullan�l�r
                nextRow = currentRow + 1; //  A1:B2  GIBI YANI 2 SATIR SECMEK ICIN EKLENDI

                // TabloName'i "-" karakteri ile ay�r
                var tabloNameParts = excelKimyasalTablo.TabloName.Split('-');
                // A:E Anahtar h�cresi
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

                // F : h�cresi
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Merge = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Value = ":";
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Font.Bold = true;
                worksheet.Cells[$"E{currentRow}:E{nextRow}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                // G:O De�er h�cresi
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
                // ...TableData Dinamik olarak sahalar� olu�tupundan sahalar� burdan al�r 
                var ExcelChemicaTableData = new List<ExcelKDynamiclTablo>
            {
                new() { TabloName = "Proses Ad�m No-Process Step No",TargetProp="adimNumarasi" },
                new() { TabloName = "�zin Verilen Uygulama S�resi-Allowed Exposure Time", TargetProp="izinVerilenUygulamaSuresi"},
                new() { TabloName = "�stenilen Uygulama S�resi-Desired Exposure Time", TargetProp="minUygulamaSuresi"},
                new() { TabloName = "�stenilen Havuz S�cakl���-Desired Bath Temperature", TargetProp="istenilenHavuzSicakligi"},
                new() { TabloName = "Havuza Giri� (Tarih, Saat)-Bath Entrance(Date,Time)", TargetProp="havuzaGirisTarihi"},
                new() { TabloName = "Havuzdan ��k�� (Tarih, Saat)-Bath Exit (Date,Time)", TargetProp="havuzdanCikisTarihi"},
                new() { TabloName = "Uygulama S�resi-Exposure Time", TargetProp="uygulamaSuresi"  },
                new() { TabloName = "Havuz S�cakl���- Bath Temperature", TargetProp="havuzSicakligi"},
                new() { TabloName = "Personel Kimlik No-Personnel Identification No", TargetProp="personelKimlikNo"}
                };
                var ExcelKWaterTableData = new List<ExcelKDynamiclTablo>
            {
            new() { TabloName = "Proses Ad�m No-Process Step No",TargetProp="adimNumarasi" },
            new() { TabloName = "�stenilen Uygulama S�resi-Desired Exposure Time", TargetProp="minUygulamaSuresi"},
            new() { TabloName = "�stenilen Havuz S�cakl���-Desired Bath Temperature", TargetProp="istenilenHavuzSicakligi"},
            new() { TabloName = "Havuza Giri� (Tarih, Saat)-Bath Entrance(Date,Time)", TargetProp="havuzaGirisTarihi"},
            new() { TabloName = "Havuzdan ��k�� (Tarih, Saat)-Bath Exit (Date,Time)", TargetProp="havuzdanCikisTarihi"},
            new() { TabloName = "Uygulama S�resi-Exposure Time", TargetProp="uygulamaSuresi"  },
            new() { TabloName = "Havuz S�cakl���- Bath Temperature", TargetProp="havuzSicakligi"},
            new() { TabloName = "Personel Kimlik No-Personnel Identification No", TargetProp="personelKimlikNo"}
                };
                var ExcelProcInfolTabloData = new List<ExcelKDynamiclTablo>
            {
            new() { TabloName = "Re�ete Ad�-Recipe Name",TargetProp="receteAdi" },
            new() { TabloName = "Seper Numaralar�-Basket Numbers", TargetProp="sepetNumaralari"},
            new() { TabloName = "Ba�lang�� Tarihi-Start Date", TargetProp="baslangicTarihi"},

                };
                var ExcelControlStepTabloData = new List<ExcelKDynamiclTablo>
        {
        new() { TabloName = "Kontrol Sonucu-Control Result",TargetProp="kontrolSonucu" },
        new() { TabloName = "Personel Kimlik No-Personnel Identification No", TargetProp="kontrolPersonelKimlikNo"},
        new() { TabloName = "Tarih,Saat-Date,Time", TargetProp="kontrolTarihi"},

            };

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Per�in ��leme Formu");
                    int refStartRow = 3;
                    #region Sayfa Ayarlar� (A4 ��kt�s� ��in)
                    // Sayfa ayarlar�
                    worksheet.PrinterSettings.Orientation = eOrientation.Portrait;
                    worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
                    // S�tun geni�likleri kald�r�ld�, sadece birle�tirilen alanlar ayarlanacak
                    #endregion
                    worksheet.Column(5).Width = 5; // E s�tununun geni�li�ini 20 birim yap
                                                   // Headr Tablo
                    CreateHeaderTable(worksheet, prosesBilgileri);
                    //Sabit 1. Tablo
                    CreateReportStaticTable(worksheet, ExcelProcInfolTabloData, prosesBilgileri, ref refStartRow);
                    // Dinamik tablo olu�tur 

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


                    // Dosyay� kaydet
                    await File.WriteAllBytesAsync(fullPath, package.GetAsByteArray());
                    //hata al�nmad�g�nda excel ba�ar� bir �ekilde kaydetilecektir
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
                receteAdi = "Kapalama - S�kmekkkkkkkkk",
                sepetNumaralari = new List<string> { "25L", "38S", "56M" },
                baslangicTarihi = DateTime.Parse("16.11.2024 17:18:30"),
                adimSayisi = 3,
                enBuyukSepetNo = "56M"
            };
            var AdimBilgileri = new List<AdimBilgisi>
                {
                new AdimBilgisi
                {
                    tabloAdi = "1. Alkali A��nd�rma / Alkaline Pickling * (80-T-35-0110)",
                    adimNumarasi = 1,
                    kimyasallikBilgisi = true,
                    izinVerilenUygulamaSuresi = 60,
                    minUygulamaSuresi = 45,
                    maxUygulamaS�resi = 60,
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
                    tabloAdi = "2. Y�kama Havuzu / Rinsing Bath * (80-T-35-0110/80-T-35-0090)",
                    adimNumarasi = 2,
                    kimyasallikBilgisi = false,
                    izinVerilenUygulamaSuresi = 0,
                    minUygulamaSuresi = 120,
                    maxUygulamaS�resi = 180,
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
                    maxUygulamaS�resi = 180,
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
        ///  /// <param name="excelFilePath">Giri� Excel dosyas�n�n tam yolu</param>
        /// <param name="csvFilePath"> Olu�turulacak CSV dosyas�n�n tam yolu.</param>
        /// </summary>
        public async Task ConvertExcelToCsvAsync(string excelFilePath, string csvFilePath, int sheetNumber = 1)
        {
            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    // Belirtilen sayfay� al
                    var worksheet = workbook.Worksheet(sheetNumber);

                    // CSV olarak kaydetmek i�in StreamWriter kullanarak her sat�r� yaz
                    using (var writer = new StreamWriter(csvFilePath))
                    {
                        // Sayfan�n her sat�r�n� dola�
                        foreach (var row in worksheet.RowsUsed())
                        {
                            var rowValues = new List<string>();

                            // H�creleri gez ve de�erlerini al
                            foreach (var cell in row.Cells())
                            {
                                // H�cre de�erlerini kontrol et, bo�sa varsay�lan de�er ata
                                var cellValue = cell.IsEmpty() ? "" : cell.GetValue<string>();
                                // H�cre de�erlerini CSV'ye uygun hale getir
                                rowValues.Add(cellValue.Replace(",", "")); // Virg�l ka����
                            }

                            // Sat�r� CSV format�nda yaz
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
        ///  /// <param name="excelFilePath">Giri� Excel dosyas�n�n tam yolu</param>
        /// <param name="pdfFilePath"> Olu�turulacak PDF dosyas�n�n tam yolu.</param>
        /// </summary>
        public void ConvertExcelToPdfAsync(string excelFilePath, string pdfFilePath)
        {
            try
            {
                // Workbook nesnesi olu�tur ve Excel dosyas�n� y�kle
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(excelFilePath);

                // Sayfa ayarlar�n� uygula
                foreach (Spire.Xls.Worksheet sheet in workbook.Worksheets)
                {
                    // T�m i�eri�i sayfaya s��d�r
                    sheet.PageSetup.FitToPagesWide = 1;
                    sheet.PageSetup.FitToPagesTall = 1;

                    sheet.PageSetup.Zoom = 98;
                    // Y�nlendirme ayarla (gerekiyorsa)
                    sheet.PageSetup.Orientation = PageOrientationType.Portrait; // veya Portrait
                }

                // PDF olarak kaydet
                workbook.SaveToFile(pdfFilePath, FileFormat.PDF);


            }
            catch (Exception ex)
            {
                Console.WriteLine("Hata olu�tu: " + ex.Message);
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
                Console.WriteLine("Ge�ersiz parametreler. ��lem yap�lmad�.");
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

            string excelFileName = fileName + ".xlsx"; // Excel dosya ad�
            string csvFileName = fileName + ".csv";   // CSV dosya ad�          
            string pdfFileName = fileName + ".pdf";   // PDF dosya ad�          
            // Dosya yollar�n� olu�tur
            string excelFilePath = Path.Combine(ifExistsPath, excelFileName);
            string csvFilePath = Path.Combine(ifExistsPath, csvFileName);
            string pdfFilePath = Path.Combine(ifExistsPath, pdfFileName);
            if (await CreateExcelAsync(prosesBilgileri, adimBilgisis, excelFilePath))
            {
               
                await ConvertExcelToCsvAsync(excelFilePath, csvFilePath);
                ConvertExcelToPdfAsync(excelFilePath, pdfFilePath);
                PrintFile(excelFilePath);// default yaz�c�ya excel i�in istek at�l�yor 
            }

        }

         void PrintFile(string filePath)
        {
            try
            {
                Process printProcess = new Process();
                printProcess.StartInfo.FileName = filePath;// yazd�r�lacak dosyan�n tam yolu 
                printProcess.StartInfo.Verb = "Print"; //"Print" de�eri, dosyay� varsay�lan yaz�c�ya yazd�rmak i�in i�letim sistemine talimat veriyor.
                printProcess.StartInfo.CreateNoWindow = true;
                printProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                printProcess.Start();

                //Yazd�rma i�lemi g�nderildi
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }
    }
}
