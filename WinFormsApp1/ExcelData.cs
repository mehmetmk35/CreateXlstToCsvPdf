using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1
{

    public class ProsesBilgileri
    {
        public int sarjNo { get; set; }
        public string receteAdi { get; set; }
        public List<string> sepetNumaralari { get; set; }
        public DateTime baslangicTarihi { get; set; }
        public int adimSayisi { get; set; }
        public string enBuyukSepetNo { get; set; }
    }

    public class AdimBilgisi
    {
        public string tabloAdi { get; set; }
        public int adimNumarasi { get; set; }
        public bool kimyasallikBilgisi { get; set; }
        public int? izinVerilenUygulamaSuresi { get; set; }
        public int minUygulamaSuresi { get; set; }
        public int maxUygulamaSüresi { get; set; }
        public int istenilenHavuzSicakligi { get; set; }
        public int istenilenHavuzSicaklikToleransi { get; set; }
        public DateTime havuzaGirisTarihi { get; set; }
        public DateTime havuzdanCikisTarihi { get; set; }
        public int uygulamaSuresi { get; set; }
        public int havuzSicakligi { get; set; }
        public string personelKimlikNo { get; set; }
        public bool kontrolAdimiVarlikBilgisi { get; set; }
        public string kontrolSonucu { get; set; }
        public string kontrolPersonelKimlikNo { get; set; }
        public DateTime? kontrolTarihi { get; set; }
    }

    public class ExcelKDynamiclTablo
    {
        public string TabloName { get; set; }
        public string TargetProp { get; set; }
    }
}
