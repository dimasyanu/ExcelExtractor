using ExcelExtractor.Attributes;
using ExcelExtractor.Contracts;

namespace ExcelExtractor.Models.ExcelModels
{
    public class MemberExcel : IExcelModel
    {
        [ColKey("nama")]
        public string Nama { get; set; } = "";

        [ColKey("sebutan")]
        public string Sebutan { get; set; } = "";

        [ColKey("cabang")]
        public string Cabang { get; set; } = "";

        [ColKey("akhir_datang")]
        public string AkhirDatang { get; set; } = "";

        [ColKey("reminder")]
        public string Reminder { get; set; } = "";

        [ColKey("admin")]
        public string Admin { get; set; } = "";

        [ColKey("awal")]
        public string Awal { get; set; } = "";

        [ColKey("akhir")]
        public string Akhir { get; set; } = "";

        [ColKey("status")]
        public string Status { get; set; } = "";

        [ColKey("wa")]
        public string Phone { get; set; } = "";
    }
}
