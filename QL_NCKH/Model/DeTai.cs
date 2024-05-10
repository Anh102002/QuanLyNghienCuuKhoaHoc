using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QL_NCKH
{
    class DeTai
    {
        public string MaDeTai { get; set; }
        public string TenDeTai { get; set; }
        public string LinhVuc { get; set; }
        public DateTime NgayBatDau { get; set; }
        public DateTime NgayKetThuc { get; set; }
        public string TienDo { get; set; }
        public string Capdetai { get; set; }
        public string Khoa { get; set; }
        public int Nam { get; set; }

        public DeTai() { }
        public DeTai(string MaDeTai, string TenDeTai, string LinhVuc, DateTime NgayBatDau, DateTime NgayKetThuc, string TienDo, string Capdetai, string Khoa, int Nam)
        {
            this.MaDeTai = MaDeTai;
            this.TenDeTai = TenDeTai;
            this.LinhVuc = LinhVuc;
            this.NgayBatDau = NgayBatDau;
            this.NgayKetThuc = NgayKetThuc;
            this.TienDo = TienDo;
            this.Khoa = Khoa;
            this.Nam = Nam;
        }
    }
}
