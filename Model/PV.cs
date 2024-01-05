using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ontap3.Model
{
    internal class PV
    {
        public string MaPV { get; set; }
        public string HoTen { get; set; }
        public string GioiTinh { get; set; }
        public string SoDienThoai { get; set; }
        public DateTime NgayVaoLam { get; set; }
        public byte[] HinhAnhPV { get; set; }

        public double LuongCoBan()
        {
            return 12000000;
        }

        public int thamNien()
        {
            return DateTime.Now.Year - NgayVaoLam.Year;
        }

        public bool isYellowBackgroud()
        {
            return thamNien() > 5;
        }
        public PV()
        {
        }
        public PV (PhongVien pv)
        {
            MaPV = pv.MaPV;
            HoTen = pv.HoTen;
            GioiTinh = pv.GioiTinh;
            SoDienThoai = pv.SoDienThoai;
            NgayVaoLam = (DateTime)pv.NgayVaoLam;
        }
    }
}
