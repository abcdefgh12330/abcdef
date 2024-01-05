using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ontap3.Model
{
    internal class PVTaiToaSoan : PV
    {
        public int soGioLamThem { get; set; }

        public double Luong()
        {
            return base.LuongCoBan() + soGioLamThem*1.5*100000;
        }
    }
}
