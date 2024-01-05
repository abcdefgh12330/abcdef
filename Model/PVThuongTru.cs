using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ontap3.Model
{
    internal class PVThuongTru : PV
    {
        public double phuCap { get; set; }

        public double Luong()
        {
            return base.LuongCoBan() + phuCap;
        }
    }
}
