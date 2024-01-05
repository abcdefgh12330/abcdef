using ontap3.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace ontap3
{
    public partial class Form1 : Form
    {
        private QLPVEntities db = new QLPVEntities();
        private List<PhongVien> listPV = new List<PhongVien>();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            setDefault();
            hienThiDuLieuTuCSDL();
            loadDataToComboBox();
        }

        private void loadDataToComboBox()
        {
            List<bool?> listLoaiLV = db.PhongViens.Select(pv => pv.LoaiPV).Distinct().ToList();
            cb_DSLoaiPV.DataSource = listLoaiLV;
        }

        private void setDefault()
        {
            rad_Nam.Checked = true;
            dtp_NgayVaoLam.Value = DateTime.Now;
            rad_TaiToaSoan.Checked = true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát không?", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if(result == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void rad_TaiToaSoan_CheckedChanged(object sender, EventArgs e)
        {
            if (rad_TaiToaSoan.Checked)
            {
                lb_SoGioLam.Visible = true;
                lb_PhuCap.Visible = false;
                txt_SoGioLam.Visible = true;
                txt_PhuCap.Visible = false;
            }
            else
            {
                lb_SoGioLam.Visible = false;
                lb_PhuCap.Visible = true;
                txt_SoGioLam.Visible = false;
                txt_PhuCap.Visible = true;
            }
        }

        private void reset()
        {
            txt_MaNV.Text = string.Empty;
            txt_HoTen.Text  = string.Empty;
            txt_DienThoai.Text  = string.Empty;
            txt_PhuCap.Text = string.Empty;
            txt_SoGioLam.Text = string.Empty;
            setDefault();
            txt_MaNV.Focus();
        }

        private void btn_Them_Click(object sender, EventArgs e)
        {
            reset();
        }

        private bool isValidInfo()
        {
            if(String.IsNullOrEmpty(txt_MaNV.Text)|| String.IsNullOrEmpty(txt_HoTen.Text)|| String.IsNullOrEmpty(txt_DienThoai.Text))
            {
                MessageBox.Show("Vui lòng không bỏ trống thông tin!", "Cảnh báo");
                return false;
            }
            if(dtp_NgayVaoLam.Value > DateTime.Now)
            {
                MessageBox.Show("Vui lòng nhập ngày vào làm nhỏ hơn hoặc bằng ngày hiện tại!", "Cảnh báo");
                return false;
            }
            Regex regex = new Regex(@"^0[1-9]\d{9,11}$");
            if(regex.IsMatch(txt_DienThoai.Text)) {
                MessageBox.Show("Vui lòng nhập số điện thoại hợp lệ!", "Cảnh báo");
                return false;
            }
            return true;
        }

        private void luuVaoCSDL(PhongVien pv)
        {
            db.PhongViens.Add(pv);
            db.SaveChanges();
        }

        private void btn_Luu_Click(object sender, EventArgs e)
        {
            
            if(isValidInfo())
            {
                MemoryStream ms = new MemoryStream();
                byte[] imageBytes = null;
                if (pb_HinhAnhPhongVien.Image != null)
                {
                    pb_HinhAnhPhongVien.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    imageBytes = ms.ToArray();
                }
                PV pvien = null;
                double Luong=0;
                if (rad_TaiToaSoan.Checked && !String.IsNullOrEmpty(txt_SoGioLam.Text))
                {
                    pvien = new PVTaiToaSoan() { soGioLamThem = int.Parse(txt_SoGioLam.Text)};
                    Luong = ((PVTaiToaSoan)pvien).Luong();
                }
                else if (rad_ThuongTru.Checked && !String.IsNullOrEmpty(txt_PhuCap.Text))
                {
                    pvien = new PVThuongTru() { phuCap = int.Parse(txt_PhuCap.Text) };
                    Luong = ((PVThuongTru)pvien).Luong();
                }
                PhongVien pv = new PhongVien()
                {
                    MaPV = txt_MaNV.Text,
                    HoTen = txt_HoTen.Text,
                    GioiTinh = rad_Nam.Checked ? "Nam" : "Nữ",
                    SoDienThoai = txt_DienThoai.Text,
                    NgayVaoLam = dtp_NgayVaoLam.Value,
                    LoaiPV = rad_TaiToaSoan.Checked ? true : false,
                    HinhAnhPV = ms.ToArray(),
                    Luong = (decimal?)Luong
                };
                var timPV = db.PhongViens.Find(pv.MaPV);
                if(timPV == null)
                {
                    lv_DSPV.SelectedItems.Clear();
                    string[] data = { pv.MaPV, pv.HoTen, pv.GioiTinh, pv.NgayVaoLam.ToString() };
                    ListViewItem lvi = new ListViewItem(data);
                    PV pvModel = new PV(pv);
                    if (pvModel.isYellowBackgroud())
                    {
                        lvi.BackColor = Color.Yellow;
                    }
                    lvi.Selected = true;
                    lv_DSPV.Items.Add(lvi);
                    luuVaoCSDL(pv);
                }
                else
                {
                    MessageBox.Show("Mã phóng viên đã tồn tại!", "Cảnh báo");
                }
            }
        }

        private void btn_Xoa_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có muốn xóa Phóng viên này hay không?", "Câu hỏi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(lv_DSPV.SelectedItems.Count > 0 && result == DialogResult.Yes)
            {
                ListViewItem lvi = lv_DSPV.SelectedItems[0];
                int nextIndex = lvi.Index;
                db.PhongViens.Remove(db.PhongViens.Find(lvi.SubItems[0].Text));
                lv_DSPV.Items.Remove(lvi);
                if(lv_DSPV.Items.Count == 0)
                {
                    reset();
                }
                else if(lv_DSPV.Items.Count > nextIndex)
                {
                    lv_DSPV.Items[nextIndex].Selected = true;
                }
                MessageBox.Show("Đã xóa thành công!", "Thông báo");
                db.SaveChanges();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn 1 dòng để xóa!", "Thông báo");
            }
        }

        private void hienThiDuLieuTuCSDL()
        {
            listPV = db.PhongViens.ToList();
            foreach(PhongVien pv in listPV)
            {
                string[] data = { pv.MaPV, pv.HoTen, pv.GioiTinh, pv.NgayVaoLam.ToString() };
                ListViewItem lvi = new ListViewItem(data);
                PV pvModel = new PV(pv);
                if (pvModel.isYellowBackgroud())
                {
                    lvi.BackColor = Color.Yellow;
                }
                lv_DSPV.Items.Add(lvi);
            }
        }

        private void lv_DSPV_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(lv_DSPV.SelectedItems.Count > 0)
            {
                ListViewItem lvi = lv_DSPV.SelectedItems[0];
                PhongVien pv = db.PhongViens.Find(lvi.SubItems[0].Text);
                if(pv != null) 
                {
                    txt_MaNV.Text = pv.MaPV;
                    txt_HoTen.Text = pv.HoTen;
                    txt_DienThoai.Text = pv.SoDienThoai;
                    if(pv.GioiTinh == "Nam")
                    {
                        rad_Nam.Checked = true;
                    }
                    else
                    {
                        rad_Nu.Checked = true;
                    }
                    dtp_NgayVaoLam.Value = (DateTime)pv.NgayVaoLam;
                    if(pv.LoaiPV == true)
                    {
                        rad_TaiToaSoan.Checked = true;
                    }
                    else
                    {
                        rad_ThuongTru.Checked = true;
                    }
                    if (pv.HinhAnhPV.Length > 2)
                    {
                        MemoryStream ms = new MemoryStream(pv.HinhAnhPV);
                        pb_HinhAnhPhongVien.Image = Image.FromStream(ms);
                    }
                }
            }
        }

        private void btn_Sua_Click(object sender, EventArgs e)
        {
            if (lv_DSPV.SelectedItems.Count > 0)
            {
                ListViewItem lvi = lv_DSPV.SelectedItems[0];
                PhongVien pv = db.PhongViens.Find(lvi.SubItems[0].Text);
                if (pv != null)
                {
                    MemoryStream ms = new MemoryStream();
                    byte[] imageBytes = null;
                    if (pb_HinhAnhPhongVien.Image != null)
                    {
                        pb_HinhAnhPhongVien.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        imageBytes = ms.ToArray();
                    }
                    PV pvien = null;
                    double Luong = (double)pv.Luong;
                    if (rad_TaiToaSoan.Checked && !String.IsNullOrEmpty(txt_SoGioLam.Text))
                    {
                        pvien = new PVTaiToaSoan() { soGioLamThem = int.Parse(txt_SoGioLam.Text) };
                        Luong = ((PVTaiToaSoan)pvien).Luong();
                    }
                    else if (rad_ThuongTru.Checked && !String.IsNullOrEmpty(txt_PhuCap.Text))
                    {
                        pvien = new PVThuongTru() { phuCap = int.Parse(txt_PhuCap.Text) };
                        Luong = ((PVThuongTru)pvien).Luong();
                    }

                    pv.MaPV = txt_MaNV.Text;
                    lvi.SubItems[0].Text = pv.MaPV;
                    pv.HoTen = txt_HoTen.Text;
                    lvi.SubItems[1].Text = pv.HoTen;
                    pv.GioiTinh = rad_Nam.Checked ? "Nam" : "Nữ";
                    lvi.SubItems[2].Text = pv.GioiTinh;
                    pv.SoDienThoai = txt_DienThoai.Text;
                    pv.NgayVaoLam = dtp_NgayVaoLam.Value;
                    lvi.SubItems[3].Text = pv.NgayVaoLam.ToString();
                    pv.LoaiPV = rad_TaiToaSoan.Checked ? true : false;
                    pv.HinhAnhPV = ms.ToArray();
                    pv.Luong = (decimal?)Luong;
                    MessageBox.Show("Đã cập nhật thành công!", "Thông báo");
                    db.SaveChanges();
                }
                else
                {
                    MessageBox.Show("Không tìm thấy phóng viên được chọn", "Thông báo");
                }

            }
            else
            {
                MessageBox.Show("Vui lòng chọn 1 dòng để sửa!", "Thông báo");
            }
        }

        private void btn_SapXep_Click(object sender, EventArgs e)
        {
            listPV = (List<PhongVien>)db.PhongViens.OrderByDescending(pv => DateTime.Now.Year - pv.NgayVaoLam.Year).ThenBy(pv => pv.HoTen).ToList();
            lv_DSPV.Items.Clear();
            foreach (PhongVien pv in listPV)
            {
                string[] data = { pv.MaPV, pv.HoTen, pv.GioiTinh, pv.NgayVaoLam.ToString() };
                ListViewItem lvi = new ListViewItem(data);
                PV pvModel = new PV(pv);
                if (pvModel.isYellowBackgroud())
                {
                    lvi.BackColor = Color.Yellow;
                }
                lv_DSPV.Items.Add(lvi);
            }
        }

        private void btn_ThongKe_Click(object sender, EventArgs e)
        {
            int soPVTaiToaSoan = db.PhongViens.Where(pv => pv.LoaiPV == true).Count();
            int SoPVThuongTru = db.PhongViens.Where(pv => pv.LoaiPV == false).Count();
            double tongLuongPVTaiToaSoan = db.PhongViens.Where(pv => pv.LoaiPV == true).Sum(pv => (double)pv.Luong);
            double tongLuongPVThuongTru = db.PhongViens.Where(pv => pv.LoaiPV == false).Sum(pv => (double)pv.Luong);
            string thongBao = String.Format("Số phóng viên tại tòa soạn là: {0}\n"+
                                            "Số phóng viên thường trú là: {1}\n"+
                                            "Tổng lương trả cho phóng viên tại tòa soạn là: {2}\n"+
                                            "Tổng lương trả cho phóng viên thường trú là: {3}\n",soPVTaiToaSoan,SoPVThuongTru,tongLuongPVTaiToaSoan,tongLuongPVThuongTru);
            MessageBox.Show(thongBao, "Thống kê");
        }

        private void pb_HinhAnhPhongVien_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "ImageFiles(*.jpg; *.jpeg; *.png; *.bmp)|*.jpg; *.jpeg; *.png; *.bmp";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pb_HinhAnhPhongVien.Image =new Bitmap(openFileDialog.FileName);
            }
        }

        private void btn_XuatExcel_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWb = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet excelWs = excelWb.Worksheets[1];

            excelWs.Range["A1"].Font.Color = Color.Red;
            excelWs.Range["A1"].Font.Bold = true;
            excelWs.Range["A1"].Font.Size = 16;
            excelWs.Range["A1"].Value = "DANH SÁCH PHÓNG VIÊN";

            int row = 2;
            
            List<PhongVien> listPVThuongTru = db.PhongViens.Where(pv => pv.LoaiPV ==false).ToList();  
            foreach(PhongVien pv in listPVThuongTru)
            {
                excelWs.Range["A1"].Font.Color = Color.Red;
            }
            excelWs.Name = "DanhsachPhongVien";
            
            
            SaveFileDialog sfd = new SaveFileDialog();
            if(sfd.ShowDialog() == DialogResult.OK)
            {
                excelWb.SaveAs(sfd.FileName);
            }
              excelApp.Quit();
        }

        private void txt_SoGioLam_KeyPress(object sender, KeyPressEventArgs e)
        {
           if(!char.IsControl(e.KeyChar)|| !char.IsDigit(e.KeyChar)|| (e.KeyChar!= '.'))
            {
                e.Handled = true;
            }
        }

        private void btn_Search_Click(object sender, EventArgs e)
        {
            string maNvSearch = txt_MaNVSearch.Text.Trim();
            string hoTenSearch = txt_HoTenSearch.Text.Trim();
            double tuGia = Double.Parse(txt_GiaTu.Text);
            double denGia = Double.Parse(txt_GiaDen.Text);
            bool loaiPV = (bool)cb_DSLoaiPV.SelectedItem;

            List<PhongVien> listPVSearch = db.PhongViens.Where( pv => pv.MaPV.Contains(maNvSearch) 
                                                                && pv.HoTen.Contains(hoTenSearch)
                                                                && pv.LoaiPV == loaiPV
                                                                && pv.Luong >= (decimal?)tuGia
                                                                && pv.Luong <= (decimal?)denGia).ToList();
            foreach (PhongVien pv in listPVSearch)
            {
                lv_DSPV.Items.Clear();
                string[] data = { pv.MaPV, pv.HoTen, pv.GioiTinh, pv.NgayVaoLam.ToString() };
                ListViewItem lvi = new ListViewItem(data);
                PV pvModel = new PV(pv);
                if (pvModel.isYellowBackgroud())
                {
                    lvi.BackColor = Color.Yellow;
                }
                lv_DSPV.Items.Add(lvi);
            }
        }
    }
}
