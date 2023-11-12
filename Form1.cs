using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Xml.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using NPOI.SS.Formula.Functions;
using System.Data.OleDb;
using ExcelDataReader;
using ClosedXML.Excel;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;
using System.Windows.Forms;



namespace QuanLiSinhViennn
{
    public partial class Form1 : Form
    {
        private string MyConnection = @"Data Source=DESKTOP-4PBV4IL\SQLEXPRESS2012;Initial Catalog=qlisinhvien;Integrated Security=True";

        private int index;
        private string path;

        public Form1()
        {
            InitializeComponent();
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            lammoi();

        }
        public void lammoi()
        {
            using (SqlConnection connection = new SqlConnection(MyConnection))
            {
                try
                {
                    connection.Open();

                    string query = "SELECT masv, tensv, tuoi, gioitinh, hinhanh FROM sinhvien";

                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    // Thêm cột mới để lưu dữ liệu hình ảnh dưới dạng byte[]
                    dt.Columns.Add("Pic", Type.GetType("System.Byte[]"));
                    foreach (DataRow row in dt.Rows)
                    {
                        // Đọc dữ liệu hình ảnh từ đường dẫn và gán vào cột "Pic"
                        string imagePath = row["hinhanh"].ToString();

                        // Add a check to ensure the file path is not empty or null
                        if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                        {
                            row["Pic"] = File.ReadAllBytes(imagePath);
                        }
                        else
                        {
                            row["Pic"] = DBNull.Value; // or handle it according to your needs
                        }
                    }

                    dataGridView1.DataSource = dt;
                    connection.Close();

                    dataGridView1.ClearSelection();
                    if (index == dataGridView1.Rows.Count - 1)
                    {
                        index = dataGridView1.Rows.Count - 2;
                        indamdong();
                    }
                    else
                    {
                        indamdong();
                    }

                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }



        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void btadd_Click(object sender, EventArgs e)
        {
            try
            {
                string mssv = txtmssv.Text;
                string name = txtten.Text;
                string tuoiText = txttuoi.Text;
                bool gioiTinh = radioButtonNam.Checked;
                if (string.IsNullOrWhiteSpace(mssv) || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(tuoiText))
                {
                    MessageBox.Show("Vui lòng điền đầy đủ thông tin.");
                    return;
                }

                if (IsMssvExists(mssv))
                {
                    MessageBox.Show("Mã số sinh viên đã tồn tại. Vui lòng chọn một mã số khác.");
                    return;
                }
                if (!radioButtonNam.Checked && !radioButtonNu.Checked)
                {
                    MessageBox.Show("Bạn chưa chọn giới tính.");
                    return; // Thoát khỏi phương thức hoặc xử lý thêm lỗi khác ở đây (tuỳ vào yêu cầu).
                }
                using (SqlConnection connection = new SqlConnection(MyConnection))
                {
                    connection.Open();

                    string query = "INSERT INTO sinhvien (masv, tensv, tuoi, gioitinh, hinhanh) VALUES (@mssv, @name, @ngaysinh, @gioitinh, @hinhanh)";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@name", name);
                        command.Parameters.AddWithValue("@ngaysinh", Convert.ToDateTime(tuoiText));
                        command.Parameters.AddWithValue("@gioitinh", gioiTinh);
                        command.Parameters.AddWithValue("@mssv", mssv);
                        command.Parameters.AddWithValue("@hinhanh", path);

                        int result = command.ExecuteNonQuery();
                        if (result > 0)
                        {
                            MessageBox.Show("Thêm sinh viên " + mssv + " thành công.");
                        }
                        else
                        {
                            MessageBox.Show("Thêm sinh viên thất bại.");
                        }
                    }
                }
                lammoi();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }



        private bool IsMssvExists(string mssv)
        {
            using (SqlConnection connection = new SqlConnection(MyConnection))
            {
                connection.Open();

                string query = "SELECT COUNT(*) FROM sinhvien WHERE masv = @mssv";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@mssv", mssv);

                    int count = (int)command.ExecuteScalar();

                    return count > 0;
                }
            }
        }


        private void btchon_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opnfd = new OpenFileDialog();
                opnfd.Filter = "Image Files (*.jpg;*.jpeg;*.gif;*.png)|*.jpg;*.jpeg;*.gif;*.png";

                if (opnfd.ShowDialog() == DialogResult.OK)
                {
                    // Kiểm tra kích thước tối đa cho hình ảnh (ví dụ: giới hạn kích thước 300x300)
                    int maxWidth = 500;
                    int maxHeight = 500;

                    System.Drawing.Image selectedImage = System.Drawing.Image.FromFile(opnfd.FileName);

                    // Kiểm tra kích thước của hình ảnh
                    if (selectedImage.Width <= maxWidth && selectedImage.Height <= maxHeight)
                    {
                        pictureBox1.Image = new Bitmap(selectedImage);
                        path = opnfd.FileName;
                    }
                    else
                    {
                        MessageBox.Show("Hình ảnh vượt quá kích thước tối đa (" + maxWidth + "x" + maxHeight + "). Vui lòng chọn hình ảnh khác.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "Có vẻ như thiếu hình rồi");
            }
        }

        private void btLamMoi_Click(object sender, EventArgs e)
        {

            lammoi();
            txtmssv.Text = String.Empty;
            txtmssv.Enabled = true;
            txtten.Text = String.Empty;
            txttuoi.Text = String.Empty;
            pictureBox1.Image = null;
        }

        private void btThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn thoát khỏi ứng dụng?", "Xác nhận thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Người dùng đã xác nhận thoát, đóng ứng dụng
                Close();
            }
            else
            {
                // Người dùng đã hủy việc thoát, không làm gì cả
            }
        }


        private void btXoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count > 0 && index >= 0)
                {
                    // Kiểm tra nếu danh sách không rỗng và index hợp lệ thì thực hiện xóa
                    DataGridViewRow row = dataGridView1.Rows[index];
                    string mssvToDelete = row.Cells[0].Value.ToString().Trim();

                    // Hiển thị hộp thoại Yes/No để xác nhận việc xóa
                    DialogResult dialogResult = MessageBox.Show($"Bạn có chắc chắn muốn xóa sinh viên có MSSV {mssvToDelete} không?", "Xác nhận xóa", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.Yes)
                    {
                        using (SqlConnection connection = new SqlConnection(MyConnection))
                        {
                            connection.Open();

                            string query = "DELETE FROM sinhvien WHERE masv = @mssv";

                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@mssv", mssvToDelete);
                                int result = command.ExecuteNonQuery();
                                if (result > 0)
                                {
                                    MessageBox.Show($"Xóa thông tin sinh viên {mssvToDelete} thành công.");
                                    lammoi();
                                    if (dataGridView1.Rows.Count == 0)
                                    {
                                        int? index = null;
                                    }

                                    // Không cần tăng index vì dòng đã bị xóa
                                }
                                else
                                {
                                    MessageBox.Show($"Xóa thông tin sinh viên {mssvToDelete} thất bại.");
                                    lammoi();
                                }
                            }
                        }
                    }
                    // Ngược lại, không làm gì cả nếu người dùng chọn No
                }
                else
                {
                    MessageBox.Show("Không có dữ liệu hoặc bạn chưa chọn một sinh viên để xóa.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void btSua_Click(object sender, EventArgs e)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txtmssv.Text) || string.IsNullOrWhiteSpace(txtten.Text) || string.IsNullOrWhiteSpace(txttuoi.Text) || pictureBox1.Image == null)
                    {
                        MessageBox.Show("Vui lòng điền đầy đủ thông tin.");
                        return;
                    }
                    if (index < 0)
                    {
                        MessageBox.Show("Vui lòng chọn một sinh viên để sửa.");
                        return;
                    }
                    if (!radioButtonNam.Checked && !radioButtonNu.Checked)
                    {
                        MessageBox.Show("Bạn chưa chọn giới tính.");
                        return; // Thoát khỏi phương thức hoặc xử lý thêm lỗi khác ở đây (tuỳ vào yêu cầu).
                    }
                    string mssv = txtmssv.Text;
                    string name = txtten.Text;
                    string tuoiText = txttuoi.Text;
                    bool gioiTinh = radioButtonNam.Checked;

                    // Kiểm tra xem Mssv đã thay đổi so với giá trị ban đầu
                    string mssv2 = dataGridView1.Rows[index].Cells[0].Value.ToString().Trim();
                    if (!mssv.Equals(mssv2))
                    {
                        MessageBox.Show("Không được thay đổi MSSV");
                        return;
                    }

                    using (SqlConnection connection = new SqlConnection(MyConnection))
                    {
                        connection.Open();

                    string query = "UPDATE sinhvien SET masv=@mssv , tensv = @name, tuoi = @ngaysinh, gioitinh = @gioitinh, hinhanh = @hinhanh WHERE masv = @mssv2";

                    string hinhanhBytes = path ?? dataGridView1.Rows[index].Cells["hinhanh"].Value.ToString();
                

                    using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@name", name);
                            command.Parameters.AddWithValue("@ngaysinh", Convert.ToDateTime(tuoiText));
                            command.Parameters.AddWithValue("@gioitinh", gioiTinh);
                            command.Parameters.AddWithValue("@mssv", mssv);
                            command.Parameters.AddWithValue("@hinhanh", hinhanhBytes);
                            command.Parameters.AddWithValue("@mssv2", mssv2); // MSSV cũ

                            int result = command.ExecuteNonQuery();
                            if (result > 0)
                            {
                                MessageBox.Show("Sửa thông tin sinh viên " + mssv + " thành công.");
                            }
                            else
                            {
                                MessageBox.Show(result.ToString());
                            }
                        }
                    }
                    lammoi();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.Message);
                }
            }


        public void indamdong()
        {
            DataGridViewRow selectedRow = dataGridView1.Rows[index];
            selectedRow.Selected = true;
            selectedRow.DefaultCellStyle.SelectionBackColor = Color.Aquamarine; // Màu nền khi chọn
            selectedRow.DefaultCellStyle.SelectionForeColor = Color.Black; // Màu chữ khi chọn
        }



        public void clickchuot(DataGridViewCellEventArgs e)
        {
            try
            {
                index = e.RowIndex;
                DataGridViewRow row = dataGridView1.Rows[index];
                //txtmssv.Enabled = false;
                txtmssv.Text = row.Cells[0].Value.ToString().Trim();
                txtten.Text = row.Cells[1].Value.ToString().Trim();
                txttuoi.Text = row.Cells[2].Value.ToString().Trim();

                // Kiểm tra giới tính
                bool checkgioitinh = (bool)row.Cells[3].Value;
                if (checkgioitinh)
                {
                    radioButtonNam.Checked = true;
                }
                else
                {
                    radioButtonNu.Checked = true;
                }

                // Hiển thị hình ảnh
                byte[] imgData = (byte[])row.Cells["Pic"].Value;
                if (imgData != null && imgData.Length > 0)
                {
                    MemoryStream ms = new MemoryStream(imgData);
                    pictureBox1.Image = System.Drawing.Image.FromStream(ms);
                }
                else
                {
                    pictureBox1.Image = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\ndưới đó chưa có dữ liệu đâu");
            }

        }


        public void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            index = e.RowIndex;
            if (index >= 0)
            {
                // Bạn muốn thay đổi màu của hàng, hãy đặt màu cho tất cả ô trong hàng
                indamdong();
            }
            clickchuot(e);
        }
        public void ExportToPdf(DataGridView dataGridView, string filePath)
        {
            try
            {
                Document document = new Document();
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
                document.Open();

                PdfPTable pdfTable = new PdfPTable(dataGridView.Columns.Count);
                pdfTable.DefaultCell.Padding = 3;
                pdfTable.WidthPercentage = 100;
                pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                // Tiêu đề của các cột
                foreach (DataGridViewColumn column in dataGridView.Columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                    cell.BackgroundColor = new BaseColor(240, 240, 240);
                    pdfTable.AddCell(cell);
                }

                // Dữ liệu từ DataGridView
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        DataGridViewCell cell = row.Cells[i];

                        if (i != 4) // Kiểm tra nếu đây không phải là cột thêm vào từ hàm lammoi()
                        {
                            pdfTable.AddCell(cell.Value?.ToString() ?? string.Empty);
                        }
                        else
                        {
                            if (cell.Value != null)
                            {
                                string imagePath = cell.Value.ToString();
                                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                                {
                                    iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(imagePath);
                                    pdfTable.AddCell(pdfImage);
                                }
                            }
                        }
                    }
                }

                // Thêm bảng vào tệp PDF
                document.Add(pdfTable);

                // Thêm hình ảnh từ PictureBox (nếu có)
                document.Close();
                writer.Close();

                MessageBox.Show("Xuất tệp PDF thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void btexportpdf_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "PDF Files|*.pdf";
                saveFileDialog.Title = "Chọn vị trí và tên tệp PDF";
                saveFileDialog.FileName = "SinhVien.pdf";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string pdfFilePath = saveFileDialog.FileName;

                    Document document = new Document();
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(pdfFilePath, FileMode.Create));
                    document.Open();

                    PdfPTable pdfTable = new PdfPTable(dataGridView1.Columns.Count);
                    pdfTable.DefaultCell.Padding = 3;
                    pdfTable.WidthPercentage = 100;
                    pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                    // Tiêu đề của các cột
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                        cell.BackgroundColor = new BaseColor(240, 240, 240);
                        pdfTable.AddCell(cell);
                    }

                    // Dữ liệu từ DataGridView
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.ColumnIndex != 4) // Kiểm tra nếu đây là cột chứa hình ảnh
                            {
                                pdfTable.AddCell(cell.Value?.ToString() ?? string.Empty);
                            }
                            else
                            {
                                if (cell.Value != null)
                                {
                                    if (cell.Value is byte[] imgData)
                                    {
                                        MemoryStream ms = new MemoryStream(imgData);
                                        System.Drawing.Image image = System.Drawing.Image.FromStream(ms);

                                        if (image != null)
                                        {
                                            iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(image, ImageFormat.Jpeg);
                                            pdfTable.AddCell(pdfImage);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Thêm bảng vào tệp PDF
                    document.Add(pdfTable);
                    // Thêm hình ảnh từ PictureBox (nếu có)
                    document.Close();
                    writer.Close();

                    MessageBox.Show("Xuất tệp PDF thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void btexportexcel_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();

                // Thiết lập đường dẫn mặc định
                saveFileDialog.InitialDirectory = "G:\\";

                // Thiết lập các thuộc tính của SaveFileDialog
                saveFileDialog.Filter = "Excel Files|*.xlsx";
                saveFileDialog.Title = "Chọn vị trí và tên tệp Excel";
                saveFileDialog.FileName = "SinhVien.xlsx"; // Tên mặc định

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    // Tạo một tệp Excel mới
                    XSSFWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("Sheet1");

                    // Bắt đầu ghi dữ liệu vào sheet
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        IRow row = sheet.CreateRow(i);
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            ICell cell = row.CreateCell(j);

                            object cellValue = dataGridView1.Rows[i].Cells[j].Value;
                            if (cellValue != null)
                            {
                                if (cellValue is byte[] imageData) // Kiểm tra xem dữ liệu có phải là hình ảnh không
                                {
                                    // Ghi hình ảnh vào tệp Excel
                                    var pictureIdx = workbook.AddPicture(imageData, PictureType.JPEG);
                                    var drawing = sheet.CreateDrawingPatriarch();
                                    var anchor = new XSSFClientAnchor(0, 0, 0, 0, j, i, j + 1, i + 1);
                                    var picture = drawing.CreatePicture(anchor, pictureIdx);
                                }
                                else
                                {
                                    cell.SetCellValue(cellValue.ToString());
                                }
                            }
                            else
                            {
                                cell.SetCellValue(string.Empty);
                            }
                        }
                    }

                    // Điều chỉnh kích thước cột cho phù hợp với nội dung
                    for (int columnIndex = 0; columnIndex < dataGridView1.Columns.Count; columnIndex++)
                    {
                        sheet.AutoSizeColumn(columnIndex);
                    }

                    // Lưu tệp Excel tại vị trí đã chọn
                    using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fileStream);
                    }

                    // Đóng workbook sau khi hoàn thành
                    workbook.Close();

                    // Thông báo thành công
                    MessageBox.Show("Xuất dữ liệu thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                // Thông báo lỗi nếu có lỗi xảy ra
                MessageBox.Show("Có lỗi xảy ra: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
