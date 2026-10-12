# Tra cứu hộ gia đình BHXH bán tự động

Công cụ tra cứu hàng loạt trên trang
[Tra cứu hộ gia đình BHXH](https://baohiemxahoi.gov.vn/tracuu/Pages/tra-cuu-ho-gia-dinh.aspx).

Trang có **mã xác nhận (ảnh)** bắt buộc cho mỗi lượt tra, nên công cụ chạy bán tự động:
script tự điền form cho từng người trong file Excel, bạn chỉ cần gõ mã xác nhận,
kết quả được ghi ra file Excel.

## Cài đặt (một lần)

1. Cài [Python 3.9+](https://www.python.org/downloads/) (Windows: tích chọn "Add Python to PATH").
2. Mở cửa sổ dòng lệnh trong thư mục này và chạy:

```
pip install -r requirements.txt
playwright install chromium
```

(Nếu máy đã có Google Chrome, script sẽ dùng Chrome luôn.)

## Chuẩn bị danh sách

Sao chép `mau_danh_sach.xlsx` và điền mỗi người một dòng. Hàng đầu là tiêu đề:

| Cột | Bắt buộc | Ghi chú |
|---|---|---|
| Tỉnh/TP | Có | Ghi đúng tên trên trang, có thể bỏ "Tỉnh"/"Thành phố" (vd: `Hà Nội`) |
| Quận/Huyện, Phường/Xã, Thôn/Xóm | Không | Ghi đúng tên như trên trang |
| Họ tên | Có | Có dấu hoặc không dấu đều được |
| CCCD | Có | Nên định dạng ô là Text để giữ số 0 đầu |
| Ngày sinh | Không | dd/mm/yyyy hoặc năm sinh |
| Mã BHXH | Không | |

## Chạy

```
python tra_cuu_bhxh.py danh_sach.xlsx
```

Trình duyệt sẽ mở ra và tự điền người đầu tiên. Nhìn ảnh mã xác nhận trên trình duyệt rồi gõ mã vào cửa sổ dòng lệnh:

- gõ mã rồi Enter: tra cứu
- `r`: đổi ảnh mã khác
- `s`: bỏ qua người này
- `q`: dừng (kết quả đã tra vẫn được giữ)

Sau khi thấy kết quả, Enter để lưu và sang người tiếp, hoặc `l` để tra lại (vd: gõ sai mã).

Tùy chọn: `-o ket_qua.xlsx` để đặt tên file kết quả, `--bat-dau 20` để chạy tiếp từ người thứ 20.

## Kết quả

File `ket_qua_<ngày giờ>.xlsx` gồm:

- **KetQua**: mỗi người một dòng, gồm thông tin đã tra, trạng thái, thông báo của trang và kết quả dạng văn bản.
- **ChiTiet**: từng dòng của các bảng kết quả trang trả về.

Nếu script không tìm thấy một ô nào đó trên trang, nó sẽ báo trong cửa sổ dòng lệnh; bạn có thể điền tay ô đó trên trình duyệt rồi gõ mã xác nhận như bình thường.
