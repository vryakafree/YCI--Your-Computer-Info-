# 📘 README – CETVN Computer Inventory Script  
*(Tiếng Việt & English)*

---

# 🇻🇳 GIỚI THIỆU (Tiếng Việt)

`CETVN_COMPUTER_LIST` là một script **Batch + PowerShell** dùng để tự động thu thập thông tin phần cứng & hệ thống của máy tính Windows và lưu vào file CSV.  
Script hỗ trợ **ghi đè hoặc cập nhật** theo tên máy tính, đồng thời **kiểm tra file CSV có đang mở hay không** để tránh lỗi ghi dữ liệu.

---

## 🚀 Tính năng

- Kiểm tra file CSV có bị khóa (đang mở trong Excel) hay không.
- Tự động quét và ghi lại các thông tin quan trọng:
  - User đang đăng nhập
  - Tên máy tính
  - Địa chỉ IP thật (ưu tiên Wi‑Fi > LAN)
  - Serial Number
  - Mainboard
  - CPU
  - RAM + Loại RAM + Số khe
  - Ổ cứng (HDD/SSD/NVMe)
  - Card đồ họa (GPU)
  - Phiên bản Windows + Build
  - Thông tin màn hình + kích thước inch
  - Thời gian log gần nhất
- Ghi dữ liệu vào file CSV theo dạng danh sách.
- Tự động đánh số thứ tự (STT).
- Cập nhật dòng cũ nếu máy đã tồn tại trong danh sách.

---

## 📂 File tạo ra

- **CETVN_COMPUTER_LIST.csv**  
  Được lưu tại cùng thư mục với file `.bat`.

---

## 📦 Cách sử dụng

1. Lưu file batch (`.bat`) vào thư mục mong muốn.  
2. Nhấp đôi để chạy script.  
3. Nếu file CSV đang mở, chương trình sẽ yêu cầu bạn đóng lại.  
4. Sau khi hoàn tất, thông báo sẽ xuất hiện:

   ```
   HOAN THANH! Du lieu may <COMPUTERNAME> da duoc cap nhat.
   ```

---

## 📄 Cấu trúc cột CSV

| Cột | Mô tả |
|-----|-------|
| STT | Số thứ tự |
| User | Người dùng đang đăng nhập |
| ComputerName | Tên máy tính |
| IPAddress | Địa chỉ IPv4 |
| SerialNumber | Số seri BIOS |
| Mainboard | Hãng + Model mainboard |
| CPU | Tên CPU |
| RAM | Tổng RAM + loại RAM |
| RAM_Slots | Số khe RAM đã dùng / tổng khe |
| GPU | Card đồ họa |
| OS | Phiên bản Windows |
| Disks | Danh sách ổ cứng |
| Display | Thông tin màn hình |
| Logtimes | Thời điểm ghi log |

---

# 🇺🇸 INTRODUCTION (English)

`CETVN_COMPUTER_LIST` is a **Batch + PowerShell inventory script** used to automatically scan hardware & system information on Windows computers and store the results into a CSV file.  
The script supports **append/update** based on the computer name and checks whether the CSV file is currently open to avoid write errors.

---

## 🚀 Features

- Detects if the CSV file is locked (opened in Excel).
- Collects detailed system information:
  - Logged‑in user
  - Computer name
  - Real IPv4 (Wi‑Fi prioritized)
  - BIOS Serial Number
  - Mainboard information
  - CPU model
  - RAM size + type + slot usage
  - Physical disks (HDD/SSD/NVMe)
  - GPU(s)
  - Windows version + Build
  - Monitor names & size in inches
  - Latest log timestamp
- Automatically writes to a CSV list.
- Keeps sorted order and auto numbering.
- If the computer already exists in the list, the entry is updated instead of duplicated.

---

## 📂 Output File

- **CETVN_COMPUTER_LIST.csv**  
  Saved in the same folder as the `.bat` file.

---

## 📦 Usage

1. Place the `.bat` file in your desired folder.  
2. Double‑click to run.  
3. If the CSV is open, the script will pause and ask you to close it.  
4. Once complete, you will see:

   ```
   COMPLETED! System information for <COMPUTERNAME> has been updated.
   ```

---

## 📄 CSV Columns Overview

| Column | Description |
|--------|-------------|
| STT | Index number |
| User | Logged‑in user |
| ComputerName | Device hostname |
| IPAddress | IPv4 address |
| SerialNumber | BIOS serial |
| Mainboard | Motherboard manufacturer + model |
| CPU | Processor name |
| RAM | Total RAM + RAM type |
| RAM_Slots | Used / total slots |
| GPU | Graphics card |
| OS | Windows version |
| Disks | Disk list |
| Display | Monitor details |
| Logtimes | Timestamp |
