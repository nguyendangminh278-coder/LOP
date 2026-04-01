TSA CLASS MANAGER - HOÀN CHỈNH

Bộ file này gồm:
1. Code.gs          -> backend Apps Script
2. appsscript.json  -> manifest Apps Script
3. index.html       -> frontend local chạy bằng VS Code / Live Server

THÔNG TIN ĐÃ CẬP NHẬT
- Spreadsheet ID: 1w0JZfHJAAQdDy8OfvOxTz4pkPh9I-BdB0d-Mq1rt8lQ
- Sheet name: baocao
- Web App URL: https://script.google.com/macros/s/AKfycbz4xhQu480PBAMwrRbmBLIHrfnZTr6jlcMu5TMEsL4t3TkIbfwhvGj4bjwZY_E8cD7Zeg/exec

CÁCH DÙNG
A. Apps Script
1. Mở Google Sheet đích.
2. Vào Extensions > Apps Script.
3. Dán Code.gs vào file Code.gs.
4. Nếu muốn, bật appsscript.json rồi dán manifest.
5. Save.
6. Deploy > Manage deployments > chỉnh web app hiện tại hoặc deploy mới.
7. Execute as: Me.
8. Who has access: Anyone.

B. Frontend local
1. Mở file index.html bằng VS Code.
2. Chạy bằng Live Server.
3. Form sẽ dùng JSONP để gọi Apps Script, nên không cần nhúng HTML vào Apps Script.

GHI CHÚ
- Ca học có 20h00, 21h30 và Tự nhập.
- QLL có Đăng Minh và Đăng Khoa.
- Ngày mặc định là hôm qua theo giờ Việt Nam.
- Dữ liệu được ghi vào sheet 'baocao' theo cột:
  Date | Lớp học | Môn học | Thầy/cô | Tổng | Sĩ số lớp | Ca học | QLL | Link record buổi học (youtube) | Buổi số (các em ghi đủ số buổi nhé )
