# Quy tắc Quản lý Git (Git Rules)

## Nghiêm cấm Commit và Push code tự động

1. **KHÔNG TỰ ĐỘNG COMMIT**:
   - Tuyệt đối KHÔNG tự ý chạy lệnh `git commit` (hoặc `git commit -m ...`) sau khi chỉnh sửa code hoặc hoàn thành tác vụ.
   - Chỉ thực hiện `git commit` khi và chỉ khi có yêu cầu cụ thể, rõ ràng từ người dùng (ví dụ: "hãy commit thay đổi này").

2. **KHÔNG TỰ ĐỘNG PUSH**:
   - Tuyệt đối KHÔNG bao giờ tự ý chạy lệnh `git push` (hoặc `git push origin ...`) lên remote repository.
   - Việc đẩy mã nguồn lên remote phải do người dùng tự thực hiện hoặc người dùng ra lệnh cụ thể.

3. **QUY TRÌNH TIÊU CHUẨN**:
   - Khi hoàn thành yêu cầu: Chỉ thực hiện chỉnh sửa file, chạy build/bundle, kiểm tra cú pháp và test tính đúng đắn.
   - Báo cáo kết quả thay đổi cho người dùng và để người dùng tự quyết định việc commit và push code.
