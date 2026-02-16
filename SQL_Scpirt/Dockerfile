FROM mcr.microsoft.com/mssql/server:2025-latest

# Chuyển sang root để cấu hình quyền
USER root

WORKDIR /usr/src/app

# Copy toàn bộ nội dung thư mục SQL_Scpirt vào container
COPY . /usr/src/app

# Sửa lỗi định dạng xuống dòng (CRLF sang LF) và cấp quyền thực thi
RUN sed -i 's/\r$//' /usr/src/app/entrypoint.sh && \
    chmod +x /usr/src/app/entrypoint.sh


USER root

# Sửa lỗi xuống dòng và cấp quyền thực thi
RUN sed -i 's/\r$//' /usr/src/app/entrypoint.sh && \
    chmod +x /usr/src/app/entrypoint.sh

# QUAN TRỌNG: Trả lại quyền cho user mssql trước khi chạy CMD
USER mssql

# Chạy script khởi động
CMD ["/bin/bash", "./entrypoint.sh"]