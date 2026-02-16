#!/bin/bash

# 1. Khởi động SQL Server ở chế độ background
/opt/mssql/bin/sqlservr &
pid=$!

echo "Đang chờ SQL Server khởi động..."

# 2. Vòng lặp kiểm tra: Chờ đến khi SQL Server sẵn sàng nhận kết nối
for i in {1..60}; do
    # Thử chạy lệnh SELECT 1. Nếu thành công ($? -eq 0) thì thoát vòng lặp.
    /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$MSSQL_SA_PASSWORD" -C -Q "SELECT 1" > /dev/null 2>&1
    if [ $? -eq 0 ]; then
        echo "SQL Server đã sẵn sàng sau $i lần thử!"
        
        # 3. Chạy các file .sql trong thư mục scripts
        echo "Bắt đầu chạy các script khởi tạo..."
        for file in /usr/src/app/scripts/*.sql; do
            if [ -f "$file" ]; then
                echo "Đang thực thi: $file"
                /opt/mssql-tools18/bin/sqlcmd -S localhost -U sa -P "$MSSQL_SA_PASSWORD" -d master -i "$file" -C
            fi
        done
        echo "Khởi tạo hoàn tất!"
        break
    fi
    echo "Tiếp tục chờ..."
    sleep 2
done

# 4. Đưa tiến trình SQL Server về foreground để container không bị tắt
wait $pid