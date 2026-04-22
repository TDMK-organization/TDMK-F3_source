USE [master];
GO

-- Kiểm tra nếu DB chưa tồn tại thì mới tạo
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'OK2SHIP_SMT')
BEGIN
    CREATE DATABASE [OK2SHIP_SMT];
END
GO