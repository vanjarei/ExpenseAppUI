-- Step 1: Create and use the correct database
CREATE DATABASE ExpenseDB;
GO

USE ExpenseDB;
GO

-- Step 2: Create Categories table
CREATE TABLE Categories (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    CategoryName NVARCHAR(50) NOT NULL UNIQUE
);
GO

-- Step 3: Insert sample data
INSERT INTO Categories (CategoryName) VALUES
('Food'), ('Travel'), ('Bills'), ('Shopping'), ('Other');
GO

-- Step 4: Create Expenses table
CREATE TABLE Expenses (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    ExpenseDate DATE NOT NULL,
    Category NVARCHAR(50) NOT NULL,
    Description NVARCHAR(255),
    Amount DECIMAL(10, 2) NOT NULL
);
GO
