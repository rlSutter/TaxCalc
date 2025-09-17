-- TaxCalc Database Schema DDL Script
-- Generated from analysis of TaxCalc web service code
-- This script creates the database tables and structures used by the TaxCalc tax computation engine

-- =============================================
-- Database Creation
-- =============================================
-- Note: Create databases as needed for your environment
-- CREATE DATABASE [reports];
-- CREATE DATABASE [siebeldb];

-- =============================================
-- Main Order Tables
-- =============================================

-- S_ORDER - Main order table
CREATE TABLE [dbo].[S_ORDER] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [DISCNT_AMT] DECIMAL(18,2) NULL,
    [DISCNT_PERCENT] DECIMAL(5,2) NULL,
    [FRGHT_AMT] DECIMAL(18,2) NULL,
    [ACCNT_ID] NVARCHAR(15) NULL,
    [SHIP_ADDR_ID] NVARCHAR(15) NULL,
    [SHIP_PER_ADDR_ID] NVARCHAR(15) NULL,
    [PR_SHIP_ADDR_ID] NVARCHAR(15) NULL,
    [PR_BL_ADDR_ID] NVARCHAR(15) NULL,
    [PR_ADDR_ID] NVARCHAR(15) NULL,
    [CREATED] DATETIME NULL,
    [STATUS_CD] NVARCHAR(30) NULL,
    [X_REGISTRATION_FLG] NVARCHAR(1) NULL,
    [X_AUDIT_FLG] NVARCHAR(1) NULL,
    [X_INVOICE_NUM] NVARCHAR(30) NULL,
    [REQ_SHIP_DT] DATETIME NULL,
    [X_TAX_AMT] DECIMAL(18,2) NULL,
    [X_TAX_RATE] DECIMAL(5,4) NULL,
    [X_FRGHT_TAX_AMT] DECIMAL(18,2) NULL,
    [X_FRGHT_TAX_RATE] DECIMAL(5,4) NULL
);

-- S_ORDER_ITEM - Order line items
CREATE TABLE [dbo].[S_ORDER_ITEM] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [ORDER_ID] NVARCHAR(15) NOT NULL,
    [CRSE_OFFR_ID] NVARCHAR(15) NULL,
    [UNIT_PRI] DECIMAL(18,2) NULL,
    [BASE_UNIT_PRI] DECIMAL(18,2) NULL,
    [QTY_REQ] DECIMAL(18,2) NULL,
    [DISCNT_AMT] DECIMAL(18,2) NULL,
    [DISCNT_PERCENT] DECIMAL(5,2) NULL,
    [X_TAX_AMT] DECIMAL(18,2) NULL,
    [X_TAX_RATE] DECIMAL(5,4) NULL,
    FOREIGN KEY ([ORDER_ID]) REFERENCES [dbo].[S_ORDER]([ROW_ID])
);

-- S_ORDER_ITEM_X - Order item extensions
CREATE TABLE [dbo].[S_ORDER_ITEM_X] (
    [PAR_ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [ATTRIB_08] NVARCHAR(255) NULL,
    FOREIGN KEY ([PAR_ROW_ID]) REFERENCES [dbo].[S_ORDER_ITEM]([ROW_ID])
);

-- =============================================
-- Organization and Account Tables
-- =============================================

-- S_ORG_EXT - Organization/Account information
CREATE TABLE [dbo].[S_ORG_EXT] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [X_ACCOUNT_NUM] NVARCHAR(30) NULL,
    [BU_ID] NVARCHAR(15) NULL,
    [NAME] NVARCHAR(100) NULL
);

-- =============================================
-- Address Tables
-- =============================================

-- S_ADDR_ORG - Organization addresses
CREATE TABLE [dbo].[S_ADDR_ORG] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [STATE] NVARCHAR(2) NULL,
    [CITY] NVARCHAR(50) NULL,
    [ZIPCODE] NVARCHAR(10) NULL,
    [COUNTY] NVARCHAR(50) NULL
);

-- S_ADDR_PER - Personal addresses
CREATE TABLE [dbo].[S_ADDR_PER] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [STATE] NVARCHAR(2) NULL,
    [CITY] NVARCHAR(50) NULL,
    [ZIPCODE] NVARCHAR(10) NULL,
    [COUNTY] NVARCHAR(50) NULL
);

-- =============================================
-- Course and Product Tables
-- =============================================

-- S_CRSE_OFFR - Course offerings
CREATE TABLE [dbo].[S_CRSE_OFFR] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [X_HELD_ADDRESS_ID] NVARCHAR(15) NULL,
    FOREIGN KEY ([X_HELD_ADDRESS_ID]) REFERENCES [dbo].[S_ADDR_ORG]([ROW_ID])
);

-- S_LST_OF_VAL - List of values (used for taxware configuration)
CREATE TABLE [dbo].[S_LST_OF_VAL] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [TYPE] NVARCHAR(30) NULL,
    [NAME] NVARCHAR(100) NULL,
    [HIGH] NVARCHAR(255) NULL
);

-- =============================================
-- Taxware Specific Tables
-- =============================================

-- TAXWARE_CNTY_FIPS - County FIPS codes for taxware
CREATE TABLE [dbo].[TAXWARE_CNTY_FIPS] (
    [ROW_ID] NVARCHAR(15) NOT NULL PRIMARY KEY,
    [STATE_CODE] NVARCHAR(2) NULL,
    [COUNTY_CODE] NVARCHAR(3) NULL,
    [COUNTY_NAME] NVARCHAR(50) NULL,
    [FIPS_CODE] NVARCHAR(5) NULL
);

-- =============================================
-- Indexes for Performance
-- =============================================

-- Indexes on foreign keys
CREATE INDEX [IX_S_ORDER_ITEM_ORDER_ID] ON [dbo].[S_ORDER_ITEM] ([ORDER_ID]);
CREATE INDEX [IX_S_ORDER_ITEM_CRSE_OFFR_ID] ON [dbo].[S_ORDER_ITEM] ([CRSE_OFFR_ID]);
CREATE INDEX [IX_S_ORDER_ACCNT_ID] ON [dbo].[S_ORDER] ([ACCNT_ID]);
CREATE INDEX [IX_S_ORDER_SHIP_ADDR_ID] ON [dbo].[S_ORDER] ([SHIP_ADDR_ID]);
CREATE INDEX [IX_S_ORDER_SHIP_PER_ADDR_ID] ON [dbo].[S_ORDER] ([SHIP_PER_ADDR_ID]);

-- Indexes on commonly queried fields
CREATE INDEX [IX_S_ORDER_STATUS_CD] ON [dbo].[S_ORDER] ([STATUS_CD]);
CREATE INDEX [IX_S_ORDER_CREATED] ON [dbo].[S_ORDER] ([CREATED]);
CREATE INDEX [IX_S_LST_OF_VAL_TYPE_NAME] ON [dbo].[S_LST_OF_VAL] ([TYPE], [NAME]);
CREATE INDEX [IX_TAXWARE_CNTY_FIPS_STATE_COUNTY] ON [dbo].[TAXWARE_CNTY_FIPS] ([STATE_CODE], [COUNTY_CODE]);

-- =============================================
-- Sample Data Insertion (Optional)
-- =============================================

-- Insert sample taxware configuration
INSERT INTO [dbo].[S_LST_OF_VAL] ([ROW_ID], [TYPE], [NAME], [HIGH]) VALUES
('1-1', 'TAXWARE', 'DEFAULT_TAX_RATE', '0.0875'),
('1-2', 'TAXWARE', 'FREIGHT_TAXABLE', 'Y'),
('1-3', 'TAXWARE', 'DEBUG_MODE', 'N');

-- =============================================
-- Views for Common Queries
-- =============================================

-- View for order tax calculation data
CREATE VIEW [dbo].[V_ORDER_TAX_DATA] AS
SELECT 
    O.ROW_ID,
    O.DISCNT_AMT,
    O.DISCNT_PERCENT,
    O.FRGHT_AMT,
    O.X_TAX_AMT,
    O.X_TAX_RATE,
    O.X_FRGHT_TAX_AMT,
    O.X_FRGHT_TAX_RATE,
    A.STATE,
    A.CITY,
    A.ZIPCODE,
    E.X_ACCOUNT_NUM,
    E.BU_ID,
    E.NAME,
    O.CREATED,
    O.STATUS_CD,
    O.X_REGISTRATION_FLG,
    O.X_AUDIT_FLG,
    O.X_INVOICE_NUM,
    CASE WHEN O.REQ_SHIP_DT IS NULL THEN O.CREATED ELSE O.REQ_SHIP_DT END AS SHIP_DATE
FROM [dbo].[S_ORDER] O
LEFT OUTER JOIN [dbo].[S_ADDR_ORG] A ON A.ROW_ID = O.SHIP_ADDR_ID
LEFT OUTER JOIN [dbo].[S_ORG_EXT] E ON E.ROW_ID = O.ACCNT_ID;

-- View for order item tax calculation data
CREATE VIEW [dbo].[V_ORDER_ITEM_TAX_DATA] AS
SELECT 
    I.ROW_ID,
    I.ORDER_ID,
    I.UNIT_PRI,
    I.BASE_UNIT_PRI,
    I.QTY_REQ,
    I.DISCNT_AMT,
    I.DISCNT_PERCENT,
    I.X_TAX_AMT,
    I.X_TAX_RATE,
    (CASE WHEN I.BASE_UNIT_PRI IS NULL OR I.BASE_UNIT_PRI = 0 THEN I.UNIT_PRI ELSE I.BASE_UNIT_PRI END) AS EFFECTIVE_UNIT_PRICE,
    (CASE WHEN I.BASE_UNIT_PRI IS NULL OR I.BASE_UNIT_PRI = 0 THEN I.UNIT_PRI ELSE I.BASE_UNIT_PRI END - ISNULL(I.DISCNT_AMT, 0)) * I.QTY_REQ AS LINE_TOTAL
FROM [dbo].[S_ORDER_ITEM] I;

-- =============================================
-- Stored Procedures (Optional)
-- =============================================

-- Procedure to get order data for tax calculation
CREATE PROCEDURE [dbo].[sp_GetOrderTaxData]
    @OrderID NVARCHAR(15)
AS
BEGIN
    SET NOCOUNT ON;
    
    SELECT 
        O.ROW_ID,
        O.DISCNT_AMT,
        O.DISCNT_PERCENT,
        O.FRGHT_AMT,
        A.STATE,
        A.CITY,
        A.ZIPCODE,
        E.X_ACCOUNT_NUM,
        E.BU_ID,
        E.NAME,
        O.CREATED,
        O.STATUS_CD,
        O.X_REGISTRATION_FLG,
        O.X_AUDIT_FLG,
        O.X_INVOICE_NUM,
        CASE WHEN O.REQ_SHIP_DT IS NULL THEN O.CREATED ELSE O.REQ_SHIP_DT END AS SHIP_DATE
    FROM [dbo].[S_ORDER] O
    LEFT OUTER JOIN [dbo].[S_ADDR_ORG] A ON A.ROW_ID = O.SHIP_ADDR_ID
    LEFT OUTER JOIN [dbo].[S_ORG_EXT] E ON E.ROW_ID = O.ACCNT_ID
    WHERE O.ROW_ID = @OrderID;
END;

-- Procedure to update order tax amounts
CREATE PROCEDURE [dbo].[sp_UpdateOrderTax]
    @OrderID NVARCHAR(15),
    @TaxAmount DECIMAL(18,2),
    @TaxRate DECIMAL(5,4),
    @FreightTaxAmount DECIMAL(18,2) = NULL,
    @FreightTaxRate DECIMAL(5,4) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE [dbo].[S_ORDER] 
    SET 
        X_TAX_AMT = @TaxAmount,
        X_TAX_RATE = @TaxRate,
        X_FRGHT_TAX_AMT = ISNULL(@FreightTaxAmount, X_FRGHT_TAX_AMT),
        X_FRGHT_TAX_RATE = ISNULL(@FreightTaxRate, X_FRGHT_TAX_RATE)
    WHERE ROW_ID = @OrderID;
END;

-- Procedure to update order item tax amounts
CREATE PROCEDURE [dbo].[sp_UpdateOrderItemTax]
    @ItemID NVARCHAR(15),
    @TaxAmount DECIMAL(18,2),
    @TaxRate DECIMAL(5,4),
    @FreightTaxability NVARCHAR(1) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    UPDATE [dbo].[S_ORDER_ITEM] 
    SET 
        X_TAX_AMT = @TaxAmount,
        X_TAX_RATE = @TaxRate
    WHERE ROW_ID = @ItemID;
    
    IF @FreightTaxability IS NOT NULL
    BEGIN
        UPDATE [dbo].[S_ORDER_ITEM_X] 
        SET ATTRIB_08 = @FreightTaxability
        WHERE PAR_ROW_ID = @ItemID;
    END;
END;
