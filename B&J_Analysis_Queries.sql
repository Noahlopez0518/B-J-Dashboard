-- ==========================================
-- B&J Biscuit Sales Dashboard Queries (Top 15)
-- Author: Noah Lopez
-- Description: SQL queries supporting Dashboard 1 & 2 metrics
-- ==========================================

-- 1. Total Revenue, COGS, Profit, Quantity Sold, Profit Margin
SELECT
    SUM(Revenue) AS TotalRevenue,
    SUM(COGS) AS TotalCOGS,
    SUM(Revenue - COGS) AS TotalProfit,
    SUM(Quantity) AS TotalQuantity,
    ROUND(SUM(Revenue - COGS) / SUM(Revenue) * 100, 2) AS ProfitMargin
FROM Sales;

-- 2. Top 5 Customers by Revenue Contribution
SELECT
    CustomerName,
    SUM(Revenue) AS RevenueContribution
FROM Sales
GROUP BY CustomerName
ORDER BY RevenueContribution DESC
LIMIT 5;

-- 3. Revenue Distribution by Product Price Category
SELECT
    CASE 
        WHEN Price < 10 THEN 'Low-Priced'
        ELSE 'High-Priced'
    END AS PriceCategory,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY PriceCategory;

-- 4. Revenue by Age Group and Gender
SELECT
    AgeGroup,
    Gender,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY AgeGroup, Gender
ORDER BY AgeGroup, Gender;

-- 5. Revenue by Payment Method
SELECT
    PaymentMethod,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY PaymentMethod
ORDER BY TotalRevenue DESC;

-- 6. Most Profitable Brand
SELECT
    Brand,
    SUM(Revenue - COGS) AS Profit
FROM Sales
GROUP BY Brand
ORDER BY Profit DESC
LIMIT 1;

-- 7. Most Profitable Location
SELECT
    Location,
    SUM(Revenue - COGS) AS Profit
FROM Sales
GROUP BY Location
ORDER BY Profit DESC
LIMIT 1;

-- 8. Most Profitable Salesperson
SELECT
    Salesperson,
    SUM(Revenue - COGS) AS Profit
FROM Sales
GROUP BY Salesperson
ORDER BY Profit DESC
LIMIT 1;

-- 9. Monthly Revenue Trend (YYYY-MM)
SELECT
    DATE_FORMAT(OrderDate, '%Y-%m') AS YearMonth,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY YearMonth
ORDER BY YearMonth;

-- 10. Quarter-over-Quarter (QoQ) Revenue
SELECT
    CONCAT(YEAR(OrderDate), '-Q', QUARTER(OrderDate)) AS YearQuarter,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY YearQuarter
ORDER BY YearQuarter;

-- 11. Week-over-Week (WoW) Revenue
SELECT
    WEEK(OrderDate) AS WeekNumber,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY WeekNumber
ORDER BY WeekNumber;

-- 12. Weekday vs Weekend Revenue
SELECT
    CASE
        WHEN DAYOFWEEK(OrderDate) IN (1,7) THEN 'Weekend'
        ELSE 'Weekday'
    END AS DayType,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY DayType;

-- 13. Revenue by Location
SELECT
    Location,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY Location
ORDER BY TotalRevenue DESC;

-- 14. KPI Summary by Brand and Location
SELECT
    Brand,
    Location,
    SUM(Quantity) AS TotalQuantity,
    SUM(Revenue) AS TotalRevenue,
    SUM(COGS) AS TotalCOGS,
    SUM(Revenue - COGS) AS TotalProfit,
    ROUND(SUM(Revenue - COGS) / SUM(Revenue) * 100, 2) AS ProfitMargin
FROM Sales
GROUP BY Brand, Location
ORDER BY TotalProfit DESC;

-- 15. Top 3 Revenue-Generating Products
SELECT
    Product,
    SUM(Revenue) AS TotalRevenue
FROM Sales
GROUP BY Product
ORDER BY TotalRevenue DESC
LIMIT 3;
