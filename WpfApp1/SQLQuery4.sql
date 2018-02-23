CREATE PROC insertStockDataNew4
@name varchar(50),
@price float,
@date varchar(50)
AS
	INSERT INTO StockData.dbo.[Stock_WebData](Name,Price,Date)
	VALUES (@name,@price,@date)
