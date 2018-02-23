CREATE PROC insertNewColumns
@transStartRow int,
@accountNumberPos varchar(50),
@dateColumn varchar(50),
@priceColumn varchar(50),
@balanceColumn varchar(50),
@commentColumn varchar(50)
AS
	INSERT INTO ImportFileData.dbo.[StoredColumns](TransStartRow,AccountNumberPos,DateColumn,PriceColumn,BalanceColumn,CommentColumn)
	VALUES (@transStartRow,@accountNumberPos,@dateColumn,@priceColumn,@balanceColumn,@commentColumn)
