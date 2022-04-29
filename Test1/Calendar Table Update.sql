/****** Script for SelectTopNRows command from SSMS  ******/

Declare @Cal_Table_Addon as Table (

CalDate date,
CalWeekStart date,
CalWeekEnd date )

insert into @Cal_Table_Addon (CalDate,CalWeekStart, CalWeekEnd) 

SELECT [CalDate]
       ,CASE
		when CalDayOfWeek = 1 then CalDate
		else DATEADD(DAY,(ABS(CalDayofweek) * -1)+1,CalDate)
	End as 'CalWeekStart'
	,CASE when CalDayOfWeek = 7 then CalDate
		else DATEADD(DAY,7 - CalDayOfWeek,CalDate)
	end as 'CalWeekEnd'
  FROM [Goals].[dbo].[calendar_tbl]

  Select * from @Cal_Table_Addon

  /*

  UPDATE
    Sales_Import
SET
    Sales_Import.AccountNumber = RAN.AccountNumber
FROM
    Sales_Import SI
INNER JOIN
    RetrieveAccountNumber RAN
ON 
    SI.LeadID = RAN.LeadID;


  */