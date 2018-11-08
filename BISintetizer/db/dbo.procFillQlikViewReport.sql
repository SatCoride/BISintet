CREATE PROCEDURE [dbo].[fillQliqViewReport]
      @listlimp listlimpqv READONLY
AS
BEGIN
      INSERT INTO qlickViewReport(customer,product,date,indicator,value)
      SELECT customer,product,date,indicator,value FROM @listlimp
END