CREATE PROCEDURE [dbo].[fillFixes]
      @listlimp listlimfix READONLY
AS
BEGIN
      INSERT INTO fixes(obj,fix,fixto)
      SELECT Obj,fix,fixto FROM @listlimp
END