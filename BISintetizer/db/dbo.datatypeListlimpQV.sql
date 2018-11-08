CREATE TYPE [dbo].[listlimpqv] AS TABLE(
      [customer] [varchar](100) NULL,
      [product] [varchar](50) NULL,
	  [date] [date] NULL,
      [indicator] [varchar](50) NULL,
	  [value] [int] NULL
)
GO