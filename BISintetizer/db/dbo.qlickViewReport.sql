CREATE TABLE [dbo].[qlickViewReport] (
    [Id]        INT          IDENTITY (1, 1) NOT NULL,
    [customer]  VARCHAR (50) NULL,
    [product]   VARCHAR (20) NULL,
    [date]      DATE         NULL,
    [indicator] VARCHAR (20) NULL,
    [value]     VARCHAR (20) NULL,
    PRIMARY KEY CLUSTERED ([Id] ASC)
);

