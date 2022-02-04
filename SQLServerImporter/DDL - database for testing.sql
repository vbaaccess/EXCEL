/*
   20220204_2355 - Create
   Test database structure for an xls file project importing data from the database
*/
USE [master]
GO

CREATE DATABASE [DB_XLS]
GO

USE [DB_XLS]
GO

 CREATE TABLE [dbo].[tDataTest]
([Id] [INT] IDENTITY(1,1) NOT NULL
,[Imie] [nvarchar](30)
,[Nazwisko] [nvarchar](81)
,[Paszport] [nvarchar](100)
,[Kod] [nvarchar](100)
,[DokumentSeriaNumer] [nvarchar](50)
,[Element] INT
,[OkresFrom] DateTime
,[OkresTo] DateTime
,[IdUmowy] INT
,CONSTRAINT [PK_tDataTest] PRIMARY KEY CLUSTERED ([Id] ASC ) ON [PRIMARY]
)
GO 

CREATE VIEW vForXLS AS
SELECT Id,Imie,Nazwisko,Paszport,Kod,DokumentSeriaNumer,OkresFrom,OkresTo,IdUmowy
FROM tDataTest
GO

INSERT INTO tDataTest (Imie,Nazwisko,Paszport,Kod,DokumentSeriaNumer,OkresFrom,OkresTo,IdUmowy) 
VALUES('Pablo', 'Pocasso','PP1234','Firma1','DSN1234',GETDATE()-300,NULL,123)
, ('Marco', 'Gajos','PP4321','Firma1','DSN867',GETDATE()-900,NULL,1125)
, ('Rafaello', 'Rafalo','PP3366','Firma2','DSN005',GETDATE()-125,GETDATE()-20,127)
, ('Pitero', 'Sajlezjo','PS7878','Firma3','DSN015',GETDATE()-325,NULL,128)
