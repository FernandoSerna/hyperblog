/****** Objeto: tabla [dbo].[TB_TEMP_NOTA_CREDITO]    fecha de la secuencia de comandos: 07/11/2007 09:45:44 a.m. ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_TEMP_NOTA_CREDITO]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_TEMP_NOTA_CREDITO]
GO

/****** Objeto: tabla [dbo].[TB_TEMP_NOTA_CREDITO]    fecha de la secuencia de comandos: 07/11/2007 09:45:45 a.m. ******/
CREATE TABLE [dbo].[TB_TEMP_NOTA_CREDITO] (
	[INTE_TEM_CONSECUTIVO] [int] NULL ,
	[VCHA_EMP_EMPRESA_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_UOR_UNIDAD_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[INTE_CAR_NUMERO] [int] NULL ,
	[VCHA_SER_SERIE_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[DTIM_CAR_FECHA] [datetime] NULL ,
	[VCHA_ART_ARTICULO_ID] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_ART_NOMBRE_ESPAÑOL] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[FLOA_TEM_CANTIDAD] [float] NULL ,
	[FLOA_TEM_PRECIO] [float] NULL ,
	[FLOA_TEM_DESCUENTO_1] [float] NULL ,
	[FLOA_TEM_DESCUENTO_2] [float] NULL ,
	[FLOA_TEM_DESCUENTO_3] [float] NULL ,
	[FLOA_CAR_IMPORTE_NETO] [float] NULL ,
	[VCHA_CAR_IMPORTE_LETRA] [varchar] (200) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_CLI_CLAVE_ID] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_CLI_NOMBRE] [varchar] (1000) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_CLI_DIRECCION] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_CLI_CP] [varchar] (1000) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_COL_NOMBRE] [varchar] (1000) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_MUN_NOMBRE] [varchar] (1000) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_EST_NOMBRE] [varchar] (1000) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_CIU_NOMBRE] [varchar] (1000) COLLATE Modern_Spanish_CI_AS NULL ,
	[VCHA_CLI_RFC] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[FLOA_CAR_TIPO_CAMBIO] [float] NULL ,
	[FLOA_CDE_IVA] [float] NULL ,
	[FLOA_DVE_TIPO_CAMBIO] [float] NULL ,
	[INTE_TEM_RENGLON] [int] NULL ,
	[INTE_TEM_NUMERO] [int] NULL ,
	[VCHA_TEM_NOMBRE] [varchar] (1000) COLLATE Modern_Spanish_CI_AS NULL ,
	[INTE_TEM_FACTURA] [int] NULL ,
	[FLOA_TEM_DESCUENTO_OTORGADO] [float] NULL ,
	[FLOA_TEM_DESCUENTO_APLICADO] [float] NULL ,
	[FLOA_TEM_IMPORTE] [float] NULL 
) ON [PRIMARY]
GO

