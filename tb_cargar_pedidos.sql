if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TB_CARGAR_PEDIDOS]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TB_CARGAR_PEDIDOS]
GO

CREATE TABLE [dbo].[TB_CARGAR_PEDIDOS] (
	[VCHA_CLI_CLAVE_ID] [varchar] (50) COLLATE Traditional_Spanish_CI_AS NULL ,
	[VCHA_ESB_ESTABLECIMIENTO_ID] [varchar] (50) COLLATE Traditional_Spanish_CI_AS NULL ,
	[VCHA_ART_ARTICULO_ID] [varchar] (50) COLLATE Traditional_Spanish_CI_AS NULL ,
	[FLOA_PED_CANTIDAD] [float] NULL 
) ON [PRIMARY]
GO

