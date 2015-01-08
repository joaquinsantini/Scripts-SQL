-- Script que imprime en una impresora perteneciente a la red local cualquier documento .pdf.
-- Es necesario un conversor de archivos .pdf a .ps estilo GhostScript.

DECLARE @PathPDF	VARCHAR(255)
DECLARE @PathConversor	VARCHAR(255)
DECLARE @PathImpresora	VARCHAR(255)
DECLARE @Cmd		NVARCHAR(4000)
DECLARE @Existe		INT
DECLARE @PathPS		VARCHAR(255)
DECLARE @Error		INT
DECLARE @ErrorMessage	VARCHAR(1023)

-- PathPDF: path completo donde se encuentra el archivo a imprimir.
SELECT	@PathPDF = '<PathPDF>'

-- PathConversor: path donde se encuentra el .exe del conversor a .ps.
SELECT	@PathConversor = '<PathConversor>'

-- PathImpresora: path donde se encuentra la impresora. Ejemplo: '\\MiServer\HP LaserJet P2050 Series PCL6'
SELECT	@PathImpresora = '<PathImpresora>'

-- Verifico existencia de archivo en FS
EXECUTE @Error = Master.dbo.xp_fileexist @PathPDF, @Existe OUT

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error al buscar archivo en FileSystem.'

	PRINT	@ErrorMessage

	RETURN
END

IF (@Existe = 0)
BEGIN
	SELECT	@ErrorMessage = 'El archivo que desea imprimir no existe.'

	PRINT	@ErrorMessage

	RETURN
END

SELECT	@PathPS = SUBSTRING(@PathPDF, 1, LEN(@PathPDF) - 3) + 'ps'

-- Asigno comando shell a ejecutar para crear el archivo .ps (los archivos .pdf o .docx no son imprimibles)
SELECT	@Cmd = '"' + @PathConversor + '" ' + '"' + @PathPDF + '"' + ' ' + '"' + @PathPS + '"'

EXECUTE @Error = Master.dbo.xp_cmdshell @Cmd, NO_OUTPUT

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error al enviar a imprimir el archivo.'

	PRINT	@ErrorMessage

	RETURN
END

-- Asigno comando shell a ejecutar para imprimir el archivo .ps
SELECT	@Cmd = 'print /d:"' + @PathImpresora + '" ' + '"' + @PathPS + '"'

EXECUTE @Error = Master.dbo.xp_cmdshell @Cmd, NO_OUTPUT

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error al enviar a imprimir el archivo.'

	PRINT	@ErrorMessage

	RETURN
END

-- Elimino el archivo .ps creado anteriormente
SELECT	@Cmd = 'del "' + @PathPS + '"'

EXECUTE @Error = Master.dbo.xp_cmdshell @Cmd, NO_OUTPUT

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error al eliminar archivo .ps.'

	PRINT	@ErrorMessage

	RETURN
END


RETURN