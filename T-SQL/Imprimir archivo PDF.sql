-- Script que imprime en una impresora perteneciente a la red local cualquier documento .pdf
-- Es necesario un conversor de archivos .pdf a .ps estilo GhostScript

-- Variables a utilizar
DECLARE @PathPDF	VARCHAR(255),
	@PathConversor	VARCHAR(255),
	@PathImpresora	VARCHAR(255),
	@Cmd		NVARCHAR(4000),
	@Existe		INT,
	@PathPS		VARCHAR(255),
	@Error		INT,
	@ErrorMessage	VARCHAR(1022);


-- PathPDF: path completo donde se encuentra el archivo a imprimir
SELECT	@PathPDF = '<PathPDF>';


-- PathConversor: path donde se encuentra el .exe del conversor a .ps
SELECT	@PathConversor = '<PathConversor>';


-- PathImpresora: path donde se encuentra la impresora. Ejemplo: '\\MiServer\HP LaserJet P2050 Series PCL6'
SELECT	@PathImpresora = '<PathImpresora>';

-------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- Verifico que los parámetros fueron seteados
IF (@PathPDF IS NULL OR LEN(@PathPDF) = 0)
BEGIN
	PRINT	'El parámetro @PathPDF no fue seteado.';

	RETURN;
END

IF (@PathConversor IS NULL OR LEN(@PathConversor) = 0)
BEGIN
	PRINT	'El parámetro @PathConversor no fue seteado.';

	RETURN;
END

IF (@PathImpresora IS NULL OR LEN(@PathImpresora) = 0)
BEGIN
	PRINT	'El parámetro @PathImpresora no fue seteado.';

	RETURN;
END


-- Verifico existencia de archivo en FS
EXECUTE @Error = xp_fileexist @PathPDF, @Existe OUT;

IF (@Error <> 0)
BEGIN
	PRINT	'Error al buscar archivo en FileSystem.';

	RETURN;
END

IF (@Existe = 0)
BEGIN
	PRINT	'El archivo que desea imprimir no existe.';

	RETURN;
END

SELECT	@PathPS = SUBSTRING(@PathPDF, 1, LEN(@PathPDF) - 3) + 'ps';


-- Asigno comando shell a ejecutar para crear el archivo .ps (los archivos .pdf o .docx no son imprimibles)
SELECT	@Cmd = '"' + @PathConversor + '" ' + '"' + @PathPDF + '"' + ' ' + '"' + @PathPS + '"';

EXECUTE @Error = xp_cmdshell @Cmd, NO_OUTPUT;

IF (@Error <> 0)
BEGIN
	PRINT	'Error al enviar a imprimir el archivo.';

	RETURN;
END


-- Asigno comando shell a ejecutar para imprimir el archivo .ps
SELECT	@Cmd = 'print /d:"' + @PathImpresora + '" ' + '"' + @PathPS + '"';

EXECUTE @Error = master.dbo.xp_cmdshell @Cmd, NO_OUTPUT;

IF (@Error <> 0)
BEGIN
	PRINT	'Error al enviar a imprimir el archivo.';

	RETURN;
END


-- Elimino el archivo .ps creado anteriormente
SELECT	@Cmd = 'del "' + @PathPS + '"';

EXECUTE @Error = xp_cmdshell @Cmd, NO_OUTPUT;

IF (@Error <> 0)
BEGIN
	PRINT	'Error al eliminar archivo .ps.';

	RETURN;
END


RETURN;