-- Script que desde un Excel guardado como "Texto separado por tabulaciones" exporta un script para hacer inserts en la tabla deseada.
-- El archivo Excel debe tener:
--	a. En la primer fila los nombres de las columnas distintas de NULL. Las columnas nulleables no se especifican.
--	b. Como última columna deben estar todos los registros numerados empezando por la primer fila. Es decir, la primer fila (la que tiene el nombre de las columnas)
--	debe tener el número 1, la fila de abajo el número 2, y así sucesivamente.
--	c. Al exportar el archivo Excel a .txt, sacarle el último "Enter". Esto es, el cursor debe quedar "titilando" después de la última numeración de la última fila.

USE [Master]
GO

SET NOCOUNT ON

-- Variables a utilizar
DECLARE @Ambiente		VARCHAR(255)
DECLARE @BBDD			VARCHAR(255)
DECLARE @Tabla			VARCHAR(255)
DECLARE @TablaLog		VARCHAR(255)
DECLARE @CantidadColumnas	INT
DECLARE @Error			INT
DECLARE @ErrorMessage		VARCHAR(1023)
DECLARE @PathExportacion	VARCHAR(255)
DECLARE @PathTxt		VARCHAR(255)
DECLARE @Existe			INT
DECLARE @Query			NVARCHAR(4000)
DECLARE @QueryParams		NVARCHAR(4000)

DECLARE @TMP_Existe TABLE (
	ExisteArchivo		INT,
	ExisteDirectorio	INT,
	ExisteDirectorioPadre	INT
)

-- Ambiente: linked server deseado.
SELECT	@Ambiente = '<Ambiente>'

-- BBDD: base de datos donde se encuentra la tabla.
SELECT	@BBDD = '<Base>'

-- Tabla: tabla donde se quieren hacer los inserts.
SELECT	@Tabla = '<Tabla>'

-- TablaLog: tabla de log donde se realizan los inserts. Si no existe tabla de Log asignar NULL.
SELECT	@TablaLog = '<TablaLog>'

-- PathExportacion: path donde se va a exportar el archivo .sql.
SELECT	@PathExportacion = '<Path>'

-- PathTxt: path donde se encuentra el archivo txt a leer.
SELECT	@PathTxt = '<PathTxt>'

-- CantidadColumnas: cantidad de columnas distintas de NULL que contiene el archivo txt.
SELECT	@CantidadColumnas = 5

----------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- Verifico existencia de archivo txt en FS.
INSERT INTO @TMP_Existe (ExisteArchivo, ExisteDirectorio, ExisteDirectorioPadre)
	EXECUTE @Error = Master.dbo.xp_fileexist @PathTxt

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error al buscar archivo txt en FileSystem.'

	PRINT	@ErrorMessage

	RETURN
END

IF ((SELECT ExisteArchivo FROM @TMP_Existe) = 0)
BEGIN
	SELECT	@ErrorMessage = 'El archivo txt que desea importar no existe.'

	PRINT	@ErrorMessage

	RETURN
END

DELETE @TMP_Existe

-- Verifico existencia del path de exportación en FS.
INSERT INTO @TMP_Existe (ExisteArchivo, ExisteDirectorio, ExisteDirectorioPadre)
	EXECUTE @Error = Master.dbo.xp_fileexist @PathExportacion

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error al verificar path de exportación en FileSystem.'

	PRINT	@ErrorMessage

	RETURN
END

IF ((SELECT ExisteDirectorio FROM @TMP_Existe) = 0)
BEGIN
	SELECT	@ErrorMessage = 'El path de exportación no es válido.'

	PRINT	@ErrorMessage

	RETURN
END

-- Verifico que se haya ingresado una base de datos.
IF (@BBDD IS NULL)
BEGIN
	SELECT	@ErrorMessage = 'No indicó la base de datos donde se encuentra la tabla.'

	PRINT	@ErrorMessage

	RETURN
END

-- Verifico que se haya ingresado una tabla.
IF (@Tabla IS NULL)
BEGIN
	SELECT	@ErrorMessage = 'No indicó la tabla donde se van a realizar los insert.'

	PRINT	@ErrorMessage

	RETURN
END


----------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- Tabla donde se van a almacenar las líneas del script a exportar.
CREATE TABLE Script (
	Linea		INT IDENTITY NOT NULL PRIMARY KEY,
	Texto		VARCHAR(7900) NOT NULL
)

-- Inserto header del script.
INSERT INTO Script(Texto) VALUES('USE [' + @BBDD + ']')
INSERT INTO Script(Texto) VALUES('')
INSERT INTO Script(Texto) VALUES('SET LANGUAGE ''us_english''')
INSERT INTO Script(Texto) VALUES('SET NOCOUNT ON')
INSERT INTO Script(Texto) VALUES('GO')
INSERT INTO Script(Texto) VALUES('')

-- Verifico que se haya ingresado una tabla de log.
IF (@TablaLog IS NOT NULL)
BEGIN
	INSERT INTO Script(Texto) VALUES('DECLARE @UsuarioLog	VARCHAR(15)')
	INSERT INTO Script(Texto) VALUES('DECLARE @SistemaLog	VARCHAR(15)')
	INSERT INTO Script(Texto) VALUES('')
	INSERT INTO Script(Texto) VALUES('SELECT	@UsuarioLog = ''MECANUS''')
	INSERT INTO Script(Texto) VALUES('SELECT	@SistemaLog = ''SISTEMA'' --Reemplazar por el que corresponda.')
	INSERT INTO Script(Texto) VALUES('')
END

-- Inserto inicio de transacción.
INSERT INTO Script(Texto) VALUES('BEGIN TRANSACTION Insert_Tabla')
INSERT INTO Script(Texto) VALUES('')

DECLARE @Contador INT

SELECT	@Contador = 0

-- En @Query almaceno la query de creación de una tabla temporal dependiendo de la cantidad de columnas que tenga el archivo .txt. Debe coincidir con
-- lo ingresado en la variable @CantidadColumnas.
SELECT	@Query = 'CREATE TABLE ##TMP_Registros ( '

-- Loope para buscar las columnas.
WHILE (@Contador <> @CantidadColumnas)
BEGIN
	SELECT	@Contador = @Contador + 1

	SELECT	@Query = @Query + ' Columna' + CONVERT(VARCHAR(10), @Contador) + '	VARCHAR(7900)'

	IF (@Contador <> @CantidadColumnas)
	BEGIN
		SELECT	@Query = @Query + ', '
	END ELSE BEGIN
		SELECT	@Query = @Query + ', RegistroId VARCHAR(3) )'
	END
END

EXECUTE @Error = sp_executesql @Query

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error creando tabla temporal ##TMP_Registros.'

	PRINT	@ErrorMessage

	DROP TABLE Script

	RETURN
END

-- En @Query almaceno el bulk insert del archivo .txt en la tabla temporal ##TMP_Registros
SELECT @Query =	'BULK INSERT ##TMP_Registros
			FROM ' + '''' + REPLACE(@PathTxt, '''', '''''') + '''
			WITH (	CODEPAGE = ''RAW'',
				DATAFILETYPE = ''char'',
				FIELDTERMINATOR = ''\t'',
				ROWTERMINATOR = ''\n'')'

EXECUTE @Error = sp_executesql @Query

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error en bulk insert.'

	PRINT	@ErrorMessage

	DROP TABLE Script

	DROP TABLE ##TMP_Registros

	RETURN
END

DECLARE @Columnas VARCHAR(7900)

SELECT	@Columnas = '('

SELECT	@Contador = 0

-- Loop para determinar la cadena de values.
WHILE (@Contador <> @CantidadColumnas)
BEGIN
	SELECT	@Contador = @Contador + 1

	SELECT	@Query = 'SELECT @Columnas = @Columnas + (SELECT Columna' + CONVERT(VARCHAR(10), @Contador) + ' FROM ##TMP_Registros WHERE RegistroId = 1)'

	SELECT	@QueryParams = '@Columnas VARCHAR(7900) OUTPUT'

	EXECUTE @Error = sp_executesql @Query, @QueryParams, @Columnas = @Columnas OUTPUT

	IF (@Error <> 0)
	BEGIN
		SELECT	@ErrorMessage = 'Error obteniendo las columnas de la tabla (query dinámica).'

		PRINT	@ErrorMessage

		DROP TABLE Script

		DROP TABLE ##TMP_Registros

		RETURN
	END

	IF (@Contador <> @CantidadColumnas)
	BEGIN
		SELECT	@Columnas = @Columnas + ', '
	END ELSE BEGIN
		SELECT	@Columnas = @Columnas + ')'
	END
END

-- Variables de uso dinámico.
DECLARE @CantidadRegistros	INT
DECLARE @ContadorAux		INT
DECLARE @Values			VARCHAR(7900)
DECLARE @ColumnasLog		VARCHAR(7900)
DECLARE @ValuesLog		VARCHAR(7900)
DECLARE @ColumnaValor		VARCHAR(7900)
DECLARE @ColumnaNombre		VARCHAR(7900)

SELECT	@CantidadRegistros = COUNT(*) - 1 FROM ##TMP_Registros

SELECT	@Contador = 0

-- Tabla temporal donde se va a almacenar el tipo de columna para determinar el insert.
CREATE TABLE ##TMP_Columnas (
	Table_Cat		VARCHAR(255),
	Table_Schem		VARCHAR(255),
	Table_Name		VARCHAR(255),
	Column_Name		VARCHAR(255),
	Data_Type		INT,
	Type_Name		VARCHAR(255),
	Column_Size		INT,
	Buffer_Length		INT,
	Decimal_Digits		INT,
	Num_Prec_Radix		INT,
	Nullable		INT,
	Remarks			VARCHAR(255),
	Column_Def		VARCHAR(255),
	Sql_Data_Type		INT,
	Sql_Datetime_Sub	INT,
	Char_Octet_Length	INT,
	Ordinal_Position	INT,
	Is_Nullable		VARCHAR(255),
	Ss_Data_Type		INT
)

-- Loop principal donde se recorren los registros del archivo .txt.
WHILE (@Contador <> @CantidadRegistros)
BEGIN
	SELECT	@Contador = @Contador + 1

	SELECT	@ContadorAux = 0

	SELECT	@Values = '('

	-- Loop para determinar variables @Columnas y @Values.
	WHILE (@ContadorAux <> @CantidadColumnas)
	BEGIN
		SELECT	@ContadorAux = @ContadorAux + 1

		SELECT	@Query = 'SELECT @ColumnaValor = Columna' + CONVERT(VARCHAR(10), @ContadorAux) + ' FROM ##TMP_Registros WHERE RegistroId = @Contador + 1'

		SELECT	@QueryParams = '@Contador INT, @ColumnaValor VARCHAR(7900) OUTPUT'

		EXECUTE @Error = sp_executesql @Query, @QueryParams,
					@Contador	= @Contador,
					@ColumnaValor	= @ColumnaValor OUTPUT

		IF (@Error <> 0)
		BEGIN
			SELECT	@ErrorMessage = 'Error obteniendo ColumnaValor (query dinámica).'

			PRINT	@ErrorMessage

			DROP TABLE Script

			DROP TABLE ##TMP_Registros

			DROP TABLE ##TMP_Columnas

			RETURN
		END

		SELECT	@Query = '	SELECT	@ColumnaNombre = Columna' + CONVERT(VARCHAR(10), @ContadorAux) + ' FROM ##TMP_Registros WHERE RegistroId = 1

					DELETE ##TMP_Columnas

					INSERT INTO ##TMP_Columnas
						EXECUTE sp_columns_ex
							@Table_Server	= @Ambiente,
							@Table_Name	= @Tabla,
							@Table_Schema	= ''dbo'',
							@Table_Catalog	= @BBDD,
							@Column_Name	= @ColumnaNombre

					IF ((SELECT Type_Name FROM ##TMP_Columnas) = ''int'')
					BEGIN
						SELECT	@Values = @Values + (SELECT Columna' + CONVERT(VARCHAR(10), @ContadorAux) + ' FROM ##TMP_Registros WHERE RegistroId = @Contador + 1)
					END ELSE BEGIN
						SELECT	@Values = @Values + '''''''' + ''' + @ColumnaValor + ''' + ''''''''
					END'

		SELECT	@QueryParams = '@Contador INT, @Ambiente VARCHAR(255), @Tabla VARCHAR(255), @BBDD VARCHAR(255), @ColumnaNombre VARCHAR(7900) OUTPUT, @Values VARCHAR(7900) OUTPUT'

		EXECUTE @Error = sp_executesql @Query, @QueryParams,
					@Contador	= @Contador,
					@Ambiente	= @Ambiente,
					@Tabla		= @Tabla,
					@BBDD		= @BBDD,
					@ColumnaNombre	= @ColumnaNombre OUTPUT,
					@Values		= @Values OUTPUT

		IF (@Error <> 0)
		BEGIN
			SELECT	@ErrorMessage = 'Error obteniendo los registros de la tabla (query dinámica).'

			PRINT	@ErrorMessage

			DROP TABLE Script

			DROP TABLE ##TMP_Registros

			DROP TABLE ##TMP_Columnas

			RETURN
		END

		IF (@ContadorAux <> @CantidadColumnas)
		BEGIN
			SELECT	@Values = @Values + ', '
		END ELSE BEGIN
			SELECT	@Values = @Values + ')'
		END
	END

	-- Inserto registro.
	INSERT INTO Script(Texto) VALUES('------------------------ Registro ' + CONVERT(VARCHAR(10), @Contador) + ' ------------------------')
	INSERT INTO Script(Texto) VALUES('INSERT INTO ' + @Tabla + ' ' + @Columnas)
	INSERT INTO Script(Texto) VALUES('	VALUES ' + @Values)
	INSERT INTO Script(Texto) VALUES('')
	INSERT INTO Script(Texto) VALUES('IF (@@ERROR <> 0)')
	INSERT INTO Script(Texto) VALUES('BEGIN')
	INSERT INTO Script(Texto) VALUES('	ROLLBACK TRANSACTION Insert_Tabla')
	INSERT INTO Script(Texto) VALUES('')
	INSERT INTO Script(Texto) VALUES('	PRINT ''Error insertando registro ' + CONVERT(VARCHAR(10), @Contador) + ' en tabla.''')
	INSERT INTO Script(Texto) VALUES('')
	INSERT INTO Script(Texto) VALUES('	RETURN')
	INSERT INTO Script(Texto) VALUES('END')
	INSERT INTO Script(Texto) VALUES('')

	IF (@TablaLog IS NOT NULL)
	BEGIN
		SELECT	@ColumnasLog = SUBSTRING(@Columnas, 1, LEN(@Columnas) - 1)

		SELECT	@ColumnasLog = @ColumnasLog + ', FechaLog, SistemaLog, UsuarioLog)'

		SELECT	@ValuesLog = SUBSTRING(@Values, 1, LEN(@Values) - 1)

		SELECT	@ValuesLog = @ValuesLog + ', GETDATE(), @SistemaLog, @UsuarioLog)'

		INSERT INTO Script(Texto) VALUES('INSERT INTO ' + @TablaLog + ' ' + @ColumnasLog)
		INSERT INTO Script(Texto) VALUES('	VALUES ' + @ValuesLog)
		INSERT INTO Script(Texto) VALUES('')
		INSERT INTO Script(Texto) VALUES('')
	END ELSE BEGIN
		INSERT INTO Script(Texto) VALUES('')
	END
END

-- Inserto commit de transacción.
INSERT INTO Script(Texto) VALUES('COMMIT TRANSACTION Insert_Tabla')
INSERT INTO Script(Texto) VALUES('')
INSERT INTO Script(Texto) VALUES('')
INSERT INTO Script(Texto) VALUES('RETURN')

-- Generación del script.
IF (RIGHT(@PathExportacion, 1) <> '\')
BEGIN
	SELECT	@PathExportacion = @PathExportacion + '\'
END

DECLARE @PathScript VARCHAR(1023)

SELECT	@PathScript = @PathExportacion + 'INSERT - ' + @Tabla + '.sql'

DECLARE @Cmd	NVARCHAR(4000)

SELECT	@Cmd = 'type nul > "' + @PathScript + '"'

EXECUTE @Error = Master.dbo.xp_CmdShell @Cmd, NO_OUTPUT

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error creando archivo .sql.'

	PRINT	@ErrorMessage

	DROP TABLE Script

	DROP TABLE ##TMP_Registros

	DROP TABLE ##TMP_Columnas

	RETURN
END

-- Copiado masivo de datos a archivo .sql
SELECT	@Cmd = 'bcp "SELECT NULLIF(Texto, '''') FROM Master.dbo.Script ORDER BY Linea" queryout "' + @PathScript + '" -c -T'

EXECUTE @Error = xp_CmdShell @Cmd, NO_OUTPUT

IF (@Error <> 0)
BEGIN
	SELECT	@ErrorMessage = 'Error copiando contenido de tabla en archivo .sql.'

	PRINT	@ErrorMessage

	DROP TABLE Script

	DROP TABLE ##TMP_Registros

	DROP TABLE ##TMP_Columnas

	RETURN
END

-- Elimino tablas usadas.
DROP TABLE Script

DROP TABLE ##TMP_Registros

DROP TABLE ##TMP_Columnas

-- Printeo el path donde se generó el script con éxito.
PRINT	'Script generado en ' + @PathScript


RETURN