-- Script que desde un Excel guardado como "Texto separado por tabulaciones" exporta un script para hacer inserts en la tabla deseada.
-- El archivo Excel debe tener en la primer fila los nombres de las columnas distintas de NULL. Las columnas nulleables no se especifican.

USE [master]
GO

SET NOCOUNT ON;


-- Variables a utilizar
DECLARE @BBDD			VARCHAR(255),
	@Owner			VARCHAR(255),
	@Tabla			VARCHAR(255),
	@TablaLog		VARCHAR(255),
	@CantidadColumnas	INT,
	@Error			INT,
	@ErrorMessage		VARCHAR(1023),
	@PathExportacion	VARCHAR(255),
	@PathTxt		VARCHAR(255),
	@Existe			INT,
	@Query			NVARCHAR(4000),
	@QueryParams		NVARCHAR(4000);

DECLARE @TMP_Existe TABLE (
	ExisteArchivo		INT,
	ExisteDirectorio	INT,
	ExisteDirectorioPadre	INT
);

CREATE TABLE Script (
	Linea		INT IDENTITY NOT NULL PRIMARY KEY,
	Texto		VARCHAR(7900) NOT NULL
);


-- @BBDD: base de datos donde se encuentra la tabla
SELECT	@BBDD = '<Base>';


-- @Owner: owner donde se encuentra la tabla. El defecto es dbo.
SELECT	@Owner = 'dbo';


-- @Tabla: tabla donde se quieren hacer los inserts
SELECT	@Tabla = '<Tabla>';


-- @TablaLog: tabla de log donde se realizan los inserts. Si no existe tabla de Log asignar NULL
SELECT	@TablaLog = NULL;


-- @PathExportacion: carpeta donde se va a exportar el archivo .sql. Debe ser una ruta compartida
SELECT	@PathExportacion = '<Carpeta>';


-- @PathTxt: path completo del archivo .txt a leer. Debe estar en una ruta compartida
SELECT	@PathTxt = '<Archivo>';


-- @CantidadColumnas: cantidad de columnas distintas de NULL que contiene el archivo txt
SELECT	@CantidadColumnas = 0;

----------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- Verifico existencia de archivo txt en FS
INSERT INTO @TMP_Existe (ExisteArchivo, ExisteDirectorio, ExisteDirectorioPadre)
	EXECUTE @Error = master.dbo.xp_fileexist @PathTxt;

IF (@Error <> 0)
BEGIN
	PRINT	'Error al buscar archivo txt en FileSystem.';

	DROP TABLE Script;

	RETURN;
END

IF ((SELECT ExisteArchivo FROM @TMP_Existe) = 0)
BEGIN
	PRINT	'El archivo txt que desea importar no existe.';

	DROP TABLE Script;

	RETURN;
END

DELETE @TMP_Existe;

-- Verifico existencia del path de exportación en FS
INSERT INTO @TMP_Existe (ExisteArchivo, ExisteDirectorio, ExisteDirectorioPadre)
	EXECUTE @Error = master.dbo.xp_fileexist @PathExportacion;

IF (@Error <> 0)
BEGIN
	PRINT	'Error al verificar path de exportación en FileSystem.';

	DROP TABLE Script;

	RETURN;
END

IF ((SELECT ExisteDirectorio FROM @TMP_Existe) = 0)
BEGIN
	PRINT	'El path de exportación no es válido.';

	DROP TABLE Script;

	RETURN;
END

-- Verifico que se haya ingresado una base de datos
IF (@BBDD IS NULL OR LEN(@BBDD) = 0)
BEGIN
	PRINT	'No indicó la base de datos donde se encuentra la tabla.';

	DROP TABLE Script;

	RETURN;
END

-- Verifico que se haya ingresado un owner
IF (@Owner IS NULL OR LEN(@Owner) = 0)
BEGIN
	PRINT	'No indicó el owner donde se encuentra la tabla.';

	DROP TABLE Script;

	RETURN;
END

-- Verifico que se haya ingresado una tabla
IF (@Tabla IS NULL OR LEN(@Tabla) = 0)
BEGIN
	PRINT	'No indicó la tabla donde se van a realizar los insert.';

	DROP TABLE Script;

	RETURN;
END

-- Verifico que la tabla se encuentre en la base de datos
SELECT	@Query = '	IF (EXISTS(SELECT 1 FROM <<BBDD>>.<<OWNER>>.sysobjects WHERE Type = ''U'' AND Name = @Tabla))
			BEGIN
				SELECT	@Existe = 1;
			END ELSE BEGIN
				SELECT	@Existe = 0;
			END';

SELECT	@Query = REPLACE(REPLACE(@Query, '<<BBDD>>', @BBDD), '<<OWNER>>', @Owner);

SELECT	@QueryParams = '@Tabla VARCHAR(255), @Existe INT OUTPUT';

EXECUTE @Error = sp_executesql @Query, @QueryParams, @Tabla = @Tabla, @Existe = @Existe OUTPUT;

IF (@Error <> 0)
BEGIN
	PRINT	'Error verificando existencia de tabla en base de datos.';

	DROP TABLE Script;

	RETURN;
END

IF (@Existe = 0)
BEGIN
	PRINT	'La tabla especificada no existe en la base de datos especificada.';

	DROP TABLE Script;

	RETURN;
END

-- Verifico la cantidad de columnas ingresadas
IF (@CantidadColumnas <= 0)
BEGIN
	PRINT	'La cantidad de columnas debe ser mayor a cero.';

	DROP TABLE Script;

	RETURN;
END


-- Inserto header del script
INSERT INTO Script(Texto) VALUES('USE [' + @BBDD + '];');
INSERT INTO Script(Texto) VALUES('');
INSERT INTO Script(Texto) VALUES('SET LANGUAGE ''us_english'';');
INSERT INTO Script(Texto) VALUES('SET NOCOUNT ON;');
INSERT INTO Script(Texto) VALUES('GO');
INSERT INTO Script(Texto) VALUES('');
INSERT INTO Script(Texto) VALUES('DECLARE @Error	INT;');
INSERT INTO Script(Texto) VALUES('');


-- Verifico que se haya ingresado una tabla de log
IF (@TablaLog IS NOT NULL)
BEGIN
	INSERT INTO Script(Texto) VALUES('DECLARE @UsuarioLog	VARCHAR(15);');
	INSERT INTO Script(Texto) VALUES('DECLARE @SistemaLog	VARCHAR(15);');
	INSERT INTO Script(Texto) VALUES('');
	INSERT INTO Script(Texto) VALUES('SELECT	@UsuarioLog = ''<<USUARIOLOG>>''; --Reemplazar por el que corresponda.');
	INSERT INTO Script(Texto) VALUES('SELECT	@SistemaLog = ''<<SISTEMALOG>>''; --Reemplazar por el que corresponda.');
	INSERT INTO Script(Texto) VALUES('');
END


-- Inserto inicio de transacción
INSERT INTO Script(Texto) VALUES('BEGIN TRANSACTION Insert_Tabla;');
INSERT INTO Script(Texto) VALUES('');

DECLARE @Contador INT;

SELECT	@Contador = 0;


-- En @Query almaceno la query de creación de una tabla temporal dependiendo de la cantidad de columnas que tenga el archivo .txt. Debe coincidir con
-- lo ingresado en la variable @CantidadColumnas
SELECT	@Query = 'CREATE TABLE ##TMP_Registros ( ';


-- LOOP para buscar las columnas
WHILE (@Contador <> @CantidadColumnas)
BEGIN
	SELECT	@Contador = @Contador + 1;

	SELECT	@Query = @Query + ' Columna' + CONVERT(VARCHAR(10), @Contador) + '	VARCHAR(7900),';
END

SELECT	@Query = SUBSTRING(@Query, 1, LEN(@Query) - 1) + ');';

EXECUTE @Error = sp_executesql @Query;

IF (@Error <> 0)
BEGIN
	PRINT	'Error creando tabla temporal ##TMP_Registros.';

	DROP TABLE Script;

	RETURN;
END


-- En @Query almaceno el bulk insert del archivo .txt en la tabla temporal ##TMP_Registros
SELECT	@Query = '	BULK INSERT ##TMP_Registros
				FROM ' + '''' + REPLACE(@PathTxt, '''', '''''') + '''
				WITH (	CODEPAGE = ''RAW'',
					DATAFILETYPE = ''widechar'',
					FIELDTERMINATOR = ''\t'',
					ROWTERMINATOR = ''\n'')';

EXECUTE @Error = sp_executesql @Query;

IF (@Error <> 0)
BEGIN
	PRINT	'Error en bulk insert.';

	DROP TABLE Script;

	DROP TABLE ##TMP_Registros;

	RETURN;
END

DECLARE @Columnas VARCHAR(7900);

SELECT	@Columnas = '(';

SELECT	@Contador = 0;


-- Agrego la columna RegistroId a la tabla temporal ##TMP_Registros
ALTER TABLE ##TMP_Registros
	ADD RegistroId INT IDENTITY(1, 1);


-- LOOP para determinar la cadena de values
WHILE (@Contador <> @CantidadColumnas)
BEGIN
	SELECT	@Contador = @Contador + 1;

	SELECT	@Query = 'SELECT @Columnas = @Columnas + (SELECT Columna' + CONVERT(VARCHAR(10), @Contador) + ' FROM ##TMP_Registros WHERE RegistroId = 1)';

	SELECT	@QueryParams = '@Columnas VARCHAR(7900) OUTPUT';

	EXECUTE @Error = sp_executesql @Query, @QueryParams, @Columnas = @Columnas OUTPUT;

	IF (@Error <> 0)
	BEGIN
		PRINT	'Error obteniendo las columnas de la tabla (query dinámica).';

		DROP TABLE Script;

		DROP TABLE ##TMP_Registros;

		RETURN;
	END

	IF (@Contador <> @CantidadColumnas)
	BEGIN
		SELECT	@Columnas = @Columnas + ', ';
	END ELSE BEGIN
		SELECT	@Columnas = @Columnas + ')';
	END
END


-- Variables de uso dinámico
DECLARE @CantidadRegistros	INT,
	@ContadorAux		INT,
	@Values			VARCHAR(7900),
	@ColumnasLog		VARCHAR(7900),
	@ValuesLog		VARCHAR(7900),
	@ColumnaValor		VARCHAR(7900),
	@ColumnaNombre		VARCHAR(7900),
	@ColumnaTipo		VARCHAR(255);

SELECT	@CantidadRegistros = COUNT(*) - 1 FROM ##TMP_Registros;

SELECT	@Contador = 0;


-- LOOP principal donde se recorren los registros del archivo .txt
WHILE (@Contador <> @CantidadRegistros)
BEGIN
	SELECT	@Contador = @Contador + 1;

	SELECT	@ContadorAux = 0;

	SELECT	@Values = '(';

	-- LOOP para determinar variables @Columnas y @Values
	WHILE (@ContadorAux <> @CantidadColumnas)
	BEGIN
		SELECT	@ContadorAux = @ContadorAux + 1;

		SELECT	@Query = 'SELECT @ColumnaValor = Columna' + CONVERT(VARCHAR(10), @ContadorAux) + ' FROM ##TMP_Registros WHERE RegistroId = @Contador + 1';

		SELECT	@QueryParams = '@Contador INT, @ColumnaValor VARCHAR(7900) OUTPUT';

		EXECUTE @Error = sp_executesql @Query, @QueryParams,
					@Contador	= @Contador,
					@ColumnaValor	= @ColumnaValor OUTPUT;

		IF (@Error <> 0)
		BEGIN
			PRINT	'Error obteniendo ColumnaValor (query dinámica).';

			DROP TABLE Script;

			DROP TABLE ##TMP_Registros;

			DROP TABLE ##TMP_Columnas;

			RETURN;
		END

		SELECT	@Query = '	SELECT	@ColumnaNombre = Columna' + CONVERT(VARCHAR(10), @ContadorAux) + ' FROM ##TMP_Registros WHERE RegistroId = 1;

					SELECT	@ColumnaTipo = data_type FROM <<BBDD>>.information_schema.columns WHERE table_schema = @Owner AND table_name = @Tabla AND column_name = @ColumnaNombre;

					IF (@ColumnaTipo = ''int'')
					BEGIN
						SELECT	@Values = @Values + (SELECT Columna' + CONVERT(VARCHAR(10), @ContadorAux) + ' FROM ##TMP_Registros WHERE RegistroId = @Contador + 1);
					END ELSE BEGIN
						SELECT	@Values = @Values + '''''''' + ''' + @ColumnaValor + ''' + '''''''';
					END';

		SELECT	@Query = REPLACE(@Query, '<<BBDD>>', @BBDD);

		SELECT	@QueryParams = '@Contador INT, @Tabla VARCHAR(255), @Owner VARCHAR(255), @ColumnaTipo VARCHAR(255) OUTPUT, @ColumnaNombre VARCHAR(7900) OUTPUT, @Values VARCHAR(7900) OUTPUT';

		EXECUTE @Error = sp_executesql @Query, @QueryParams,
					@Contador	= @Contador,
					@Tabla		= @Tabla,
					@Owner		= @Owner,
					@ColumnaTipo	= @ColumnaTipo OUTPUT,
					@ColumnaNombre	= @ColumnaNombre OUTPUT,
					@Values		= @Values OUTPUT;

		IF (@Error <> 0)
		BEGIN
			PRINT	'Error obteniendo los registros de la tabla (query dinámica).';

			DROP TABLE Script;

			DROP TABLE ##TMP_Registros;

			RETURN;
		END

		IF (@ContadorAux <> @CantidadColumnas)
		BEGIN
			SELECT	@Values = @Values + ', ';
		END ELSE BEGIN
			SELECT	@Values = @Values + ')';
		END
	END

	-- Inserto registro
	INSERT INTO Script(Texto) VALUES('------------------------ Registro ' + CONVERT(VARCHAR(10), @Contador) + ' ------------------------');
	INSERT INTO Script(Texto) VALUES('INSERT INTO ' + @Tabla + ' ' + @Columnas);
	INSERT INTO Script(Texto) VALUES('	VALUES ' + @Values + ';');
	INSERT INTO Script(Texto) VALUES('');
	INSERT INTO Script(Texto) VALUES('SELECT	@Error = @@ERROR;');
	INSERT INTO Script(Texto) VALUES('');
	INSERT INTO Script(Texto) VALUES('IF (@Error <> 0)');
	INSERT INTO Script(Texto) VALUES('BEGIN');
	INSERT INTO Script(Texto) VALUES('	ROLLBACK TRANSACTION Insert_Tabla;');
	INSERT INTO Script(Texto) VALUES('');
	INSERT INTO Script(Texto) VALUES('	PRINT ''Error insertando registro ' + CONVERT(VARCHAR(10), @Contador) + ' en tabla.'';');
	INSERT INTO Script(Texto) VALUES('');
	INSERT INTO Script(Texto) VALUES('	RETURN;');
	INSERT INTO Script(Texto) VALUES('END');
	INSERT INTO Script(Texto) VALUES('');

	IF (@TablaLog IS NOT NULL)
	BEGIN
		SELECT	@ColumnasLog = SUBSTRING(@Columnas, 1, LEN(@Columnas) - 1);

		SELECT	@ColumnasLog = @ColumnasLog + ', FechaLog, SistemaLog, UsuarioLog)';

		SELECT	@ValuesLog = SUBSTRING(@Values, 1, LEN(@Values) - 1);

		SELECT	@ValuesLog = @ValuesLog + ', GETDATE(), @SistemaLog, @UsuarioLog);'

		INSERT INTO Script(Texto) VALUES('INSERT INTO ' + @TablaLog + ' ' + @ColumnasLog);
		INSERT INTO Script(Texto) VALUES('	VALUES ' + @ValuesLog + ';');
		INSERT INTO Script(Texto) VALUES('');
		INSERT INTO Script(Texto) VALUES('SELECT	@Error = @@ERROR;');
		INSERT INTO Script(Texto) VALUES('');
		INSERT INTO Script(Texto) VALUES('IF (@Error <> 0)');
		INSERT INTO Script(Texto) VALUES('BEGIN');
		INSERT INTO Script(Texto) VALUES('	ROLLBACK TRANSACTION Insert_Tabla;');
		INSERT INTO Script(Texto) VALUES('');
		INSERT INTO Script(Texto) VALUES('	PRINT ''Error logueando registro ' + CONVERT(VARCHAR(10), @Contador) + ' en tabla log.'';');
		INSERT INTO Script(Texto) VALUES('');
		INSERT INTO Script(Texto) VALUES('	RETURN;');
		INSERT INTO Script(Texto) VALUES('END');
		INSERT INTO Script(Texto) VALUES('');
	END ELSE BEGIN
		INSERT INTO Script(Texto) VALUES('');
	END
END


-- Inserto commit de transacción
INSERT INTO Script(Texto) VALUES('COMMIT TRANSACTION Insert_Tabla;');
INSERT INTO Script(Texto) VALUES('');
INSERT INTO Script(Texto) VALUES('');
INSERT INTO Script(Texto) VALUES('RETURN;');


-- Generación del script
IF (RIGHT(@PathExportacion, 1) <> '\')
BEGIN
	SELECT	@PathExportacion = @PathExportacion + '\';
END

DECLARE @PathScript VARCHAR(1023);

SELECT	@PathScript = @PathExportacion + 'INSERT - ' + @Tabla + '.sql';

DECLARE @Cmd	NVARCHAR(4000);

SELECT	@Cmd = 'type nul > "' + @PathScript + '"';

EXECUTE @Error = master.dbo.xp_cmdshell @Cmd, NO_OUTPUT;

IF (@Error <> 0)
BEGIN
	PRINT	'Error creando archivo .sql.';

	DROP TABLE Script;

	DROP TABLE ##TMP_Registros;

	RETURN;
END


-- Copiado masivo de datos a archivo .sql
SELECT	@Cmd = 'bcp "SELECT NULLIF(Texto, '''') FROM master.dbo.Script ORDER BY Linea" queryout "' + @PathScript + '" -w -T';

EXECUTE @Error = xp_CmdShell @Cmd, NO_OUTPUT;

IF (@Error <> 0)
BEGIN
	PRINT	'Error copiando contenido de tabla en archivo .sql.';

	DROP TABLE Script;

	DROP TABLE ##TMP_Registros;

	RETURN;
END


-- Elimino tablas usadas
DROP TABLE Script;

DROP TABLE ##TMP_Registros;


-- Printeo el path donde se generó el script con éxito
PRINT	'Script generado en ' + @PathScript;


RETURN;