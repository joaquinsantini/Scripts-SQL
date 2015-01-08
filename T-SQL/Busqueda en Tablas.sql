-- Script que busca un String en tablas de una determinada base de datos (puede ser en todas las tablas y todas sus columnas).

USE [<Base de datos>]
GO

-- Compruebo existencia de tabla temporal.
IF (ISNULL(OBJECT_ID('tempdb..#TMP_ColumnasTablas'), '') <> '')
BEGIN
	DROP TABLE #TMP_ColumnasTablas
END

-- Creo tabla temporal.
CREATE TABLE #TMP_ColumnasTablas (
	NombreTabla	VARCHAR(255),
	NombreColumna	VARCHAR(255),
	FlagProcesado	INT,
	FlagMostrar	INT
)

-- Lleno la tabla temporal con las tablas donde se quiere buscar (pueden ser todas las tablas de una base).
INSERT INTO #TMP_ColumnasTablas (NombreTabla, NombreColumna, FlagProcesado, FlagMostrar)
	SELECT	SO.name		AS 'NombreTabla',
		SC.name		AS 'NombreColumna',
		0		AS 'FlagProcesado',
		0		AS 'FlagMostrar'
		FROM	sys.objects SO
		INNER JOIN sys.columns SC
			ON SO.OBJECT_ID = SC.OBJECT_ID
		WHERE	SO.type = 'U' AND <Condiciones sobre las tablas y las columnas>
		/* Ejemplo donde se buscan columnas que contengan Pers o Perm en su nombre y no sean de baja ni de alta ni de modificacion
		WHERE	SO.type = 'U' AND
			SO.name NOT LIKE '%Log%' AND
			(SC.name LIKE '%Pers%' OR SC.name LIKE '%Perm%') AND
			(SC.name NOT LIKE '%Alta%' AND SC.name NOT LIKE '%Baja%' AND SC.name NOT LIKE '%Modif%') */
		ORDER BY SO.name, SC.name

DECLARE @Tabla		VARCHAR(255)
DECLARE @Columna	VARCHAR(255)
DECLARE @SQL		NVARCHAR(MAX)
DECLARE @Cantidad	INT
DECLARE @Error		INT

-- Loop principal de búsqueda.
WHILE (EXISTS(SELECT TOP (1) * FROM #TMP_ColumnasTablas WHERE FlagProcesado = 0))
BEGIN
	SELECT TOP (1)	@Tabla = TMP.NombreTabla,
			@Columna = TMP.NombreColumna
		FROM	#TMP_ColumnasTablas TMP
		WHERE	TMP.FlagProcesado = 0

	SELECT	@SQL = '	IF (EXISTS(SELECT TOP (1) * FROM ' + @Tabla + ' WHERE ' + @Columna + ' LIKE ''%<String a buscar>%''))
				BEGIN
					SELECT @Cantidad = 1
				END ELSE BEGIN
					SELECT @Cantidad = 0
				END'

	EXECUTE @Error = sp_executesql @SQL, N'@Cantidad INT OUTPUT', @Cantidad OUTPUT

	IF (@Error <> 0)
	BEGIN
		PRINT	'Error al ejecutar consulta dinámica.'

		DROP TABLE #TMP_ColumnasTablas

		RETURN
	END

	IF (@Cantidad <> 0)
	BEGIN
		UPDATE #TMP_ColumnasTablas
			SET	FlagMostrar = 1
			WHERE	NombreTabla = @Tabla AND
				NombreColumna = @Columna
	END

	UPDATE #TMP_ColumnasTablas
		SET	FlagProcesado = 1
		WHERE	NombreTabla = @Tabla AND
			NombreColumna = @Columna
END

-- Muestro resultados
SELECT	TMP.NombreTabla		AS 'Nombre Tabla',
	TMP.NombreColumna	AS 'Nombre Columna'
	FROM	#TMP_ColumnasTablas TMP
	WHERE	TMP.FlagMostrar = 1

-- Elimino tabla temporal.
DROP TABLE #TMP_ColumnasTablas


RETURN