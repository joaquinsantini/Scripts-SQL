-- Script que devuelve una tabla con el número de ocurrencia y la posición de donde empieza una frase que se busca en otra.
DECLARE @Frase		VARCHAR(MAX)
DECLARE @FraseABuscar	VARCHAR(MAX)

SELECT	@Frase = '<Frase en donde se va a buscar>'
SELECT	@FraseABuscar = '<Frase que se quiere buscar>'

CREATE TABLE #Posiciones (
	Ocurrencia	INT,
	Posicion	INT
)

DECLARE @FraseAux	VARCHAR(MAX)
DECLARE @Contador	INT
DECLARE @Salir		INT
DECLARE @Pos		INT
DECLARE @PosAnterior	INT

SELECT	@Contador = 0
SELECT	@Salir = 0
SELECT	@PosAnterior = 0

WHILE (@Salir = 0)
BEGIN
	SELECT	@Pos = CHARINDEX(@FraseABuscar, @Frase)
	IF (@Pos <> 0)
	BEGIN
		SELECT	@Contador = @Contador + 1

		INSERT INTO #Posiciones (Ocurrencia, Posicion)
			VALUES (@Contador, @PosAnterior + @Pos)

		SELECT	@PosAnterior = @PosAnterior + @Pos

		SELECT	@Frase = SUBSTRING(@Frase, @Pos, 4000)
		SELECT	@FraseAux = SUBSTRING(@Frase, 2, 4000)
		SELECT	@Pos = CHARINDEX(@FraseABuscar, @FraseAux)
		IF (@Pos <> 0)
		BEGIN
			SELECT	@Frase = SUBSTRING(@Frase, 2, 4000)
		END ELSE BEGIN
			SELECT	@Salir = 1
		END
	END ELSE BEGIN
		SELECT	@Salir = 1
	END
END

SELECT * FROM #Posiciones

DROP TABLE #Posiciones


RETURN