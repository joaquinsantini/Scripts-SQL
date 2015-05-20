-- Script que devuelve una tabla con el número de ocurrencias y la posición de donde empieza una frase que se busca en otra

SET NOCOUNT ON;

-- Variables a utilizar
DECLARE @Frase		VARCHAR(7900),
	@FraseABuscar	VARCHAR(7900),
	@FraseAux	VARCHAR(7900),
	@Contador	INT,
	@Salir		INT,
	@Pos		INT,
	@PosAnterior	INT;

CREATE TABLE #Posiciones (
	Ocurrencia	INT,
	Posicion	INT
);


-- @Frase: Frase en donde se va a buscar
SELECT	@Frase = '<Frase en donde se va a buscar>';


-- @FraseABuscar: Frase que se quiere buscar
SELECT	@FraseABuscar = '<Frase que se quiere buscar>';

-------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- Seteo variables
SELECT	@Contador = 0,
	@Salir = 0,
	@PosAnterior = 0;

-- LOOP principal de búsqueda
WHILE (@Salir = 0)
BEGIN
	SELECT	@Pos = CHARINDEX(@FraseABuscar, @Frase);

	IF (@Pos <> 0)
	BEGIN
		SELECT	@Contador = @Contador + 1;

		INSERT INTO #Posiciones (Ocurrencia, Posicion)
			VALUES (@Contador, @PosAnterior + @Pos);

		SELECT	@PosAnterior = @PosAnterior + @Pos;

		SELECT	@Frase = SUBSTRING(@Frase, @Pos, 7900),
			@FraseAux = SUBSTRING(@Frase, 2, 7900),
			@Pos = CHARINDEX(@FraseABuscar, @FraseAux);

		IF (@Pos <> 0)
		BEGIN
			SELECT	@Frase = SUBSTRING(@Frase, 2, 7900);
		END ELSE BEGIN
			SELECT	@Salir = 1;
		END
	END ELSE BEGIN
		SELECT	@Salir = 1;
	END
END


-- Muestro resultados
SELECT * FROM #Posiciones;

DROP TABLE #Posiciones;


RETURN;