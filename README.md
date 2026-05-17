# Reto Power Query Avanzado - Festivos de España 2023

## Descripción del proyecto

Este repositorio contiene la resolución del reto de **Power Query avanzado** en Power BI.

El objetivo del ejercicio es obtener una tabla de fechas con los **días festivos de España para el año 2023**, indicando la fecha y el nombre del festivo correspondiente.

La extracción y transformación de los datos se ha realizado íntegramente mediante **Lenguaje M** dentro del Editor de Power Query de Power BI.

---

## Fuente de datos

Los datos se obtienen desde la siguiente página web:

https://www.cuandoenelmundo.com/calendario/espana/2023

La página contiene el calendario de España para el año 2023 y los festivos nacionales, autonómicos y locales reflejados en cada fecha.

---

## Documentación utilizada

Para la realización del reto se ha utilizado como referencia la documentación oficial de Microsoft sobre funciones de Power Query M:

https://learn.microsoft.com/es-es/powerquery-m/power-query-m-function-reference

---

## Herramientas utilizadas

- Power BI Desktop
- Power Query
- Lenguaje M
- Git
- GitHub
- Visual Studio Code

---

## Objetivo del reto

Crear una tabla final que contenga:

| Campo | Descripción |
|---|---|
| Fecha | Fecha del día festivo |
| Festivo | Nombre del festivo correspondiente |

---

## Resultado final

La consulta final genera una tabla llamada:

```text
Festivos_Espana_2023
```

Con las siguientes columnas:

| Fecha | Festivo |
|---|---|
| 01/01/2023 | Año Nuevo |
| 06/01/2023 | Epifanía del Señor |
| 07/04/2023 | Viernes Santo |
| 01/05/2023 | Fiesta del Trabajo |
| ... | ... |

En los casos en los que una misma fecha contiene varios festivos, se ha añadido un separador para mejorar la legibilidad:

```text
Día de la Región de Murcia (MC) | Día de La Rioja (RI)
```

---

## Transformaciones realizadas

Todas las transformaciones se han realizado utilizando **Lenguaje M** en Power Query.

Los pasos principales han sido:

1. Definición de la URL base del calendario.
2. Creación de una tabla auxiliar con los 12 meses del año 2023.
3. Creación de una función personalizada para obtener los festivos de cada mes.
4. Conexión a la web mediante `Web.Contents`.
5. Lectura del contenido HTML mediante `Web.Page`.
6. Filtrado de las tablas obtenidas desde la web.
7. Combinación de las tablas mensuales.
8. Identificación de las filas que contienen festivos.
9. Extracción del día y del nombre del festivo.
10. Creación de la fecha completa.
11. Conversión de tipos de datos.
12. Limpieza de textos.
13. Ordenación final por fecha.

---

## Funciones de Lenguaje M utilizadas

Durante el desarrollo del reto se han utilizado las siguientes funciones indicadas en el enunciado:

```powerquery
Web.Contents
Web.Page
Table.SelectRows
Table.Combine
Table.CombineColumns
Table.TransformColumnTypes
```

Además, se han utilizado otras funciones complementarias de Power Query M para completar la transformación:

```powerquery
Table.AddColumn
Table.Distinct
Table.Sort
Table.TransformColumns
Date.FromText
Text.Trim
Text.From
Text.Lower
Text.Replace
List.Transform
List.Select
List.PositionOfAny
List.Accumulate
Record.FieldValues
Number.FromText
```

---

## Código M utilizado

```powerquery
let
    // URL base del calendario de España 2023
    UrlBase = "https://www.cuandoenelmundo.com/calendario/espana/2023",

    // Tabla auxiliar con los meses
    Meses = #table(
        type table [
            NumeroMes = Int64.Type,
            Mes = text,
            Slug = text
        ],
        {
            {1, "enero", "enero"},
            {2, "febrero", "febrero"},
            {3, "marzo", "marzo"},
            {4, "abril", "abril"},
            {5, "mayo", "mayo"},
            {6, "junio", "junio"},
            {7, "julio", "julio"},
            {8, "agosto", "agosto"},
            {9, "septiembre", "septiembre"},
            {10, "octubre", "octubre"},
            {11, "noviembre", "noviembre"},
            {12, "diciembre", "diciembre"}
        }
    ),

    // Función para obtener los festivos de cada mes
    ObtenerFestivosMes = (numeroMes as number, nombreMes as text, slugMes as text) as table =>
        let
            UrlMes = UrlBase & "/" & slugMes,

            // Carga de la página mensual
            Origen = Web.Page(Web.Contents(UrlMes)),

            // Nos quedamos solo con las tablas de la página
            Tablas = Table.SelectRows(
                Origen,
                each [Source] = "Table"
            ),

            // Eliminamos tablas vacías
            TablasConDatos = Table.SelectRows(
                Tablas,
                each try Table.RowCount([Data]) > 0 otherwise false
            ),

            // Combinamos las tablas encontradas
            TablaCombinada = Table.Combine(TablasConDatos[Data]),

            // Lista de días de la semana para identificar filas del calendario mensual
            DiasSemana = {
                "lunes",
                "martes",
                "miércoles",
                "miercoles",
                "jueves",
                "viernes",
                "sábado",
                "sabado",
                "domingo"
            },

            // Convertimos cada fila en una lista de valores limpios
            ValoresFila = Table.AddColumn(
                TablaCombinada,
                "Valores",
                each
                    List.Select(
                        List.Transform(
                            Record.FieldValues(_),
                            each try Text.Trim(Text.From(_)) otherwise ""
                        ),
                        each _ <> "" and _ <> "null"
                    )
            ),

            // Buscamos la posición del día de la semana dentro de cada fila
            PosDiaSemana = Table.AddColumn(
                ValoresFila,
                "PosDiaSemana",
                each List.PositionOfAny(
                    List.Transform([Valores], each Text.Lower(_)),
                    DiasSemana
                ),
                Int64.Type
            ),

            // El día del mes está justo antes del día de la semana
            DiaExtraido = Table.AddColumn(
                PosDiaSemana,
                "Dia",
                each try [Valores]{[PosDiaSemana] - 1} otherwise null,
                type text
            ),

            // El festivo, cuando existe, está justo después del día de la semana
            FestivoExtraido = Table.AddColumn(
                DiaExtraido,
                "Festivo",
                each try [Valores]{[PosDiaSemana] + 1} otherwise null,
                type text
            ),

            // Convertimos el día a número para validar que sea correcto
            DiaNumero = Table.AddColumn(
                FestivoExtraido,
                "DiaNumero",
                each try Number.FromText([Dia]) otherwise null,
                Int64.Type
            ),

            // Filtramos solo filas con festivo real
            FilasFestivos = Table.SelectRows(
                DiaNumero,
                each
                    [PosDiaSemana] > -1
                    and [DiaNumero] <> null
                    and [DiaNumero] >= 1
                    and [DiaNumero] <= 31
                    and [Festivo] <> null
                    and Text.Trim([Festivo]) <> ""
            ),

            // Dejamos las columnas necesarias
            ColumnasNecesarias = Table.SelectColumns(
                FilasFestivos,
                {"Dia", "Festivo"}
            ),

            // Añadimos el mes
            MesAgregado = Table.AddColumn(
                ColumnasNecesarias,
                "Mes",
                each nombreMes,
                type text
            ),

            // Añadimos el año
            AnioAgregado = Table.AddColumn(
                MesAgregado,
                "Anio",
                each "2023",
                type text
            ),

            // Cambiamos los tipos a texto antes de crear la fecha
            TiposTexto = Table.TransformColumnTypes(
                AnioAgregado,
                {
                    {"Dia", type text},
                    {"Mes", type text},
                    {"Anio", type text},
                    {"Festivo", type text}
                }
            ),

            // Combinamos día, mes y año en una columna de texto
            FechaTexto = Table.CombineColumns(
                TiposTexto,
                {"Dia", "Mes", "Anio"},
                Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),
                "FechaTexto"
            ),

            // Convertimos el texto a tipo fecha
            FechaCreada = Table.AddColumn(
                FechaTexto,
                "Fecha",
                each Date.FromText(
                    [FechaTexto],
                    [Format = "d MMMM yyyy", Culture = "es-ES"]
                ),
                type date
            ),

            // Resultado final de cada mes
            ResultadoMes = Table.SelectColumns(
                FechaCreada,
                {"Fecha", "Festivo"}
            )
        in
            ResultadoMes,

    // Aplicamos la función a cada uno de los meses
    DatosPorMes = Table.AddColumn(
        Meses,
        "Datos",
        each ObtenerFestivosMes([NumeroMes], [Mes], [Slug])
    ),

    // Combinamos todos los meses en una única tabla
    TablaFestivos = Table.Combine(DatosPorMes[Datos]),

    // Quitamos duplicados por seguridad
    SinDuplicados = Table.Distinct(TablaFestivos),

    // Añadimos separador cuando varios festivos vienen pegados en la misma celda
    FestivosSeparados = Table.TransformColumns(
        SinDuplicados,
        {
            {
                "Festivo",
                each
                    Text.Trim(
                        List.Accumulate(
                            {
                                "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
                                "K", "L", "M", "N", "Ñ", "O", "P", "Q", "R", "S",
                                "T", "U", "V", "W", "X", "Y", "Z",
                                "Á", "É", "Í", "Ó", "Ú"
                            },
                            Text.From(_),
                            (textoActual, letra) =>
                                Text.Replace(
                                    textoActual,
                                    ")" & letra,
                                    ") | " & letra
                                )
                        )
                    ),
                type text
            }
        }
    ),

    // Tipos finales de la tabla
    TiposFinales = Table.TransformColumnTypes(
        FestivosSeparados,
        {
            {"Fecha", type date},
            {"Festivo", type text}
        }
    ),

    // Ordenamos por fecha ascendente
    TablaFinal = Table.Sort(
        TiposFinales,
        {{"Fecha", Order.Ascending}}
    )
in
    TablaFinal
```

---

## Estructura del repositorio

```text
reto-festivos-espana-2023/
│
├── README.md
├── reto_festivos_espana_2023.pbix
└── .gitignore
```

---

## Archivo principal

El archivo principal del reto es:

```text
reto_festivos_espana_2023.pbix
```

Este fichero contiene la consulta desarrollada en Power Query y la tabla final de festivos de España para el año 2023.

---

## Cómo abrir el proyecto

Para revisar el proyecto:

1. Descargar o clonar este repositorio.
2. Abrir el archivo `.pbix` con Power BI Desktop.
3. Ir a **Transformar datos**.
4. Revisar la consulta `Festivos_Espana_2023`.
5. Verificar los pasos aplicados en el panel derecho de Power Query.

---

## Pasos para clonar el repositorio

```bash
git clone https://github.com/KvieraS/reto-festivos-espana-2023.git
```

Después, abrir el archivo `.pbix` con Power BI Desktop.

---

## Control de versiones

El proyecto ha sido subido a GitHub mediante Git desde Visual Studio Code.

Comandos principales utilizados:

```bash
git init
git add .
git commit -m "Primer commit del reto festivos España 2023"
git remote add origin https://github.com/KvieraS/reto-festivos-espana-2023.git
git branch -M main
git push -u origin main
```

---

## Conclusiones

Con este reto se ha practicado la extracción de datos desde una página web utilizando Power Query y Lenguaje M.

El resultado final es una tabla limpia, ordenada y preparada para ser utilizada en Power BI, con los festivos de España del año 2023.

Durante el proceso se han trabajado conceptos importantes como:

- Conexión a una fuente web.
- Lectura de tablas HTML.
- Filtrado de datos.
- Combinación de tablas.
- Creación de funciones personalizadas en Lenguaje M.
- Transformación de columnas.
- Conversión de tipos de datos.
- Limpieza de texto.
- Ordenación de la tabla final.

Este ejercicio permite reforzar el uso de Power Query como herramienta ETL dentro de Power BI.

---

## Autor

Proyecto realizado por Kevin Jesus Santoveña Viera como parte del módulo de Power Query avanzado.
