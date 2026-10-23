# ##############################################################################
#
#   EDUCACIÓN EN MÉXICO Y EN EL MUNDO: GASTO, COBERTURA Y RESULTADOS
#   Tablero en Shiny con los datos de la carpeta "DATOS DANIELA"
#
# ##############################################################################
#
#   Todo el tablero está en este script, dividido en cuatro partes:
#
#     PARTE 1. LIBRERÍAS Y FUNCIONES
#              1.1 Librerías
#              1.2 Colores, letra y paletas viridis
#              1.3 Funciones para leer los Excel
#              1.4 Funciones de formato (números y años)
#              1.5 Funciones de gráficas (plotly)
#              1.6 Piezas de la interfaz (tarjetas, encabezados, notas)
#              1.7 Módulos: cada hoja del tablero (UI + server de esa hoja)
#
#     PARTE 2. DATOS
#              2.1  Carpeta de datos
#              2.2  Gasto en educación como % del PIB
#              2.3  Gasto por estudiante
#              2.4  Escuelas y maestros por entidad
#              2.5  Gasto público en educación por entidad
#              2.6  Grado promedio de escolaridad por sexo
#              2.7  Tasa neta de matriculación
#              2.8  Tasa de abandono escolar
#              2.9  Catálogo de indicadores por entidad
#              2.10 Datos internacionales: gasto y PISA
#              2.11 Mapa de los estados
#              2.12 Fuentes y notas que se muestran en el tablero
#
#     PARTE 3. UI      (cómo se ve: pestañas, barras laterales, tarjetas)
#     PARTE 4. SERVER  (qué hace: arranca la lógica de cada hoja)
#
#   ¿Cómo se corre?
#     - En Positron o RStudio: botón "Run App" de este archivo, o
#     - seleccionar todo (Cmd/Ctrl + A) y correr (Cmd/Ctrl + Enter).
#     La última línea, shinyApp(ui, server), abre el tablero.
#
#   El script lee los Excel directamente de la carpeta "DATOS DANIELA"
#   (la busca en la carpeta de trabajo y hasta dos niveles arriba). Si se
#   corrige un Excel, basta con volver a correr el script.
#
#   Paquetes (se instalan una sola vez):
#   install.packages(c("shiny", "bslib", "bsicons", "plotly", "dplyr", "tidyr",
#                      "purrr", "stringr", "readxl", "jsonlite", "viridisLite",
#                      "stringi"))
#
# ##############################################################################



# ==============================================================================
# PARTE 1. LIBRERÍAS Y FUNCIONES
# ==============================================================================


# ------------------------------------------------------------------------------
# 1.1 Librerías
# ------------------------------------------------------------------------------

library(shiny)      # el tablero
library(bslib)      # tema, pestañas, tarjetas y barras laterales
library(bsicons)    # íconos
library(plotly)     # gráficas y mapas interactivos
library(dplyr)      # manipular tablas
library(tidyr)      # pasar tablas de ancho a largo
library(purrr)      # aplicar funciones a listas
library(stringr)    # limpiar texto
library(readxl)     # leer los Excel
# jsonlite, viridisLite y stringi se usan con "paquete::función".


# ------------------------------------------------------------------------------
# 1.2 Colores, letra y paletas viridis
# ------------------------------------------------------------------------------
# Verdes "naturaleza" para la interfaz y la familia viridis para TODAS las
# gráficas y mapas.

colores <- list(
  bosque   = "#1E4D2B",  # barra superior y títulos
  hoja     = "#2F7D4A",  # botones, pestañas activas, deslizadores
  musgo    = "#6E9254",
  salvia   = "#B9D3AE",  # bordes de tarjetas
  menta    = "#E3EFDC",  # fondo de las barras laterales
  fondo    = "#EEF5E9",  # fondo de la página
  texto    = "#1B2E21",
  rejilla  = "#E2EBDD",  # líneas guía de las gráficas
  eje      = "#9FB89A",
  sin_dato = "#D5DBD2",  # estados sin dato en los mapas
  apagado  = "#A9B8A4"   # puntos de países no elegidos
)

paletas_viridis <- c(
  "Viridis" = "viridis", "Magma" = "magma", "Plasma" = "plasma", "Inferno" = "inferno",
  "Cividis" = "cividis", "Mako" = "mako", "Rocket" = "rocket", "Turbo" = "turbo"
)

familia_letra <- "'Nunito Sans', 'Segoe UI', Helvetica, Arial, sans-serif"

# local = FALSE: la letra se pide a Google Fonts desde el navegador; sin
# internet se usa la siguiente letra de la lista.
tema_app <- bs_theme(
  version = 5,
  bg = colores$fondo, fg = colores$texto,
  primary = colores$hoja, secondary = colores$musgo, success = colores$hoja,
  base_font = font_collection(font_google("Nunito Sans", local = FALSE), "Segoe UI", "Helvetica", "Arial", "sans-serif"),
  heading_font = font_collection(font_google("Poppins", local = FALSE), "Segoe UI", "Helvetica", "sans-serif"),
  font_scale = 1.1,
  "border-radius" = "0.85rem",
  "card-border-color" = colores$salvia,
  "card-cap-bg" = "#FFFFFF",
  "card-bg" = "#FFFFFF"
)

# Colores para series (un estado, un país, un nivel). Se corta el amarillo más
# claro (fin = 0.85) para que las líneas se lean sobre fondo blanco.
colores_viridis <- function(n, opcion = "viridis", invertir = FALSE, inicio = 0.05, fin = 0.85) {
  cols <- substr(viridisLite::viridis(n, option = opcion, begin = inicio, end = fin), 1, 7)
  if (invertir) rev(cols) else cols
}

# Un color con nombre para cada serie.
colores_series <- function(nombres, opcion = "viridis", invertir = FALSE) {
  setNames(colores_viridis(max(length(nombres), 1), opcion, invertir)[seq_along(nombres)], nombres)
}

# Escala continua (mapas, barras del ranking y mapas de calor) en el formato de plotly.
escala_plotly <- function(opcion = "viridis", invertir = FALSE) {
  cols <- colores_viridis(11, opcion, invertir, inicio = 0, fin = 0.93)
  lapply(seq_along(cols), function(i) list((i - 1) / 10, cols[i]))
}


# ------------------------------------------------------------------------------
# 1.3 Funciones para leer los Excel
# ------------------------------------------------------------------------------

# Busca una carpeta o archivo en la carpeta de trabajo y hasta dos niveles arriba.
buscar_ruta <- function(nombre, tipo = c("carpeta", "archivo")) {
  tipo <- match.arg(tipo)
  candidatos <- c(nombre, file.path("..", nombre), file.path("..", "..", nombre))
  existe <- if (tipo == "carpeta") dir.exists(candidatos) else file.exists(candidatos)
  encontrada <- candidatos[existe][1]
  if (is.na(encontrada)) stop("No encuentro '", nombre, "'. Corre el script desde la carpeta SHiny.")
  encontrada
}

# Lee una hoja completa como texto, sin encabezados: así vemos el Excel tal cual
# y decidimos nosotras qué fila es el encabezado.
leer_excel <- function(carpeta, archivo, hoja = 1) {
  read_excel(file.path(carpeta, archivo), sheet = hoja, col_names = FALSE,
             col_types = "text", .name_repair = "minimal") |>
    as.data.frame()
}

# Texto a número; quita las comas de miles ("87,416,053,078").
a_numero <- function(x) suppressWarnings(as.numeric(str_remove_all(x, ",")))

# Claves INEGI de las entidades. Todos los archivos se unen por clave, nunca por
# nombre, porque cada archivo escribe los nombres distinto.
entidades <- tibble(
  cve_ent = sprintf("%02d", 0:32),
  entidad = c("Nacional", "Aguascalientes", "Baja California", "Baja California Sur", "Campeche",
              "Coahuila", "Colima", "Chiapas", "Chihuahua", "Ciudad de México", "Durango",
              "Guanajuato", "Guerrero", "Hidalgo", "Jalisco", "México", "Michoacán", "Morelos",
              "Nayarit", "Nuevo León", "Oaxaca", "Puebla", "Querétaro", "Quintana Roo",
              "San Luis Potosí", "Sinaloa", "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala",
              "Veracruz", "Yucatán", "Zacatecas")
)

# Quita acentos, guiones bajos y mayúsculas: "Nuevo_león" -> "nuevo leon".
normalizar <- function(x) {
  x |> stringi::stri_trans_general("Latin-ASCII") |> tolower() |>
    str_replace_all("_", " ") |> str_squish()
}

# Otras formas en que aparecen las entidades en los archivos.
otros_nombres <- c(
  "pais" = "00", "total" = "00", "estados unidos mexicanos" = "00",
  "cdmx" = "09", "distrito federal" = "09", "coahuila de zaragoza" = "05",
  "michoacan de ocampo" = "16", "veracruz de ignacio de la llave" = "30", "estado de mexico" = "15"
)

# Nombre de entidad -> clave INEGI. Si aparece un nombre desconocido, se detiene
# (mejor un error que un estado mal unido).
clave_entidad <- function(nombres) {
  diccionario <- c(otros_nombres, setNames(entidades$cve_ent, normalizar(entidades$entidad)))
  cve <- unname(diccionario[normalizar(nombres)])
  if (anyNA(cve)) stop("Entidades sin clave: ", paste(unique(nombres[is.na(cve)]), collapse = ", "))
  cve
}


# ------------------------------------------------------------------------------
# 1.4 Funciones de formato (números y años)
# ------------------------------------------------------------------------------

# Decimales según la magnitud: pesos sin decimales, porcentajes con uno,
# escolaridad o maestros por escuela con dos.
decimales_para <- function(x) {
  m <- median(abs(x), na.rm = TRUE)
  if (is.na(m)) 1 else if (m >= 100) 0 else if (m >= 10) 1 else 2
}

# Separador decimal "." y coma de miles (formato de México).
formato_num <- function(x, dec = 1) {
  ifelse(is.na(x), "sin dato", formatC(x, format = "f", digits = dec, big.mark = ","))
}
formato_d3 <- function(dec) paste0(",.", dec, "f")   # lo mismo, para plotly

# Si el año elegido no tiene dato (series bienales o por censo), usamos el más cercano.
anio_cercano <- function(anio, disponibles) {
  if (length(disponibles) == 0) return(NA_integer_)
  disponibles[which.min(abs(disponibles - anio))]
}

# Paso del deslizador de años: 1 si la serie es anual, 5 si es cada cinco años...
paso_anios <- function(anios) {
  difs <- unique(diff(sort(unique(anios))))
  if (length(difs) == 0) return(1)
  Reduce(function(a, b) if (b == 0) a else Recall(b, a %% b), difs)
}

primero_o_na <- function(x) if (length(x) == 0) NA_real_ else x[[1]]


# ------------------------------------------------------------------------------
# 1.5 Funciones de gráficas (plotly)
# ------------------------------------------------------------------------------
# Todas comparten: letra grande (15-17 px), ejes con título y unidad, colores
# viridis y separador decimal ".".

# --- Estilo común ---------------------------------------------------------------

eje_plotly <- function(titulo = NULL, ...) {
  modifyList(list(
    title = list(text = titulo, font = list(size = 17), standoff = 10),
    tickfont = list(size = 15), gridcolor = colores$rejilla, zeroline = FALSE,
    showline = TRUE, linecolor = colores$eje, linewidth = 1.5,
    ticks = "outside", tickcolor = colores$eje, automargin = TRUE
  ), list(...))
}

# Eje de años: con pocos años se marcan solo los que tienen dato.
eje_anios <- function(anios, ...) {
  anios <- sort(unique(anios))
  if (length(anios) <= 14) {
    eje_plotly("Año", tickvals = anios, ticktext = anios, ...)
  } else {
    eje_plotly("Año", tickformat = "d", dtick = if (diff(range(anios)) > 20) 5 else 2, tickangle = 0, ...)
  }
}

config_plotly <- function(p) {
  p |>
    layout(
      font = list(family = familia_letra, size = 16, color = colores$texto),
      separators = ".,",
      paper_bgcolor = "rgba(0,0,0,0)", plot_bgcolor = "rgba(0,0,0,0)",
      hoverlabel = list(font = list(size = 16, family = familia_letra),
                        bgcolor = "white", bordercolor = colores$salvia),
      legend = list(font = list(size = 15), orientation = "h", x = 0, y = -0.2,
                    title = list(text = ""))
    ) |>
    config(
      displaylogo = FALSE, locale = "es",
      modeBarButtonsToRemove = c("lasso2d", "select2d", "autoScale2d"),
      toImageButtonOptions = list(format = "png", scale = 2)
    )
}

# Con muchas series la leyenda va a la derecha, en columna.
leyenda_lateral <- function(p, n_series) {
  if (n_series <= 8) return(p)
  p |> layout(legend = list(orientation = "v", x = 1.02, y = 1, font = list(size = 13)))
}

# Parte un texto largo en renglones (títulos de barras de color).
partir <- function(texto, ancho = 18) gsub("\n", "<br>", str_wrap(texto, ancho))

# Barra de color horizontal debajo de los mapas, para que el mapa use todo el ancho.
barra_color_mapa <- function(unidad, dec) {
  list(orientation = "h", x = 0.5, xanchor = "center", y = 0, yanchor = "top",
       len = 0.85, thickness = 18, outlinewidth = 0,
       title = list(text = unidad, side = "top", font = list(size = 15)),
       tickfont = list(size = 15), tickformat = formato_d3(dec))
}

grafica_vacia <- function(mensaje) {
  plot_ly(type = "scatter", mode = "markers") |>
    layout(xaxis = list(visible = FALSE), yaxis = list(visible = FALSE),
           annotations = list(list(text = mensaje, showarrow = FALSE, font = list(size = 18),
                                   xref = "paper", yref = "paper", x = 0.5, y = 0.5))) |>
    config_plotly()
}

# --- Mapa por estado --------------------------------------------------------------
# Capa 1: los 32 estados en gris (un estado sin dato no desaparece del mapa).
# Capa 2: los estados con dato, coloreados con viridis.

graf_mapa_estados <- function(d, estados, geo, unidad, dec, opcion, invertir, rango, anio, fuente) {
  if (nrow(d) == 0) return(grafica_vacia("No hay datos para este año"))

  texto <- sprintf("<b>%s</b><br>%s %s<br>Lugar %d de %d (1 = valor más alto)<br>Año %s",
                   d$entidad, formato_num(d$valor, dec), unidad, d$lugar, nrow(d), anio)

  plot_ly(source = fuente) |>
    add_trace(
      type = "choropleth", geojson = geo, featureidkey = "properties.cve_ent",
      locations = estados$cve_ent, z = rep(1, nrow(estados)), customdata = estados$cve_ent,
      colorscale = list(list(0, colores$sin_dato), list(1, colores$sin_dato)),
      showscale = FALSE, marker = list(line = list(color = "white", width = 1)),
      text = paste0("<b>", estados$entidad, "</b><br>Sin dato en ", anio),
      hovertemplate = "%{text}<extra></extra>"
    ) |>
    add_trace(
      type = "choropleth", geojson = geo, featureidkey = "properties.cve_ent",
      locations = d$cve_ent, z = d$valor, customdata = d$cve_ent,
      zmin = rango[1], zmax = rango[2], colorscale = escala_plotly(opcion, invertir),
      marker = list(line = list(color = "white", width = 1.2)),
      text = texto, hovertemplate = "%{text}<extra></extra>",
      colorbar = barra_color_mapa(unidad, dec)
    ) |>
    layout(
      geo = list(fitbounds = "locations", visible = FALSE, bgcolor = "rgba(0,0,0,0)",
                 projection = list(type = "mercator")),
      margin = list(l = 0, r = 0, t = 10, b = 95), dragmode = FALSE
    ) |>
    config_plotly() |>
    event_register("plotly_click")
}

# --- Ranking de estados ---------------------------------------------------------------
# Barras horizontales ordenadas; los estados elegidos aparecen con borde y "◀".
# Admite valores negativos (el abandono escolar puede ser negativo).

graf_ranking <- function(d, nacional, seleccion, unidad, dec, opcion, invertir, rango, fuente) {
  if (nrow(d) == 0) return(grafica_vacia("No hay datos para este año"))

  d <- d |> arrange(valor)
  elegido <- d$cve_ent %in% seleccion
  etiquetas_y <- ifelse(elegido, paste0("<b>", d$entidad, " ◀</b>"), d$entidad)
  maximo <- max(c(d$valor, nacional, 0), na.rm = TRUE)
  minimo <- min(c(d$valor, nacional, 0), na.rm = TRUE)

  linea_nacional <- list()
  texto_nacional <- list()
  if (!is.na(nacional)) {
    linea_nacional <- list(list(type = "line", x0 = nacional, x1 = nacional, yref = "paper", y0 = 0, y1 = 1,
                                layer = "below", line = list(color = colores$texto, width = 2, dash = "dash")))
    texto_nacional <- list(list(x = nacional, y = 1, yref = "paper", yanchor = "bottom", yshift = 6,
                                showarrow = FALSE, xanchor = "center", font = list(size = 15),
                                text = paste0("<b>Nacional: ", formato_num(nacional, dec), "</b>")))
  }

  plot_ly(
    d, x = ~valor, y = ~entidad, type = "bar", orientation = "h", source = fuente,
    customdata = ~cve_ent,
    marker = list(
      color = ~valor, colorscale = escala_plotly(opcion, invertir), cmin = rango[1], cmax = rango[2],
      line = list(color = ifelse(elegido, colores$texto, "rgba(0,0,0,0)"), width = ifelse(elegido, 3, 0))
    ),
    text = ~formato_num(valor, dec), textposition = "outside", cliponaxis = FALSE,
    textfont = list(size = 14, color = colores$texto),
    hovertext = sprintf("<b>%s</b><br>%s %s<br>Lugar %d de %d", d$entidad,
                        formato_num(d$valor, dec), unidad, d$lugar, nrow(d)),
    hovertemplate = "%{hovertext}<extra></extra>"
  ) |>
    layout(
      xaxis = eje_plotly(unidad, range = c(minimo * 1.25, maximo * 1.16), tickformat = formato_d3(dec)),
      yaxis = eje_plotly(NULL, categoryorder = "array", categoryarray = d$entidad,
                         tickvals = d$entidad, ticktext = etiquetas_y, tickfont = list(size = 14),
                         showgrid = FALSE),
      shapes = linea_nacional, annotations = texto_nacional, bargap = 0.18,
      uniformtext = list(minsize = 13, mode = "show"),
      margin = list(l = 10, r = 30, t = 45, b = 10)
    ) |>
    config_plotly() |>
    event_register("plotly_click")
}

# --- Evolución en el tiempo (estados) -------------------------------------------------

graf_tendencia_estados <- function(d, seleccion, mostrar_nacional, anio, unidad, dec, opcion, invertir) {
  d <- d |> filter(cve_ent %in% c(if (mostrar_nacional) "00", seleccion))
  if (nrow(d) == 0) return(grafica_vacia("Elige estados en la barra lateral\no haz clic en el mapa"))

  # Los estados en el orden de la clave INEGI, cada uno con su color.
  seleccion <- sort(intersect(seleccion, unique(d$cve_ent)))
  cols <- setNames(colores_viridis(max(length(seleccion), 1), opcion, invertir)[seq_along(seleccion)], seleccion)
  formato <- paste0("%{y:", formato_d3(dec), "}")
  p <- plot_ly()

  for (cve in seleccion) {
    s <- d |> filter(cve_ent == cve) |> arrange(anio)
    p <- p |> add_trace(
      data = s, x = ~anio, y = ~valor, name = s$entidad[1], type = "scatter", mode = "lines+markers",
      connectgaps = TRUE, line = list(color = cols[[cve]], width = 3.5),
      marker = list(color = cols[[cve]], size = 9),
      hovertemplate = paste0("<b>", s$entidad[1], "</b>: ", formato, "<extra></extra>")
    )
  }
  if (mostrar_nacional && any(d$cve_ent == "00")) {
    s <- d |> filter(cve_ent == "00") |> arrange(anio)
    p <- p |> add_trace(
      data = s, x = ~anio, y = ~valor, name = "Nacional", type = "scatter", mode = "lines+markers",
      connectgaps = TRUE, line = list(color = colores$texto, width = 4.5, dash = "dot"),
      marker = list(color = colores$texto, size = 9, symbol = "diamond"),
      hovertemplate = paste0("<b>Nacional</b>: ", formato, "<extra></extra>")
    )
  }

  p |>
    layout(
      xaxis = eje_anios(d$anio),
      yaxis = eje_plotly(unidad, tickformat = formato_d3(dec)),
      hovermode = if (length(seleccion) > 10) "closest" else "x unified",
      shapes = list(list(type = "line", x0 = anio, x1 = anio, yref = "paper", y0 = 0, y1 = 1,
                         line = list(color = colores$musgo, width = 2, dash = "dash"))),
      margin = list(l = 10, r = 20, t = 20, b = 10)
    ) |>
    config_plotly() |>
    leyenda_lateral(length(seleccion) + mostrar_nacional)
}

# --- Mapa de calor: estados x categorías (niveles educativos o años) ------------------

graf_calor <- function(d, columnas, etiquetas, unidad, dec, opcion, invertir, titulo_barra = unidad) {
  if (nrow(d) == 0) return(grafica_vacia("No hay datos para este año"))

  filas <- d |> distinct(cve_ent, entidad) |> arrange(desc(cve_ent))   # Nacional arriba
  matriz <- d |>
    select(cve_ent, columna, valor) |>
    pivot_wider(names_from = columna, values_from = valor) |>
    right_join(filas, by = "cve_ent") |>
    arrange(match(cve_ent, filas$cve_ent))
  columnas <- intersect(columnas, names(matriz))

  plot_ly(
    x = partir(unname(etiquetas[columnas]), 13),
    y = ifelse(matriz$cve_ent == "00", "<b>Nacional</b>", matriz$entidad),
    z = as.matrix(matriz[, columnas]), type = "heatmap",
    colorscale = escala_plotly(opcion, invertir), xgap = 2, ygap = 2,
    texttemplate = paste0("%{z:.", dec, "f}"), textfont = list(size = 13),
    hovertemplate = paste0("<b>%{y}</b><br>%{x}<br>%{z:", formato_d3(dec), "} ", unidad, "<extra></extra>"),
    colorbar = list(title = list(text = partir(titulo_barra), font = list(size = 15)),
                    tickfont = list(size = 15), thickness = 20)
  ) |>
    layout(
      # type = "category": que los años de la brecha se lean como etiquetas, no como números.
      xaxis = eje_plotly(NULL, side = "top", tickangle = 0, tickfont = list(size = 14), showline = FALSE,
                         ticks = "", type = "category"),
      yaxis = eje_plotly(NULL, tickfont = list(size = 14), showline = FALSE, ticks = ""),
      margin = list(l = 10, r = 10, t = 10, b = 10)
    ) |>
    config_plotly()
}

# --- Líneas genéricas (países, niveles) -------------------------------------------------
# d: columnas serie, anio, valor. cols: un color por serie.

graf_lineas <- function(d, cols, unidad, dec, resaltar = NULL, anio_marca = NULL,
                        vacio = "Elige al menos una serie") {
  if (nrow(d) == 0) return(grafica_vacia(vacio))
  formato <- paste0("%{y:", formato_d3(dec), "}")
  series <- intersect(names(cols), unique(d$serie))
  p <- plot_ly()
  for (s in series) {
    x <- d |> filter(serie == s) |> arrange(anio)
    grueso <- s %in% resaltar
    p <- p |> add_trace(
      data = x, x = ~anio, y = ~valor, name = s, type = "scatter", mode = "lines+markers",
      connectgaps = TRUE,
      line = list(color = cols[[s]], width = if (grueso) 6 else 3),
      marker = list(color = cols[[s]], size = if (grueso) 12 else 9),
      hovertemplate = paste0("<b>", s, "</b>: ", formato, "<extra></extra>")
    )
  }
  marca <- if (is.null(anio_marca)) list() else
    list(list(type = "line", x0 = anio_marca, x1 = anio_marca, yref = "paper", y0 = 0, y1 = 1,
              line = list(color = colores$musgo, width = 2, dash = "dash")))
  p |>
    layout(
      xaxis = eje_anios(d$anio),
      yaxis = eje_plotly(unidad, tickformat = formato_d3(dec)),
      hovermode = if (length(series) > 10) "closest" else "x unified",
      shapes = marca, margin = list(l = 10, r = 20, t = 20, b = 10)
    ) |>
    config_plotly() |>
    leyenda_lateral(length(series))
}

# --- Gasto en educación como % del PIB --------------------------------------------------
# "areas": público + privado apilados (suman el nacional) y el total punteado encima.
# "lineas": las tres series como líneas.

graf_pib <- function(d, cols, modo, dec = 1) {
  if (nrow(d) == 0) return(grafica_vacia("Sin datos en el periodo"))
  formato <- paste0("%{y:", formato_d3(dec), "} % del PIB")
  p <- plot_ly()

  if (modo == "areas") {
    for (s in c("Gasto público", "Gasto privado")) {
      x <- d |> filter(serie == s) |> arrange(anio)
      p <- p |> add_trace(
        data = x, x = ~anio, y = ~valor, name = s, type = "scatter", mode = "lines",
        stackgroup = "uno", line = list(color = cols[[s]], width = 1.5),
        fillcolor = paste0(cols[[s]], "CC"),
        hovertemplate = paste0("<b>", s, "</b>: ", formato, "<extra></extra>")
      )
    }
    x <- d |> filter(serie == "Gasto nacional (total)") |> arrange(anio)
    p <- p |> add_trace(
      data = x, x = ~anio, y = ~valor, name = "Gasto nacional (total)", type = "scatter",
      mode = "lines+markers+text", text = ~formato_num(valor, dec), textposition = "top center",
      textfont = list(size = 13, color = colores$texto),
      line = list(color = colores$texto, width = 3, dash = "dot"),
      marker = list(color = colores$texto, size = 7),
      hovertemplate = paste0("<b>Gasto nacional (total)</b>: ", formato, "<extra></extra>")
    )
    techo <- max(d$valor) * 1.15
  } else {
    for (s in names(cols)) {
      x <- d |> filter(serie == s) |> arrange(anio)
      p <- p |> add_trace(
        data = x, x = ~anio, y = ~valor, name = s, type = "scatter", mode = "lines+markers",
        line = list(color = cols[[s]], width = if (s == "Gasto nacional (total)") 5 else 3),
        marker = list(color = cols[[s]], size = 9),
        hovertemplate = paste0("<b>", s, "</b>: ", formato, "<extra></extra>")
      )
    }
    techo <- max(d$valor) * 1.1
  }

  p |>
    layout(
      xaxis = eje_anios(d$anio),
      yaxis = eje_plotly("% del PIB", range = c(0, techo), tickformat = formato_d3(dec)),
      hovermode = "x unified", margin = list(l = 10, r = 20, t = 20, b = 10)
    ) |>
    config_plotly()
}

# Composición: de cada 100 pesos del gasto nacional, cuántos son públicos y cuántos privados.
graf_composicion <- function(d, cols) {
  if (nrow(d) == 0) return(grafica_vacia("Sin datos en el periodo"))
  p <- plot_ly()
  for (s in c("Gasto público", "Gasto privado")) {
    x <- d |> filter(serie == s) |> arrange(anio)
    p <- p |> add_trace(
      data = x, x = ~anio, y = ~pct, name = s, type = "bar", marker = list(color = cols[[s]]),
      text = ~formato_num(pct, 0), textposition = "inside", insidetextanchor = "middle",
      textfont = list(size = 13, color = "white"),
      hovertemplate = paste0("<b>", s, "</b>: %{y:.1f} % del gasto nacional<extra></extra>")
    )
  }
  p |>
    layout(
      barmode = "stack", bargap = 0.15,
      xaxis = eje_anios(d$anio),
      yaxis = eje_plotly("% del gasto nacional en educación", range = c(0, 100), ticksuffix = " %"),
      uniformtext = list(minsize = 12, mode = "hide"),
      hovermode = "x unified", margin = list(l = 10, r = 20, t = 20, b = 10)
    ) |>
    config_plotly()
}

# Cambio porcentual respecto a un año base (d: serie, anio, valor).
graf_indice <- function(d, cols, base) {
  if (nrow(d) == 0) return(grafica_vacia("Elige al menos un nivel"))
  d <- d |>
    group_by(serie) |>
    mutate(cambio = 100 * (valor / valor[anio == base] - 1)) |>
    ungroup() |>
    filter(!is.na(cambio))
  p <- plot_ly()
  for (s in intersect(names(cols), unique(d$serie))) {
    x <- d |> filter(serie == s) |> arrange(anio)
    p <- p |> add_trace(
      data = x, x = ~anio, y = ~cambio, name = s, type = "scatter", mode = "lines+markers",
      line = list(color = cols[[s]], width = 3.5), marker = list(color = cols[[s]], size = 9),
      hovertemplate = paste0("<b>", s, "</b>: %{y:+.1f} %<extra></extra>")
    )
  }
  p |>
    layout(
      xaxis = eje_anios(d$anio),
      yaxis = eje_plotly(paste0("Cambio respecto a ", base, " (%)"), ticksuffix = " %",
                         zeroline = TRUE, zerolinecolor = colores$texto, zerolinewidth = 2),
      hovermode = "x unified", margin = list(l = 10, r = 20, t = 20, b = 10)
    ) |>
    config_plotly()
}

# --- Gasto en educación vs. PISA ---------------------------------------------------------
# Un punto por país. Los países elegidos van con color y nombre; el resto en gris
# verdoso, para ver dónde queda cada uno frente a todos los demás.

graf_gasto_pisa <- function(d, elegidos, cols, materia, anio) {
  if (nrow(d) == 0) return(grafica_vacia(paste("Ningún país tiene gasto y PISA para", anio)))
  d <- d |> mutate(elegido = pais %in% elegidos) |> arrange(elegido)   # los elegidos quedan encima
  hover <- sprintf("<b>%s</b><br>Gasto en educación: %s %% del gasto público (%d)<br>PISA %s %d: %s puntos",
                   d$pais, formato_num(d$gasto, 1), d$anio_gasto, tolower(materia), anio, formato_num(d$pisa, 0))
  color_punto <- ifelse(d$elegido, unname(cols[d$pais]), colores$apagado)

  plot_ly(
    d, x = ~gasto, y = ~pisa, type = "scatter", mode = "markers+text",
    text = ifelse(d$elegido, d$pais, ""), textposition = "top center",
    textfont = list(size = 15, color = colores$texto),
    marker = list(size = ifelse(d$elegido, 20, 12), color = color_punto,
                  opacity = ifelse(d$elegido, 1, 0.75), line = list(color = "white", width = 1.5)),
    hovertext = hover, hovertemplate = "%{hovertext}<extra></extra>"
  ) |>
    layout(
      xaxis = eje_plotly("Gasto público en educación (% del gasto público total)", ticksuffix = " %",
                         range = c(min(d$gasto) - 1.5, max(d$gasto) + 1.5)),
      yaxis = eje_plotly(paste0("PISA ", tolower(materia), " (puntos)"),
                         range = c(min(d$pisa) - 25, max(d$pisa) + 25)),
      showlegend = FALSE, margin = list(l = 10, r = 20, t = 20, b = 10)
    ) |>
    config_plotly()
}

# --- Escolaridad de mujeres y hombres --------------------------------------------------
# Mancuernas: cada línea une a mujeres y hombres del mismo estado.

graf_mancuerna <- function(d, cols, seleccion_nombres) {
  if (nrow(d) == 0) return(grafica_vacia("Sin datos para este año"))
  d <- d |> arrange(Total)
  etiquetas_y <- ifelse(d$entidad %in% seleccion_nombres | d$cve_ent == "00",
                        paste0("<b>", d$entidad, "</b>"), d$entidad)
  segmentos <- lapply(seq_len(nrow(d)), function(i) list(
    type = "line", x0 = d$Mujeres[i], x1 = d$Hombres[i], y0 = d$entidad[i], y1 = d$entidad[i],
    layer = "below", line = list(color = colores$eje, width = 4)))

  plot_ly(d, y = ~entidad) |>
    add_trace(x = ~Mujeres, type = "scatter", mode = "markers", name = "Mujeres",
              marker = list(color = cols[["Mujeres"]], size = 13),
              hovertemplate = "<b>%{y}</b><br>Mujeres: %{x:.2f} años<extra></extra>") |>
    add_trace(x = ~Hombres, type = "scatter", mode = "markers", name = "Hombres",
              marker = list(color = cols[["Hombres"]], size = 13, symbol = "diamond"),
              hovertemplate = "<b>%{y}</b><br>Hombres: %{x:.2f} años<extra></extra>") |>
    layout(
      shapes = segmentos,
      xaxis = eje_plotly("Grado promedio de escolaridad (años)", tickformat = ".1f"),
      yaxis = eje_plotly(NULL, categoryorder = "array", categoryarray = d$entidad,
                         tickvals = d$entidad, ticktext = etiquetas_y, tickfont = list(size = 14),
                         showgrid = FALSE),
      margin = list(l = 10, r = 20, t = 40, b = 10)
    ) |>
    config_plotly() |>
    layout(legend = list(orientation = "h", x = 0, y = 1.04, font = list(size = 15)))
}


# ------------------------------------------------------------------------------
# 1.6 Piezas de la interfaz (tarjetas, encabezados, notas)
# ------------------------------------------------------------------------------

# Cada gráfica va en una tarjeta. "clave" sirve para mostrarla, ocultarla o
# moverla desde la barra lateral (lo hace el JavaScript js_app de la PARTE 3,
# sin volver a dibujar la gráfica, por eso el tablero no se traba).
tarjeta <- function(clave, titulo, icono, ..., ancho = FALSE) {
  div(
    class = paste("panel-item", if (ancho) "panel-ancho"),
    `data-panel` = clave,
    card(
      card_header(class = "titulo-tarjeta", bs_icon(icono), titulo),
      card_body(fillable = FALSE, ...)
    )
  )
}

# Contenedor de las tarjetas de una hoja.
rejilla <- function(id, ...) div(id = id, class = "panel-grid cols-2", ...)

# Nota visible (solo donde hace falta una aclaración importante).
nota <- function(...) div(class = "nota-caja", bs_icon("info-circle-fill"), div(...))

# Encabezado de cada hoja: título, etiquetas cortas y la fuente.
encabezado_hoja <- function(titulo, etiquetas, fuente, aviso = NULL) {
  etiquetas <- Filter(Negate(is.null), etiquetas)   # quita las etiquetas que no aplican
  div(
    class = "encabezado-seccion",
    h3(titulo),
    div(class = "chips", lapply(etiquetas, function(e) span(class = "chip", bs_icon(e[[1]]), e[[2]]))),
    p(class = "fuente", strong("Fuente: "), fuente),
    aviso
  )
}

# Controles que se repiten en todas las barras laterales.
controles_formato <- function(ns, tarjetas) {
  tagList(
    selectizeInput(ns("paneles"), "Gráficas a mostrar (en este orden)", choices = tarjetas,
                   selected = tarjetas, multiple = TRUE, options = list(plugins = list("remove_button"))),
    helpText(class = "ayuda", bs_icon("arrow-down-up"),
             "Quita una gráfica con la ×; al volver a agregarla se acomoda al final."),
    radioButtons(ns("columnas"), "Disposición", c("Dos columnas" = "2", "Una columna" = "1"), inline = TRUE),
    selectInput(ns("paleta"), "Paleta de color (familia viridis)", paletas_viridis),
    checkboxInput(ns("invertir"), "Invertir la paleta", FALSE)
  )
}

barra_lateral <- function(...) {
  sidebar(position = "right", width = 350, open = "desktop", bg = colores$menta,
          title = tags$span(class = "titulo-controles", bs_icon("sliders"), "Controles"), ...)
}

# Mensaje que el server manda al JavaScript para acomodar las tarjetas.
acomodar_tarjetas <- function(session, id_rejilla, orden, columnas) {
  session$sendCustomMessage("acomodar-paneles", list(
    id = id_rejilla, orden = as.list(orden %||% character(0)), columnas = columnas %||% "2"
  ))
}


# ------------------------------------------------------------------------------
# 1.7 Módulos: cada hoja del tablero
# ------------------------------------------------------------------------------
# Un módulo junta la UI y el server de una hoja. Así la misma hoja de estados se
# reusa para indicadores, matrícula y abandono, y los nombres de sus controles no
# chocan entre sí (NS() les pone un prefijo).


# ··· Hoja: gasto en educación como % del PIB ··································

series_pib <- c("Gasto nacional (total)", "Gasto público", "Gasto privado")

hoja_pib_ui <- function(id, pib, fuente) {
  ns <- NS(id)
  anios <- range(pib$anio)
  tarjetas <- c("Gasto como % del PIB" = "evolucion", "Composición: público y privado" = "composicion")

  layout_sidebar(
    class = "seccion", fillable = FALSE,
    sidebar = barra_lateral(
      radioButtons(ns("modo"), "Tipo de gráfica",
                   c("Áreas apiladas (público + privado = total)" = "areas", "Líneas" = "lineas")),
      sliderInput(ns("rango"), "Periodo", min = anios[1], max = anios[2], value = anios,
                  step = 1, sep = "", width = "100%"),
      controles_formato(ns, tarjetas)
    ),
    encabezado_hoja(
      "Gasto nacional, público y privado en educación como porcentaje del PIB, México",
      list(c("rulers", "% del PIB"), c("calendar3", paste0(anios[1], "–", anios[2])),
           c("layers", "Nacional = público + privado")),
      fuente
    ),
    rejilla(
      ns("rejilla"),
      tarjeta("evolucion", "Gasto nacional, público y privado (% del PIB)", "graph-up",
              plotlyOutput(ns("evolucion"), height = "520px"), ancho = TRUE),
      tarjeta("composicion", "Composición del gasto nacional en educación", "bar-chart-steps",
              plotlyOutput(ns("composicion"), height = "480px"),
              p(class = "pie-grafica", "Algunos años suman un poco menos de 100 % porque el archivo trae ",
                "las cifras redondeadas a un decimal."),
              ancho = TRUE)
    )
  )
}

hoja_pib_server <- function(id, pib) {
  moduleServer(id, function(input, output, session) {
    observe(acomodar_tarjetas(session, session$ns("rejilla"), input$paneles, input$columnas))

    cols <- reactive(setNames(colores_viridis(3, input$paleta, input$invertir, fin = 0.75), series_pib))
    en_rango <- reactive(pib |> filter(anio >= input$rango[1], anio <= input$rango[2]))

    output$evolucion <- renderPlotly(graf_pib(en_rango(), cols(), input$modo))

    # Parte pública y privada como % del gasto nacional de cada año.
    output$composicion <- renderPlotly({
      composicion <- en_rango() |>
        pivot_wider(names_from = serie, values_from = valor) |>
        mutate(`Gasto público` = 100 * `Gasto público` / `Gasto nacional (total)`,
               `Gasto privado` = 100 * `Gasto privado` / `Gasto nacional (total)`) |>
        select(anio, `Gasto público`, `Gasto privado`) |>
        pivot_longer(-anio, names_to = "serie", values_to = "pct")
      graf_composicion(composicion, cols())
    })
  })
}


# ··· Hoja: gasto por estudiante ·················································

hoja_estudiante_ui <- function(id, estudiante, unidad, fuente) {
  ns <- NS(id)
  anios <- range(estudiante$anio)
  niveles <- levels(estudiante$nivel)
  tarjetas <- c("Gasto por estudiante" = "evolucion", "Cambio respecto a un año base" = "cambio")

  layout_sidebar(
    class = "seccion", fillable = FALSE,
    sidebar = barra_lateral(
      checkboxGroupInput(ns("niveles"), "Niveles educativos", niveles, selected = niveles),
      selectInput(ns("base"), "Año base del cambio", seq(anios[1], anios[2]), selected = anios[1]),
      controles_formato(ns, tarjetas)
    ),
    encabezado_hoja(
      "Gasto en educación por estudiante, México",
      list(c("rulers", unidad), c("calendar3", paste0(anios[1], "–", anios[2]))),
      fuente
    ),
    rejilla(
      ns("rejilla"),
      tarjeta("evolucion", "Gasto por estudiante y nivel", "graph-up",
              plotlyOutput(ns("evolucion"), height = "500px")),
      tarjeta("cambio", "Cambio respecto al año base", "arrow-left-right",
              plotlyOutput(ns("cambio"), height = "500px"))
    )
  )
}

hoja_estudiante_server <- function(id, estudiante, unidad) {
  moduleServer(id, function(input, output, session) {
    observe(acomodar_tarjetas(session, session$ns("rejilla"), input$paneles, input$columnas))

    niveles <- levels(estudiante$nivel)
    datos <- estudiante |> mutate(serie = as.character(nivel))
    cols <- reactive(setNames(colores_viridis(length(niveles), input$paleta, input$invertir), niveles))
    elegidos <- reactive(datos |> filter(serie %in% input$niveles))

    output$evolucion <- renderPlotly(graf_lineas(elegidos(), cols(), unidad, 1, vacio = "Elige al menos un nivel"))
    output$cambio <- renderPlotly(graf_indice(elegidos(), cols(), as.integer(input$base)))
  })
}


# ··· Hoja con datos por entidad federativa ·····································
# Se usa tres veces: indicadores (escuelas, maestros, gasto, escolaridad),
# matrícula y abandono. Siempre trae mapa, ranking y evolución; además:
#   extras = "sexo"  -> escolaridad de mujeres y hombres + brecha por estado
#   extras = "calor" -> mapa de calor con todos los niveles por estado
#
# Los indicadores de tipo "conteo" (escuelas, maestros, pesos) se ven como
# totales en el ranking, pero el mapa usa su % del total nacional ("id_mapa"),
# porque un mapa de totales solo reflejaría qué tan grande es cada estado.

hoja_estados_ui <- function(id, meta, datos, etiqueta_selector, inicial,
                            extras = character(0), nota_calor = NULL) {
  ns <- NS(id)
  visibles <- meta |> filter(visible)
  anios <- sort(unique(datos$anio[datos$id %in% visibles$id]))
  estados <- datos |> filter(cve_ent != "00") |> distinct(cve_ent, entidad) |> arrange(cve_ent)

  # Opciones del selector, agrupadas por familia (Escuelas, Maestros, ...).
  grupos <- split(setNames(visibles$id, visibles$etiqueta), factor(visibles$familia, unique(visibles$familia)))
  opciones <- if (length(grupos) == 1) grupos[[1]] else lapply(grupos, as.list)

  # Estados que aparecen al abrir: el valor más alto y el más bajo del último año.
  ultimo <- datos |> filter(id == inicial, cve_ent != "00") |> filter(anio == max(anio))
  al_abrir <- c(ultimo$cve_ent[which.max(ultimo$valor)], ultimo$cve_ent[which.min(ultimo$valor)])

  tarjetas <- c(
    "Mapa por estado" = "mapa", "Ranking de estados" = "ranking", "Evolución en el tiempo" = "tendencia",
    if ("calor" %in% extras) c("Todos los niveles por estado" = "calor"),
    if ("sexo" %in% extras) c("Escolaridad de mujeres y hombres" = "mancuerna",
                              "Brecha entre mujeres y hombres" = "brecha")
  )

  layout_sidebar(
    class = "seccion", fillable = FALSE,
    sidebar = barra_lateral(
      selectInput(ns("id"), etiqueta_selector, choices = opciones, selected = inicial),
      sliderInput(ns("anio"), "Año (mapa y ranking)", min = min(anios), max = max(anios), value = max(anios),
                  step = paso_anios(anios), sep = "", ticks = TRUE, width = "100%",
                  animate = animationOptions(interval = 1800, loop = FALSE)),
      radioButtons(ns("escala"), "Escala de color",
                   c("Del año elegido (más contraste)" = "anio", "Fija para todos los años" = "fija")),
      selectizeInput(ns("estados"), "Estados en la evolución", multiple = TRUE,
                     choices = setNames(estados$cve_ent, estados$entidad), selected = al_abrir,
                     options = list(plugins = list("remove_button"), placeholder = "Elige estados…")),
      div(class = "botones-estados",
          actionButton(ns("todos"), "Todos los estados", icon = icon("check-double"), class = "btn-sm"),
          actionButton(ns("ninguno"), "Quitar todos", icon = icon("eraser"), class = "btn-sm")),
      checkboxInput(ns("nacional"), "Mostrar el dato nacional", TRUE),
      helpText(class = "ayuda", bs_icon("hand-index-thumb"),
               "También puedes hacer clic en un estado del mapa o del ranking para agregarlo o quitarlo."),
      controles_formato(ns, tarjetas)
    ),

    uiOutput(ns("encabezado")),
    rejilla(
      ns("rejilla"),
      tarjeta("mapa", "Mapa por estado", "map", plotlyOutput(ns("mapa"), height = "560px")),
      tarjeta("ranking", "Ranking de estados", "bar-chart-steps", plotlyOutput(ns("ranking"), height = "760px")),
      tarjeta("tendencia", "Evolución en el tiempo", "graph-up", plotlyOutput(ns("tendencia"), height = "560px"),
              ancho = TRUE),
      if ("calor" %in% extras)
        tarjeta("calor", "Todos los niveles por estado", "grid-3x3-gap",
                plotlyOutput(ns("calor"), height = "980px"), nota_calor, ancho = TRUE),
      if ("sexo" %in% extras)
        tarjeta("mancuerna", "Escolaridad de mujeres y hombres", "gender-ambiguous",
                plotlyOutput(ns("mancuerna"), height = "900px")),
      if ("sexo" %in% extras)
        tarjeta("brecha", "Brecha entre mujeres y hombres por estado (años)", "distribute-horizontal",
                plotlyOutput(ns("brecha"), height = "900px"))
    )
  )
}

hoja_estados_server <- function(id, meta, datos, geo, fuentes, aviso_escolaridad) {
  moduleServer(id, function(input, output, session) {
    ns <- session$ns
    estados <- datos |> filter(cve_ent != "00") |> distinct(cve_ent, entidad) |> arrange(cve_ent)
    datos_de <- function(i) datos[datos$id == i, ]
    con_lugar <- function(d) d |> mutate(lugar = min_rank(desc(valor)))

    observe(acomodar_tarjetas(session, ns("rejilla"), input$paneles, input$columnas))

    # --- Indicador elegido -------------------------------------------------------
    indicador <- reactive(meta[meta$id == req(input$id), ])
    es_conteo <- reactive(indicador()$tipo == "conteo")
    id_mapa   <- reactive(if (is.na(indicador()$id_mapa)) indicador()$id else indicador()$id_mapa)

    serie  <- reactive(datos_de(indicador()$id))
    dec    <- reactive(decimales_para(serie()$valor))
    anios  <- reactive(sort(unique(serie()$anio[serie()$cve_ent != "00"])))
    anio   <- reactive(anio_cercano(req(input$anio), anios()))

    del_anio <- reactive(serie() |> filter(anio == anio(), cve_ent != "00") |> con_lugar())
    nacional <- reactive({
      if (es_conteo()) return(NA_real_)   # en los conteos el "nacional" es la suma, no un promedio
      primero_o_na(serie()$valor[serie()$anio == anio() & serie()$cve_ent == "00"])
    })

    # El mapa puede mostrar otra medida (el % del total para los conteos).
    mapa_todos <- reactive(datos_de(id_mapa()) |> filter(cve_ent != "00"))
    mapa_anio  <- reactive(mapa_todos() |> filter(anio == anio()) |> con_lugar())

    rango_color <- function(todos, del_anio) {
      v <- if (input$escala == "fija") todos else del_anio
      if (length(v) == 0) c(0, 1) else range(v)
    }

    # --- Estados en la evolución -----------------------------------------------
    observeEvent(input$todos, updateSelectizeInput(session, "estados", selected = estados$cve_ent))
    observeEvent(input$ninguno, updateSelectizeInput(session, "estados", selected = character(0)))

    # Clic en mapa o ranking: agrega o quita ese estado.
    # suppressWarnings: plotly avisa si se escucha el clic antes de que exista la
    # gráfica (pasa en pestañas que aún no se abren); no es un error.
    alternar_estado <- function(clic) {
      cve <- clic$customdata
      if (is.null(cve)) return()
      sel <- input$estados
      updateSelectizeInput(session, "estados", selected = if (cve %in% sel) setdiff(sel, cve) else c(sel, cve))
    }
    clic_mapa    <- reactive(suppressWarnings(event_data("plotly_click", source = ns("mapa"))))
    clic_ranking <- reactive(suppressWarnings(event_data("plotly_click", source = ns("ranking"))))
    observeEvent(clic_mapa(), alternar_estado(clic_mapa()))
    observeEvent(clic_ranking(), alternar_estado(clic_ranking()))

    # --- Encabezado ----------------------------------------------------------------
    output$encabezado <- renderUI({
      m <- indicador()
      encabezado_hoja(
        m$titulo,
        list(c("rulers", m$unidad),
             c("calendar3", paste0(min(anios()), "–", max(anios()), " (", length(anios()), " años con dato)")),
             if (es_conteo()) c("map", "El mapa muestra el % del total nacional")),
        fuentes[[m$fuente]],
        if (m$familia == "Escolaridad promedio") aviso_escolaridad
      )
    })

    # --- Gráficas -------------------------------------------------------------------
    output$mapa <- renderPlotly({
      unidad_mapa <- meta$unidad[meta$id == id_mapa()]
      graf_mapa_estados(mapa_anio(), estados, geo, unidad_mapa, decimales_para(mapa_todos()$valor),
                        input$paleta, input$invertir, rango_color(mapa_todos()$valor, mapa_anio()$valor),
                        anio(), fuente = ns("mapa"))
    })

    output$ranking <- renderPlotly({
      graf_ranking(del_anio(), if (isTRUE(input$nacional)) nacional() else NA, input$estados,
                   indicador()$unidad, dec(), input$paleta, input$invertir,
                   rango_color(serie()$valor[serie()$cve_ent != "00"], del_anio()$valor), fuente = ns("ranking"))
    })

    output$tendencia <- renderPlotly({
      graf_tendencia_estados(serie(), input$estados, isTRUE(input$nacional) && !es_conteo(),
                             anio(), indicador()$unidad, dec(), input$paleta, input$invertir)
    })

    # Todos los niveles de la hoja en el año elegido (matrícula y abandono).
    output$calor <- renderPlotly({
      ids <- meta$id[meta$visible]
      anio_calor <- anio_cercano(input$anio, sort(unique(datos$anio[datos$id %in% ids])))
      d <- datos |> filter(id %in% ids, anio == anio_calor) |> mutate(columna = id)
      graf_calor(d, ids, setNames(meta$etiqueta, meta$id), indicador()$unidad, 1,
                 input$paleta, input$invertir, titulo_barra = paste(indicador()$unidad, "·", anio_calor))
    })

    # Escolaridad por sexo (solo en la hoja de indicadores).
    escolaridad <- reactive({
      datos |>
        filter(id %in% c("escol_total", "escol_mujeres", "escol_hombres")) |>
        mutate(sexo = recode(id, escol_total = "Total", escol_mujeres = "Mujeres", escol_hombres = "Hombres")) |>
        select(cve_ent, entidad, anio, sexo, valor) |>
        pivot_wider(names_from = sexo, values_from = valor)
    })
    cols_sexo <- reactive(setNames(colores_viridis(2, input$paleta, input$invertir, inicio = 0.15, fin = 0.75),
                                   c("Mujeres", "Hombres")))

    output$mancuerna <- renderPlotly({
      d <- escolaridad()
      a <- anio_cercano(input$anio, sort(unique(d$anio)))
      graf_mancuerna(d |> filter(anio == a), cols_sexo(), estados$entidad[estados$cve_ent %in% input$estados]) |>
        layout(title = list(text = paste("Año", a), x = 0.98, xanchor = "right", font = list(size = 16)))
    })

    # Brecha = mujeres − hombres, para cada estado y cada año con dato.
    output$brecha <- renderPlotly({
      d <- datos_de("escol_brecha") |> mutate(columna = as.character(anio))
      columnas <- as.character(sort(unique(d$anio)))
      graf_calor(d, columnas, setNames(columnas, columnas), "años (mujeres − hombres)", 2,
                 input$paleta, input$invertir, titulo_barra = "Mujeres − hombres (años)")
    })
  })
}


# ··· Hoja internacional: gasto en educación y PISA ·······························

materias_pisa <- c("Promedio", "Matemáticas", "Lectura", "Ciencia")

hoja_internacional_ui <- function(id, gasto, pisa, paises, al_abrir, america_latina, fuente) {
  ns <- NS(id)
  anios <- range(c(gasto$anio, pisa$anio))
  anios_pisa <- sort(unique(pisa$anio))
  tarjetas <- c("% del gasto público en educación" = "gasto", "Resultados PISA" = "pisa",
                "Gasto en educación y PISA en un mismo año" = "relacion")

  layout_sidebar(
    class = "seccion", fillable = FALSE,
    sidebar = barra_lateral(
      selectizeInput(ns("paises"), "Países", choices = paises, selected = al_abrir, multiple = TRUE,
                     options = list(plugins = list("remove_button"), placeholder = "Elige países…")),
      div(class = "botones-estados",
          actionButton(ns("latam"), "América Latina", icon = icon("earth-americas"), class = "btn-sm"),
          actionButton(ns("todos"), "Todos", icon = icon("check-double"), class = "btn-sm"),
          actionButton(ns("inicio"), "Selección inicial", icon = icon("rotate-left"), class = "btn-sm"),
          actionButton(ns("ninguno"), "Quitar todos", icon = icon("eraser"), class = "btn-sm")),
      checkboxInput(ns("resaltar"), "Resaltar a México (línea gruesa)", TRUE),
      selectInput(ns("materia"), "Materia de PISA", materias_pisa),
      sliderInput(ns("rango"), "Periodo de las líneas", min = anios[1], max = anios[2], value = anios,
                  step = 1, sep = "", width = "100%"),
      controles_formato(ns, tarjetas)
    ),
    encabezado_hoja(
      "Gasto público en educación y resultados en PISA por país",
      list(c("flag", paste(length(paises), "países")),
           c("calendar3", paste0("Gasto ", min(gasto$anio), "–", max(gasto$anio),
                                 " · PISA ", paste(anios_pisa, collapse = ", ")))),
      fuente
    ),
    rejilla(
      ns("rejilla"),
      tarjeta("gasto", "% del gasto público que se dedica a educación", "cash-coin",
              plotlyOutput(ns("gasto"), height = "520px")),
      tarjeta("pisa", "Resultados PISA", "mortarboard",
              plotlyOutput(ns("pisa"), height = "520px")),
      tarjeta("relacion", "Gasto en educación y PISA en un mismo año", "diagram-3",
              radioButtons(ns("anio_pisa"), "Año de PISA", choices = anios_pisa,
                           selected = max(anios_pisa), inline = TRUE),
              plotlyOutput(ns("relacion"), height = "600px"),
              p(class = "pie-grafica",
                "Cada punto es un país. Se usa el gasto del mismo año de PISA o, si falta, el del año con dato ",
                "más cercano anterior (hasta 3 años antes). Los países elegidos aparecen con color y nombre."),
              ancho = TRUE)
    )
  )
}

hoja_internacional_server <- function(id, gasto, pisa, paises, al_abrir, america_latina) {
  moduleServer(id, function(input, output, session) {
    observe(acomodar_tarjetas(session, session$ns("rejilla"), input$paneles, input$columnas))

    # Botones para elegir países rápido.
    elegir <- function(cuales) updateSelectizeInput(session, "paises", selected = cuales)
    observeEvent(input$latam, elegir(america_latina))
    observeEvent(input$todos, elegir(paises))
    observeEvent(input$inicio, elegir(al_abrir))
    observeEvent(input$ninguno, elegir(character(0)))

    elegidos  <- reactive(intersect(paises, input$paises))
    cols      <- reactive(colores_series(elegidos(), input$paleta, input$invertir))
    resaltar  <- reactive(if (isTRUE(input$resaltar)) "México" else NULL)
    anio_pisa <- reactive(as.integer(req(input$anio_pisa)))
    en_rango  <- function(d) d |> filter(pais %in% elegidos(), anio >= input$rango[1], anio <= input$rango[2])

    # Las dos gráficas, una junto a la otra.
    output$gasto <- renderPlotly({
      graf_lineas(en_rango(gasto) |> rename(serie = pais), cols(), "% del gasto público total", 1,
                  resaltar(), anio_marca = anio_pisa(), vacio = "Ningún país elegido tiene dato de gasto")
    })
    output$pisa <- renderPlotly({
      d <- en_rango(pisa |> filter(materia == input$materia, !is.na(valor))) |> rename(serie = pais)
      graf_lineas(d, cols(), paste0("PISA ", tolower(input$materia), " (puntos)"), 0,
                  resaltar(), anio_marca = anio_pisa(), vacio = "Ningún país elegido tiene dato de PISA")
    })

    # Relación gasto–PISA: todos los países con ambos datos, los elegidos con color.
    output$relacion <- renderPlotly({
      a <- anio_pisa()
      puntajes <- pisa |>
        filter(materia == input$materia, anio == a, !is.na(valor)) |>
        select(pais, pisa = valor)
      gasto_cercano <- gasto |>
        filter(anio <= a, anio >= a - 3) |>
        group_by(pais) |> filter(anio == max(anio)) |> ungroup() |>
        select(pais, gasto = valor, anio_gasto = anio)
      graf_gasto_pisa(inner_join(puntajes, gasto_cercano, by = "pais"), elegidos(), cols(), input$materia, a)
    })
  })
}



# ==============================================================================
# PARTE 2. DATOS
# ==============================================================================
# Aquí se leen y limpian todos los Excel. Cada sección termina con una tabla
# "larga" (una fila por entidad o país, año y valor) lista para las gráficas.
# Las comprobaciones con stopifnot() detienen el script si un Excel cambia de
# forma, en lugar de mostrar datos equivocados.


# ------------------------------------------------------------------------------
# 2.1 Carpeta de datos
# ------------------------------------------------------------------------------

carpeta <- buscar_ruta("DATOS DANIELA")


# ------------------------------------------------------------------------------
# 2.2 Gasto en educación como % del PIB (México, 2000-2022)
# ------------------------------------------------------------------------------
# Columnas: Año, Gasto_nac, Gasto_pub, Gasto_priv. El nacional es la suma de
# público y privado.

pib_crudo <- leer_excel(carpeta, "Gasto_Educ_2000-2022_R..xlsx")
fila_titulos <- which(pib_crudo[[1]] == "Año")

gasto_pib <- pib_crudo[(fila_titulos + 1):nrow(pib_crudo), 1:4] |>
  setNames(c("anio", "Gasto nacional (total)", "Gasto público", "Gasto privado")) |>
  filter(str_detect(anio, "^\\d{4}$")) |>
  pivot_longer(-anio, names_to = "serie", values_to = "texto") |>
  filter(!is.na(texto), texto != "") |>
  transmute(anio = as.integer(anio), serie, valor = a_numero(texto))

stopifnot(!anyNA(gasto_pib$valor))


# ------------------------------------------------------------------------------
# 2.3 Gasto por estudiante (miles de pesos de 2025, 2015-2025)
# ------------------------------------------------------------------------------

estudiante_crudo <- leer_excel(carpeta, "GastoNac_Estud_2015-2025_R..xlsx")
fila_titulos <- which(estudiante_crudo[[1]] == "Año")
niveles_estudiante <- as.character(estudiante_crudo[fila_titulos, 2:5])

# La unidad viene en una fila entre paréntesis: "(Miles de MXN de 2025 por estudiante)".
unidad_estudiante <- estudiante_crudo[[1]][str_detect(coalesce(estudiante_crudo[[1]], ""), "^\\(")][1] |>
  str_remove_all("[()]")

gasto_estudiante <- estudiante_crudo[(fila_titulos + 1):nrow(estudiante_crudo), 1:5] |>
  setNames(c("anio", niveles_estudiante)) |>
  filter(str_detect(anio, "^\\d{4}$")) |>
  pivot_longer(-anio, names_to = "nivel", values_to = "texto") |>
  filter(!is.na(texto), texto != "") |>
  transmute(anio = as.integer(anio), nivel = factor(nivel, niveles_estudiante), valor = a_numero(texto))

stopifnot(!anyNA(gasto_estudiante$valor))


# ------------------------------------------------------------------------------
# 2.4 Escuelas y maestros por entidad y nivel educativo (INEGI)
# ------------------------------------------------------------------------------
# "ND" = no disponible (el nivel Inicial no tiene dato antes de 2020).

em_crudo <- leer_excel(carpeta, "escuelas_maestros_EF_R..xlsx")
fila_titulos <- which(em_crudo[[1]] == "ent_fed")

escuelas_maestros <- em_crudo[(fila_titulos + 1):nrow(em_crudo), 1:5] |>
  setNames(c("ent", "anio", "nivel", "maestros", "escuelas")) |>
  filter(str_detect(anio, "^\\d{4}$")) |>
  transmute(
    cve_ent  = clave_entidad(ent),
    anio     = as.integer(anio),
    nivel    = str_replace_all(nivel, "_", " "),
    maestros = if_else(maestros == "ND", NA_real_, a_numero(maestros)),
    escuelas = if_else(escuelas == "ND", NA_real_, a_numero(escuelas))
  ) |>
  filter(!is.na(escuelas) | !is.na(maestros))

niveles_em <- c("Total", "Inicial", "Preescolar", "Primaria", "Secundaria", "Media superior", "Superior")
stopifnot(setequal(unique(escuelas_maestros$nivel), niveles_em))


# ------------------------------------------------------------------------------
# 2.5 Gasto público en educación por entidad (INEGI, EFIPEM)
# ------------------------------------------------------------------------------
# Una columna por estado y una fila por año. Se pasa a millones de pesos para que
# los números se lean. Ciudad de México no viene en el archivo.

gasto_ef_crudo <- leer_excel(carpeta, "GastoPub_EF_2000-2025_R..xlsx")
fila_titulos <- which(gasto_ef_crudo[[2]] == "Total")
nombres_columnas <- c("anio", as.character(gasto_ef_crudo[fila_titulos, -1]))

gasto_ef <- gasto_ef_crudo[(fila_titulos + 1):nrow(gasto_ef_crudo), ] |>
  setNames(nombres_columnas) |>
  filter(str_detect(anio, "^\\d{4}$")) |>
  pivot_longer(-anio, names_to = "ent", values_to = "texto") |>
  filter(!is.na(texto), texto != "") |>
  transmute(cve_ent = clave_entidad(ent), anio = as.integer(anio), gasto = a_numero(texto))

stopifnot(!anyNA(gasto_ef$gasto))


# ------------------------------------------------------------------------------
# 2.6 Grado promedio de escolaridad por sexo (INEGI)
# ------------------------------------------------------------------------------

escol_crudo <- leer_excel(carpeta, "EscolaridaPromedio_EF_sexo_2000-2025_R.xlsx")
fila_titulos <- which(escol_crudo[[1]] == "ent_fed")

escolaridad <- escol_crudo[(fila_titulos + 1):nrow(escol_crudo), 1:5] |>
  setNames(c("ent", "anio", "Total", "Mujeres", "Hombres")) |>
  filter(str_detect(anio, "^\\d{4}$")) |>
  transmute(cve_ent = clave_entidad(ent), anio = as.integer(anio),
            Total = a_numero(Total), Mujeres = a_numero(Mujeres), Hombres = a_numero(Hombres))

stopifnot(!anyNA(escolaridad))


# ------------------------------------------------------------------------------
# 2.7 Tasa neta de matriculación por entidad y nivel (SEP, vía INEGI)
# ------------------------------------------------------------------------------
# Los años vienen como ciclos escolares ("2024/2025"); usamos el año de inicio.

# Lee los tabulados de INEGI que tienen el mismo formato: entidad, nivel y una
# columna por ciclo escolar (matrícula y abandono).
leer_tabulado_ciclos <- function(archivo, primera_columna, niveles) {
  crudo <- leer_excel(carpeta, archivo)
  fila <- which(crudo[[1]] == primera_columna)
  ciclos <- as.character(crudo[fila, 3:ncol(crudo)])
  crudo[(fila + 1):nrow(crudo), ] |>
    setNames(c("ent", "nivel", ciclos)) |>
    filter(nivel %in% niveles) |>                       # descarta notas y renglones vacíos
    pivot_longer(all_of(ciclos), names_to = "ciclo", values_to = "texto") |>
    filter(!is.na(texto), texto != "") |>
    transmute(cve_ent = clave_entidad(ent), nivel, anio = as.integer(str_sub(ciclo, 1, 4)),
              valor = a_numero(texto))
}

matricula <- leer_tabulado_ciclos("Matricula%_EF_Grados_R..xlsx", "ent_fed",
                                  c("Preescolar", "Primaria", "Secundaria", "Media superior"))
stopifnot(!anyNA(matricula$valor), n_distinct(matricula$cve_ent) == 33)


# ------------------------------------------------------------------------------
# 2.8 Tasa de abandono escolar por entidad y nivel (SEP, vía INEGI)
# ------------------------------------------------------------------------------
# Puede ser negativa: en ese ciclo hubo más altas que bajas de alumnos.

abandono <- leer_tabulado_ciclos("Baja_EF_2000-2025.xlsx", "Entidad federativa",
                                 c("Primaria", "Secundaria", "Media superior", "Superior"))
stopifnot(!anyNA(abandono$valor), n_distinct(abandono$cve_ent) == 33)


# ------------------------------------------------------------------------------
# 2.9 Catálogo de indicadores por entidad
# ------------------------------------------------------------------------------
# Cada indicador es una lista con su ficha ("meta": nombre, unidad, fuente...)
# y sus datos (cve_ent, anio, valor). Al final se juntan en dos tablas.
#   tipo    = "conteo" si son totales (el mapa usa su % del total nacional)
#   hoja    = en qué hoja del tablero aparece
#   visible = si aparece en el selector (los % del total solo los usa el mapa)

nuevo_indicador <- function(id, datos, hoja, familia, etiqueta, titulo, unidad, fuente,
                            tipo = "tasa", visible = TRUE, id_mapa = NA_character_) {
  list(
    meta = tibble(id, hoja, familia, etiqueta, titulo, unidad, fuente, tipo, visible, id_mapa),
    datos = datos |> transmute(id = id, cve_ent, anio, valor) |> filter(!is.na(valor))
  )
}

# Porcentaje que representa cada estado del total de los estados, año por año.
como_pct_del_total <- function(d) {
  d |> filter(cve_ent != "00") |> group_by(anio) |> mutate(valor = 100 * valor / sum(valor, na.rm = TRUE)) |> ungroup()
}

sin_acentos <- function(x) str_replace_all(tolower(stringi::stri_trans_general(x, "Latin-ASCII")), " ", "_")
indicadores <- list()

# Escuelas, maestros y maestros por escuela, para cada nivel educativo.
for (n in niveles_em) {
  d <- escuelas_maestros |> filter(nivel == n)
  sufijo <- sin_acentos(n)

  for (cual in c("escuelas", "maestros")) {
    id <- paste0(str_sub(cual, 1, 3), "_", sufijo)
    nombre <- str_to_sentence(cual)
    totales <- d |> transmute(cve_ent, anio, valor = .data[[cual]])
    indicadores <- c(indicadores, list(
      nuevo_indicador(id, totales, "ef", nombre, paste0(nombre, ": ", tolower(n)),
                      paste0(nombre, " por entidad federativa, nivel ", tolower(n)),
                      cual, "escuelas_maestros", tipo = "conteo", id_mapa = paste0(id, "_pct")),
      nuevo_indicador(paste0(id, "_pct"), como_pct_del_total(totales), "ef", nombre,
                      paste0(nombre, ": ", tolower(n), " (% del total)"),
                      paste0(nombre, " del nivel ", tolower(n), " como % del total nacional"),
                      "% del total nacional", "escuelas_maestros", visible = FALSE)
    ))
  }

  indicadores <- c(indicadores, list(
    nuevo_indicador(paste0("mxe_", sufijo), d |> transmute(cve_ent, anio, valor = maestros / escuelas),
                    "ef", "Maestros por escuela", paste0("Maestros por escuela: ", tolower(n)),
                    paste0("Maestros por escuela, nivel ", tolower(n)), "maestros por escuela",
                    "escuelas_maestros")
  ))
}

# Gasto público en educación: total, % del total y por maestro.
gasto_maestro <- gasto_ef |>
  inner_join(escuelas_maestros |> filter(nivel == "Total"), by = c("cve_ent", "anio")) |>
  transmute(cve_ent, anio, valor = gasto / maestros)

indicadores <- c(indicadores, list(
  nuevo_indicador("gas_total", gasto_ef |> transmute(cve_ent, anio, valor = gasto / 1e6), "ef",
                  "Gasto público en educación", "Gasto público: millones de pesos",
                  "Gasto público en educación por entidad federativa", "millones de pesos", "gasto_ef",
                  tipo = "conteo", id_mapa = "gas_pct"),
  nuevo_indicador("gas_pct", como_pct_del_total(gasto_ef |> rename(valor = gasto)), "ef",
                  "Gasto público en educación", "Gasto público: % del total nacional",
                  "Gasto público en educación de cada estado como % de la suma de los estados",
                  "% del total nacional", "gasto_ef"),
  nuevo_indicador("gas_maestro", gasto_maestro, "ef", "Gasto público en educación", "Gasto público por maestro",
                  "Gasto público en educación por maestro (todos los niveles)", "pesos por maestro", "gasto_ef")
))

# Escolaridad promedio: total, mujeres, hombres y la brecha entre ambos.
for (s in c("Total", "Mujeres", "Hombres")) {
  indicadores <- c(indicadores, list(
    nuevo_indicador(paste0("escol_", tolower(s)), escolaridad |> transmute(cve_ent, anio, valor = .data[[s]]),
                    "ef", "Escolaridad promedio", paste0("Escolaridad: ", tolower(s)),
                    paste0("Grado promedio de escolaridad de la población de 15 años y más (", tolower(s), ")"),
                    "años de escolaridad", "escolaridad")
  ))
}
indicadores <- c(indicadores, list(
  nuevo_indicador("escol_brecha", escolaridad |> transmute(cve_ent, anio, valor = Mujeres - Hombres), "ef",
                  "Escolaridad promedio", "Escolaridad: brecha mujeres − hombres",
                  "Diferencia en el grado promedio de escolaridad: mujeres menos hombres",
                  "años (mujeres − hombres)", "escolaridad")
))

# Matrícula y abandono: un indicador por nivel educativo.
for (n in unique(matricula$nivel)) {
  indicadores <- c(indicadores, list(
    nuevo_indicador(paste0("mat_", sin_acentos(n)), matricula |> filter(nivel == n), "matricula", "Matrícula", n,
                    paste0("Tasa neta de matriculación en ", tolower(n)),
                    "% de la población en edad oficial del nivel", "matricula")
  ))
}
for (n in unique(abandono$nivel)) {
  indicadores <- c(indicadores, list(
    nuevo_indicador(paste0("aba_", sin_acentos(n)), abandono |> filter(nivel == n), "abandono", "Abandono", n,
                    paste0("Tasa de abandono escolar en ", tolower(n)),
                    "% de la matrícula de inicio de cursos", "abandono")
  ))
}

meta_estados  <- map(indicadores, "meta") |> list_rbind()
datos_estados <- map(indicadores, "datos") |> list_rbind() |> left_join(entidades, by = "cve_ent")

# Indicadores y datos de una hoja (incluye los % del total que usa su mapa).
datos_de_hoja <- function(hoja) {
  meta <- meta_estados |> filter(.data$hoja == !!hoja)
  meta <- bind_rows(meta, meta_estados |> filter(id %in% meta$id_mapa, !id %in% meta$id))
  list(meta = meta, datos = datos_estados |> filter(id %in% meta$id))
}
hoja_indicadores <- datos_de_hoja("ef")
hoja_matricula   <- datos_de_hoja("matricula")
hoja_abandono    <- datos_de_hoja("abandono")


# ------------------------------------------------------------------------------
# 2.10 Datos internacionales: gasto en educación y PISA
# ------------------------------------------------------------------------------
# El gasto (Banco Mundial) trae los nombres en español y PISA en inglés, así que
# traducimos PISA para poder juntar las dos bases por país.

# --- Gasto público en educación (% del gasto público total) ---
gasto_int_crudo <- leer_excel(carpeta, "GastPubInter_Educ_Filtrado.xlsx")
fila_titulos <- which(gasto_int_crudo[[1]] == "Country Name")
anios_gasto <- as.character(gasto_int_crudo[fila_titulos, -1])

gasto_int <- gasto_int_crudo[(fila_titulos + 1):nrow(gasto_int_crudo), ] |>
  setNames(c("pais", anios_gasto)) |>
  filter(!is.na(pais)) |>
  pivot_longer(-pais, names_to = "anio", values_to = "texto") |>
  filter(!is.na(texto), texto != "") |>
  transmute(pais, anio = as.integer(anio), valor = a_numero(texto)) |>
  # Un 0 % de gasto en educación es imposible: es una celda vacía mal capturada
  # (pasa en Armenia, Belarús y Brasil). Se trata como sin dato.
  filter(valor > 0)

stopifnot(!anyNA(gasto_int$valor))

# --- PISA ---
pisa_crudo <- leer_excel(carpeta, "Pisa_paises_2000-2025_Filtrado.xlsx")
fila_titulos <- which(pisa_crudo[[1]] == "Year")

# Nombres de PISA (inglés) -> nombres del Banco Mundial (español).
traduccion_paises <- c(
  "Albania" = "Albania", "Argentina" = "Argentina", "Armenia" = "Armenia", "Australia" = "Australia",
  "Austria" = "Austria", "Azerbaijan" = "Azerbaiyán", "Belarus" = "Belarús", "Belgium" = "Bélgica",
  "Bosnia and Herzegovina" = "Bosnia y Herzegovina", "Brazil" = "Brasil", "Brunei" = "Brunei Darussalam",
  "Brunei Darussalam" = "Brunei Darussalam", "Bulgaria" = "Bulgaria", "Cambodia" = "Camboya",
  "Canada" = "Canadá", "Chile" = "Chile", "Chinese Taipei" = "Taipéi Chino", "Colombia" = "Colombia",
  "Costa Rica" = "Costa Rica", "Croatia" = "Croacia", "Cyprus" = "Chipre", "Czech Republic" = "Chequia",
  "Denmark" = "Dinamarca", "Dominican Republic" = "República Dominicana", "Ecuador" = "Ecuador",
  "El Salvador" = "El Salvador", "Estonia" = "Estonia", "Finland" = "Finlandia", "France" = "Francia",
  "Georgia" = "Georgia", "Germany" = "Alemania", "Greece" = "Grecia", "Guatemala" = "Guatemala",
  "Hong Kong" = "Hong Kong (China)", "Hungary" = "Hungría", "Iceland" = "Islandia", "Indonesia" = "Indonesia",
  "Ireland" = "Irlanda", "Israel" = "Israel", "Italy" = "Italia", "Jamaica" = "Jamaica", "Japan" = "Japón",
  "Jordan" = "Jordania", "Kazakhstan" = "Kazajstán", "Kenya" = "Kenya", "Kosovo" = "Kosovo",
  "Kyrgyzstan" = "Kirguistán", "Latvia" = "Letonia", "Lebanon" = "Líbano", "Liechtenstein" = "Liechtenstein",
  "Lithuania" = "Lituania", "Luxembourg" = "Luxemburgo", "Macao" = "Macao (China)", "Malaysia" = "Malasia",
  "Malta" = "Malta", "Mexico" = "México", "Moldova" = "República de Moldova", "Mongolia" = "Mongolia",
  "Montenegro" = "Montenegro", "Morocco" = "Marruecos", "Netherlands" = "Países Bajos",
  "New Zealand" = "Nueva Zelandia", "North Macedonia" = "Macedonia del Norte", "Norway" = "Noruega",
  "Palestinian Authority" = "Autoridad Palestina", "Panama" = "Panamá", "Paraguay" = "Paraguay",
  "Peru" = "Perú", "Philippines" = "Filipinas", "Poland" = "Polonia", "Portugal" = "Portugal",
  "Qatar" = "Qatar", "Romania" = "Rumania", "Rwanda" = "Rwanda", "Saudi Arabia" = "Arabia Saudita",
  "Serbia" = "Serbia", "Singapore" = "Singapur", "Slovakia" = "Eslovaquia", "Slovenia" = "Eslovenia",
  "South Korea" = "Corea del Sur", "Spain" = "España", "Sweden" = "Suecia", "Switzerland" = "Suiza",
  "Thailand" = "Tailandia", "Tunisia" = "Túnez", "Turkey" = "Turquía", "Ukraine" = "Ucrania",
  "United Arab Emirates" = "Emiratos Árabes Unidos", "United Kingdom" = "Reino Unido",
  "United States" = "Estados Unidos", "Uruguay" = "Uruguay", "Uzbekistan" = "Uzbekistán",
  "Viet Nam" = "Viet Nam", "Vietnam" = "Viet Nam", "Zambia" = "Zambia"
)

pisa <- pisa_crudo[(fila_titulos + 1):nrow(pisa_crudo), 1:6] |>
  setNames(c("anio", "pais_ingles", "Promedio", "Matemáticas", "Lectura", "Ciencia")) |>
  filter(str_detect(anio, "^\\d{4}$")) |>
  mutate(pais = unname(traduccion_paises[pais_ingles])) |>
  pivot_longer(all_of(materias_pisa), names_to = "materia", values_to = "texto") |>
  # Las celdas vacías se quedan como NA (p. ej., España 2018 no tiene lectura ni promedio).
  transmute(pais, anio = as.integer(anio), materia, valor = a_numero(texto))

# Si PISA agrega un país que no está en la traducción, avisamos en lugar de perderlo.
sin_traducir <- setdiff(unique(pisa_crudo[[2]][-(1:fila_titulos)]), c(names(traduccion_paises), NA))
if (length(sin_traducir) > 0) stop("Países de PISA sin traducir: ", paste(sin_traducir, collapse = ", "))

# --- Listas de países para los selectores ---
paises <- sort(unique(c(gasto_int$pais, pisa$pais)))
paises_al_abrir <- intersect(c("México", "Argentina", "Brasil", "Chile", "Colombia", "Finlandia", "Corea del Sur"), paises)
paises_america_latina <- intersect(
  c("Argentina", "Brasil", "Chile", "Colombia", "Costa Rica", "Ecuador", "El Salvador", "Guatemala",
    "Jamaica", "México", "Panamá", "Paraguay", "Perú", "República Dominicana", "Uruguay"),
  paises
)


# ------------------------------------------------------------------------------
# 2.11 Mapa de los estados
# ------------------------------------------------------------------------------
# Se usa el archivo guardado en SHiny/data/. Si no está, se descarga (requiere
# internet) del repositorio strotgen/mexico-leaflet, que trae la clave INEGI de
# cada estado (state_code = 1..32).

descargar_mapa_estados <- function() {
  url <- "https://raw.githubusercontent.com/strotgen/mexico-leaflet/master/states.geojson"
  estados <- jsonlite::read_json(url, simplifyVector = FALSE)
  estados$features <- lapply(estados$features, function(f) {
    f$geometry$coordinates <- rapply(f$geometry$coordinates, function(v) round(v, 3), how = "replace")
    list(type = "Feature", properties = list(cve_ent = sprintf("%02d", as.integer(f$properties$state_code))),
         geometry = f$geometry)
  })
  estados
}

ruta_mapa <- tryCatch(buscar_ruta(file.path("data", "estados_mx.geojson"), "archivo"), error = function(e) NA)
mapa_estados <- if (!is.na(ruta_mapa)) {
  jsonlite::read_json(ruta_mapa, simplifyVector = FALSE)
} else {
  descargar_mapa_estados()
}


# ------------------------------------------------------------------------------
# 2.12 Fuentes y notas que se muestran en el tablero
# ------------------------------------------------------------------------------

enlace <- function(texto, url) tags$a(href = url, target = "_blank", texto)

fuentes <- list(
  pib = tagList(
    "Salgado Vega, M. del C. (2022), ",
    enlace("“Gasto Nacional y Público en Educación, México 2000-2022”",
           "https://economia.uaemex.mx/images/REVISTA%20ECONOMIA%20ACTUAL/Ano15/NUMERO%20ESPECIAL%2015/15-E-12.pdf"),
    ", Economía Actual (UAEMéx). Los datos originales son del INEGI (Banco de Indicadores) y de la SEP ",
    "(4.º Informe de Labores 2021-2022)."
  ),
  estudiante = tagList(
    "CIEP (2025), ",
    enlace("Gasto educativo 2025: ampliaciones en el PEF",
           "https://ciep.mx/wp-content/uploads/2025/01/Gasto-educativo-2025.-Ampliaciones-en-el-PEF_.pdf"),
    ", Figura 1, elaborada con información de la SHCP. De 2015 a 2023 son montos ejercidos; ",
    "2024 y 2025 son montos aprobados."
  ),
  escuelas_maestros = tagList(
    "INEGI, ",
    enlace("Maestros y escuelas por entidad federativa según nivel educativo, 2000/2001 a 2024/2025",
           "https://www.inegi.org.mx/app/tabulados/interactivos/?pxq=8c29ddc6-eeca-4dcc-8def-6c3254029f19"),
    "."
  ),
  gasto_ef = tagList(
    "INEGI, ",
    enlace("Estadística de Finanzas Públicas Estatales y Municipales (EFIPEM)",
           "https://www.inegi.org.mx/sistemas/olap/proyectos/bd/continuas/finanzaspublicas/fpest.asp?s=est&c=11288&proy=efipem_fest"),
    ". Pesos corrientes (sin descontar inflación). Ciudad de México no aparece en el archivo."
  ),
  escolaridad = tagList(
    "INEGI, ",
    enlace("Grado promedio de escolaridad de la población de 15 años y más por entidad federativa según sexo, 2000 a 2025",
           "https://www.inegi.org.mx/app/tabulados/interactivos/?pxq=Educacion_Educacion_05_2f6d2a08-babc-442f-b4e0-25f7d324dfe0"),
    "."
  ),
  matricula = tagList(
    "SEP, Reporte de indicadores educativos, publicado por el INEGI en ",
    enlace("Matrícula escolar por entidad federativa según nivel educativo, 2000/2001 a 2024/2025",
           "https://www.inegi.org.mx/app/tabulados/interactivos/?px=Educacion_06&bd=Educacion"),
    "."
  ),
  abandono = tagList(
    "SEP, Reporte de indicadores educativos, publicado por el INEGI en ",
    enlace("Tasa de abandono escolar por entidad federativa según nivel educativo, 2000/2001 a 2023/2024",
           "https://www.inegi.org.mx/app/tabulados/interactivos/?pxq=Educacion_Educacion_11_c6aa7c65-4d89-4eaf-972e-431727fc686d"),
    ". Un valor negativo significa que en ese ciclo hubo más altas que bajas de alumnos."
  ),
  internacional = tagList(
    "Gasto: Banco Mundial, ",
    enlace("Gasto público en educación, total (% del gasto del gobierno), SE.XPD.TOTL.GB.ZS",
           "https://datos.bancomundial.org/indicador/SE.XPD.TOTL.GB.ZS"),
    ". PISA: OCDE, puntajes recopilados en ",
    enlace("educationbycountry.org", "https://educationbycountry.org/pisa-rankings-by-country/"),
    "."
  )
)

# Nota sobre por qué la escolaridad no tiene datos para todos los años.
aviso_escolaridad <- nota(
  strong("¿Por qué no hay datos para todos los años? "),
  "El INEGI calcula el grado promedio de escolaridad con la pregunta sobre el último grado aprobado de sus ",
  "censos (2000, 2010 y 2020) y de sus encuestas intercensales (2015 y 2025), que entrevistan a una muestra ",
  "muy grande de hogares en todo el país. Entre un levantamiento y otro no hay una medición comparable: la SEP ",
  "publica cifras anuales, pero son estimaciones hechas con otra metodología (registros escolares y ",
  "proyecciones de población), por lo que no se mezclan con esta serie. Por eso la serie no es anual y conviene ",
  "comparar solo los años que sí tienen dato."
)

# Nota del mapa de calor de matrícula: por qué puede pasar de 100 %.
nota_matricula <- nota(
  strong("¿Por qué algunos estados pasan de 100 %? "),
  "La matrícula se cuenta en el estado donde está la escuela, pero la población en edad de estudiar se cuenta ",
  "donde vive. Si muchos alumnos cruzan a estudiar a otro estado (población flotante, como quienes viven en el ",
  "Estado de México y estudian en la Ciudad de México), la tasa del estado que los recibe puede superar 100 %. ",
  "También influye que la población se calcula con proyecciones de CONAPO."
)

nota_abandono <- nota(
  "Una tasa negativa significa que durante el ciclo hubo más altas que bajas de alumnos. ",
  "Los ceros no son absolutos: el INEGI los marca como no significativos."
)



# ==============================================================================
# PARTE 3. UI
# ==============================================================================
# Cómo se ve el tablero: dos pestañas principales (México e internacional) y,
# dentro de México, una pestaña por tema.


# ------------------------------------------------------------------------------
# 3.1 Estilo (CSS) y JavaScript
# ------------------------------------------------------------------------------

css_app <- r"---(
/* Colores: bosque #1E4D2B · hoja #2F7D4A · musgo #6E9254 · salvia #B9D3AE ·
   menta #E3EFDC · fondo #EEF5E9 · texto #1B2E21 */

body {
  font-size: 17px;
  background:
    radial-gradient(circle at 0% 0%, rgba(110, 146, 84, 0.10), transparent 40%),
    radial-gradient(circle at 100% 30%, rgba(47, 125, 74, 0.08), transparent 45%),
    #EEF5E9;
}

/* Barra superior */
.navbar { box-shadow: 0 2px 10px rgba(30, 77, 43, 0.25); }
.navbar .marca { font-family: Poppins, sans-serif; font-weight: 700; font-size: 1.3rem; letter-spacing: .2px; }
.navbar .marca .bi { color: #B9D3AE; margin-right: .4rem; }
.navbar .nav-link { font-size: 1.08rem; font-weight: 600; color: rgba(255, 255, 255, .82) !important; }
.navbar .nav-link:hover { color: #FFFFFF !important; }
.navbar .nav-link.active { color: #FFFFFF !important; border-bottom: 3px solid #B9D3AE; }
.fuentes-nav { font-size: .9rem; color: rgba(255, 255, 255, .75) !important; }

/* Título de cada pestaña principal */
.intro { padding: 1.2rem .25rem .6rem; }
.intro h2 { color: #1E4D2B; font-weight: 700; font-size: 1.75rem; margin-bottom: .3rem; }
.intro p { font-size: 1.05rem; color: #3D5A45; margin-bottom: .2rem; max-width: 70rem; }

/* Pestañas */
.nav-tabs { border-bottom: 2px solid #B9D3AE; }
.nav-tabs .nav-link { font-size: 1.05rem; font-weight: 600; color: #3D5A45; }
.nav-tabs .nav-link.active { color: #1E4D2B; background: #FFFFFF; border-color: #B9D3AE #B9D3AE #FFFFFF; border-width: 2px; }

/* Hoja con barra lateral a la derecha */
.seccion.bslib-sidebar-layout {
  border: 1px solid #B9D3AE; border-radius: 1rem; background: rgba(255, 255, 255, .55);
  box-shadow: 0 4px 18px rgba(30, 77, 43, 0.08);
}
.seccion > .sidebar .titulo-controles { font-weight: 800; font-size: 1.15rem; color: #1E4D2B; }
.seccion > .sidebar .titulo-controles .bi { margin-right: .4rem; }
.seccion > .sidebar label, .seccion > .sidebar .control-label { font-weight: 700; font-size: 1rem; color: #1E4D2B; }
.seccion > .sidebar .form-check-label, .seccion > .sidebar .radio label, .seccion > .sidebar .checkbox label { font-weight: 500; }
.seccion .ayuda { font-size: .92rem; color: #3D5A45; margin-top: -.4rem; }
.seccion .ayuda .bi { margin-right: .3rem; }
.botones-estados { display: flex; flex-wrap: wrap; gap: .4rem; margin: -.4rem 0 1rem; }
.botones-estados .btn { background: #FFFFFF; border: 1px solid #B9D3AE; color: #1E4D2B; font-weight: 600; }
.botones-estados .btn:hover { background: #2F7D4A; border-color: #2F7D4A; color: #FFFFFF; }

/* Deslizadores en verde */
.irs--shiny .irs-bar { background: #2F7D4A; border-color: #2F7D4A; }
.irs--shiny .irs-single, .irs--shiny .irs-from, .irs--shiny .irs-to { background: #1E4D2B; font-size: 14px; }
.irs--shiny .irs-handle { border-color: #2F7D4A; }
.irs--shiny .irs-grid-text { font-size: 12px; color: #3D5A45; }
.irs--shiny .irs-min, .irs--shiny .irs-max { font-size: 12px; }
.slider-animate-button { font-size: 1.3rem; color: #2F7D4A; }

/* Encabezado de cada hoja */
.encabezado-seccion { margin-bottom: 1rem; }
.encabezado-seccion h3 { color: #1E4D2B; font-weight: 700; font-size: 1.45rem; line-height: 1.3; }
.chips { display: flex; flex-wrap: wrap; gap: .5rem; margin: .5rem 0; }
.chip { background: #FFFFFF; border: 1px solid #B9D3AE; border-radius: 999px; padding: .2rem .8rem; font-size: .95rem; color: #1E4D2B; }
.chip .bi { margin-right: .35rem; color: #2F7D4A; }
.fuente { font-size: .95rem; color: #3D5A45; margin-bottom: .3rem; }

/* Notas visibles y pies de gráfica */
.nota-caja {
  display: flex; gap: .6rem; align-items: flex-start; margin: .8rem 0 0;
  background: #F6FAF3; border: 1px solid #D5E4CE; border-left: 5px solid #6E9254;
  border-radius: .7rem; padding: .7rem .9rem; font-size: .98rem; color: #1B2E21; line-height: 1.45;
}
.nota-caja > .bi { color: #2F7D4A; font-size: 1.15rem; margin-top: .1rem; flex-shrink: 0; }
.pie-grafica { font-size: .95rem; color: #3D5A45; margin: .5rem 0 0; }

/* Tarjetas: se muestran, ocultan y reordenan desde la barra lateral */
.panel-grid { display: grid; gap: 1.25rem; grid-template-columns: minmax(0, 1fr); }
.panel-grid.cols-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
.panel-grid .panel-ancho { grid-column: 1 / -1; }
.panel-oculto { display: none !important; }
.panel-item { animation: aparecer .35s ease; }
@keyframes aparecer { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }
@media (max-width: 1250px) { .panel-grid.cols-2 { grid-template-columns: minmax(0, 1fr); } }
.panel-item .card { height: 100%; margin-bottom: 0; box-shadow: 0 3px 14px rgba(30, 77, 43, 0.07); }
.titulo-tarjeta {
  font-family: Poppins, sans-serif; font-weight: 600; font-size: 1.12rem; color: #1E4D2B;
  border-left: 5px solid #2F7D4A; display: flex; align-items: center; gap: .5rem;
}
.titulo-tarjeta .bi { color: #2F7D4A; }
.shiny-options-group .form-check-inline { margin-right: 1rem; }
.card-body .shiny-input-radiogroup label { font-weight: 700; color: #1E4D2B; }
)---"

# Recibe el mensaje "acomodar-paneles" del server: muestra, oculta y ordena las
# tarjetas de una hoja cambiando solo el CSS (las gráficas no se vuelven a dibujar).
js_app <- r"---(
Shiny.addCustomMessageHandler("acomodar-paneles", function (msg) {
  var rejilla = document.getElementById(msg.id);
  if (!rejilla) return;

  rejilla.classList.toggle("cols-1", msg.columnas === "1");
  rejilla.classList.toggle("cols-2", msg.columnas !== "1");

  rejilla.querySelectorAll(":scope > .panel-item").forEach(function (tarjeta) {
    var lugar = msg.orden.indexOf(tarjeta.dataset.panel);
    if (lugar < 0) {
      tarjeta.classList.add("panel-oculto");
    } else {
      tarjeta.classList.remove("panel-oculto");
      tarjeta.style.order = lugar;
    }
  });

  // Plotly y Shiny recalculan tamaños (y qué gráficas están visibles) con "resize".
  setTimeout(function () { window.dispatchEvent(new Event("resize")); }, 80);
});
)---"


# ------------------------------------------------------------------------------
# 3.2 Pestañas del tablero
# ------------------------------------------------------------------------------

ui <- page_navbar(
  title = tags$span(class = "marca", bs_icon("mortarboard-fill"), "Gasto y resultados en educación"),
  window_title = "Gasto y resultados en educación",
  theme = tema_app,
  fillable = FALSE,
  navbar_options = navbar_options(bg = colores$bosque, theme = "dark"),
  header = tags$head(tags$style(HTML(css_app)), tags$script(HTML(js_app))),

  # --- Pestaña 1: México -------------------------------------------------------
  nav_panel(
    title = "México", icon = bs_icon("geo-alt-fill"),
    div(class = "intro",
        h2("Educación en México: gasto, cobertura y resultados"),
        ),
    navset_tab(
      id = "temas_mexico",
      nav_panel("Gasto como % del PIB",
                hoja_pib_ui("pib", gasto_pib, fuentes$pib)),
      nav_panel("Gasto por estudiante",
                hoja_estudiante_ui("estudiante", gasto_estudiante, unidad_estudiante, fuentes$estudiante)),
      nav_panel("Escuelas, maestros, gasto y escolaridad por estado",
                hoja_estados_ui("indicadores", hoja_indicadores$meta, hoja_indicadores$datos,
                                etiqueta_selector = "Indicador", inicial = "escol_total", extras = "sexo")),
      nav_panel("Matrícula por estado y nivel",
                hoja_estados_ui("matricula", hoja_matricula$meta, hoja_matricula$datos,
                                etiqueta_selector = "Nivel educativo", inicial = "mat_media_superior",
                                extras = "calor", nota_calor = nota_matricula)),
      nav_panel("Abandono escolar por estado y nivel",
                hoja_estados_ui("abandono", hoja_abandono$meta, hoja_abandono$datos,
                                etiqueta_selector = "Nivel educativo", inicial = "aba_media_superior",
                                extras = "calor", nota_calor = nota_abandono))
    )
  ),

  # --- Pestaña 2: Comparación internacional ------------------------------------
  nav_panel(
    title = "Comparación internacional", icon = bs_icon("globe-americas"),
    div(class = "intro",
        h2("Gasto en educación y resultado en PISA (2000 - 2025)"),
        ),
    hoja_internacional_ui("internacional", gasto_int, pisa, paises, paises_al_abrir,
                          paises_america_latina, fuentes$internacional)
  ),

  nav_spacer(),
  nav_item(tags$span(class = "navbar-text fuentes-nav", "Fuentes: INEGI · SEP · CIEP · Banco Mundial · OCDE"))
)



# ==============================================================================
# PARTE 4. SERVER
# ==============================================================================
# Qué hace el tablero: arranca la lógica de cada hoja (definida en los módulos
# de la PARTE 1) con sus datos (PARTE 2).

server <- function(input, output, session) {

  # México
  hoja_pib_server("pib", gasto_pib)
  hoja_estudiante_server("estudiante", gasto_estudiante, unidad_estudiante)
  hoja_estados_server("indicadores", hoja_indicadores$meta, hoja_indicadores$datos,
                      mapa_estados, fuentes, aviso_escolaridad)
  hoja_estados_server("matricula", hoja_matricula$meta, hoja_matricula$datos,
                      mapa_estados, fuentes, aviso_escolaridad)
  hoja_estados_server("abandono", hoja_abandono$meta, hoja_abandono$datos,
                      mapa_estados, fuentes, aviso_escolaridad)

  # Comparación internacional
  hoja_internacional_server("internacional", gasto_int, pisa, paises, paises_al_abrir, paises_america_latina)
}


# ------------------------------------------------------------------------------
# Abrir el tablero
# ------------------------------------------------------------------------------

shinyApp(ui, server)
