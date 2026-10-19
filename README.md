# Libro-GHE · Generador de horarios ProContact

Crea el horario mensual de los ejecutivos ProContact con un solo clic. Cada mes se ajusta una planilla Excel, se presiona ▶ en el libro y sale el horario listo en Excel. Incluye turnos, sábados, pausas, colaciones y alertas.

## Qué hay en el repositorio

| Archivo / carpeta | Para qué sirve | ¿Se modifica? |
|---|---|---|
| `planificacion_turnos.ipynb` | El libro. Tiene una sola casilla con el botón ▶ | No (solo `MES` si quiere rehacer un mes) |
| `motor_turnos.py` | El programa que arma el horario | No |
| `plantillas/PLANTILLA_MES_EJEMPLO.xlsx` | Plantilla de ejemplo con ejecutivos ficticios | Se copia a `data/` y se llena |
| `plantillas/horario_mes_anterior_EJEMPLO.xlsx` | Horario de ejemplo de un mes anterior (octubre 2026, ficticio) | Solo para probar |
| `Horarios y Duración de Turnos.docx` | Documento con las reglas del negocio | No |
| `data/` *(no se sube)* | Sus datos reales: plantilla y horario del mes anterior | Sí, cada mes |
| `salida/` *(no se sube)* | Los horarios generados | No, solo se abren |

`data/` y `salida/` no se suben a git porque contienen RUT y nombres reales.

## Primera vez

1. Instale Python 3.10 o superior. Luego instale las librerías con:
   ```
   pip install pandas openpyxl jupyter
   ```
   O abra el libro en **VS Code** (con la extensión Jupyter) o en **Google Colab**.
2. Cree la carpeta `data/` y copie dentro las plantillas, quitando `_EJEMPLO` del nombre:
   - `plantillas/PLANTILLA_MES_EJEMPLO.xlsx` → `data/PLANTILLA_MES.xlsx`
   - el Excel del último horario real → `data/horario_mes_anterior.xlsx`
3. Abra `data/PLANTILLA_MES.xlsx` y reemplace los ejecutivos de ejemplo por los reales. Luego guarde y cierre el archivo.

> Si `data/PLANTILLA_MES.xlsx` no existe, al presionar ▶ se crea una plantilla vacía con las reglas originales.
>
> Para probar sin datos reales, copie también `plantillas/horario_mes_anterior_EJEMPLO.xlsx` como `data/horario_mes_anterior.xlsx`. Al presionar ▶ se creará noviembre 2026.

## Cada mes (3 pasos)

1. **Ajuste la plantilla** `data/PLANTILLA_MES.xlsx`. Cambie solo lo que cambió:
   - **EJECUTIVOS**: si alguien se va, ponga `Inactivo`. Si llega alguien, agréguelo al final.
   - **AUSENCIAS**: licencias, vacaciones o permisos, con fecha desde y hasta.
   - **FERIADOS**: revise que estén los del mes.

   Luego **guarde y cierre** el Excel.
2. **Abra `planificacion_turnos.ipynb` y presione ▶** en la casilla gris, o use `Shift + Enter`.
3. **Abra el horario** en `salida/PLANIFICACION_<MES>_<AÑO>.xlsx`.

El libro detecta solo qué mes crear: el siguiente al último horario disponible. El horario que se crea este mes sirve como "mes anterior" el próximo, así que no hay que copiar nada de nuevo.

Para **rehacer un mes**, cambie en la casilla `MES = "auto"` por, por ejemplo, `MES = "Noviembre 2026"`.

## La plantilla (`PLANTILLA_MES.xlsx`)

| Hoja | Contenido |
|---|---|
| INSTRUCCIONES | Paso a paso dentro del mismo Excel |
| EJECUTIVOS | RUT, nombre, estado y servicio (desde una lista), y horas de contrato |
| AUSENCIAS | RUT, tipo (LICENCIA / VACACIONES / PERMISO), desde y hasta |
| FERIADOS | Fechas no laborales |
| TURNOS | Los 5 turnos con sus horarios. La columna HORAS SEMANA debe dar 42 |
| PARAMETROS | Reglas: horas semanales, colación, pausas, % máximo en descanso, etc. |
| COBERTURA | Mínimo de ejecutivos conectados por hora (lunes a viernes y sábado) |

TURNOS, PARAMETROS y COBERTURA solo se cambian si cambian las reglas de la empresa.

## Qué hace el programa al presionar ▶

1. **Revisa la plantilla.** Si hay errores (RUT inválido o repetido, fechas al revés, un turno que no suma 42 h…), dice la hoja y la fila de cada uno, todos de una vez.
2. **Lee el horario anterior.** Lo toma de `data/horario_mes_anterior.xlsx` y de los horarios en `salida/`, para continuar la rotación de sábados y tardes.
3. **Arma el mes** según las reglas del documento:
   - 42 h semanales, de lunes a sábado.
   - 1 de cada 3 sábados trabajado, en forma pareja entre todos.
   - Turnos de tarde repartidos entre quienes han tenido menos.
   - Colación de 60 min de lunes a viernes y de 30 min el sábado.
   - Pausas de 15 min, sin superar el % máximo de la plataforma en descanso.
4. **Valida y exporta** un Excel con 4 hojas:
   - `PLANIFICACION_MENSUAL`: jornada diaria de cada ejecutivo.
   - `RESUMEN_EJECUTIVOS`: horas y sábados por persona, con fórmulas.
   - `MATRIZ_COBERTURA`: ejecutivos conectados por hora y día.
   - `ALERTAS`: lo que no se pudo cumplir, para revisión manual.

## Si algo sale mal

| Mensaje | Qué hacer |
|---|---|
| `Permission denied` / archivo en uso | Cierre el Excel de la plantilla o del horario y vuelva a presionar ▶ |
| `No module named motor_turnos` | `motor_turnos.py` debe estar en la misma carpeta que el libro (en Colab, súbalo también) |
| `No module named pandas` / `openpyxl` | Ejecute `pip install pandas openpyxl` |
| Errores en la plantilla | Corrija la hoja y fila que se indica, guarde, cierre y presione ▶ |
| No encuentra horario anterior | Copie el último horario como `data/horario_mes_anterior.xlsx`, o escriba el mes en `MES` |

**Alertas "DEMASIADOS EN DESCANSO A LA VEZ":** con 16 personas, no alcanza para cumplir a la vez el tope de 30% en descanso y las pausas obligatorias en el horario de 13:00 a 16:00. Para que desaparezcan, la empresa debe ajustar esa regla en PARAMETROS o aumentar la dotación.
