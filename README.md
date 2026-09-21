# Generador de reportes de servicio

**De registros de equipos y fotografías a reportes Word listos para revisión.** Esta aplicación de escritorio automatiza la elaboración de reportes de mantenimiento para equipos de unidades médicas. Permite seleccionar una unidad, filtrar los equipos que se incluirán y producir un documento individual por equipo, además de un reporte consolidado.

El proyecto reúne procesamiento de datos, búsquedas con SQLite, generación de documentos y una interfaz visual en Python. Está pensado para reducir tareas repetitivas al preparar reportes de servicio preventivo y correctivo.

## Qué hace

- **Organiza registros operativos:** importa datos desde JSON o CSV, normaliza campos y crea una base SQLite local con índices para las búsquedas.
- **Permite elegir el alcance:** busca unidades médicas y filtra por tipo de servicio, categoría de equipo o folios específicos.
- **Valida antes de generar:** identifica folios repetidos, inexistentes o que pertenecen a otra unidad o tipo de servicio.
- **Completa una plantilla Word:** sustituye campos de identificación del equipo, marca el tipo de servicio e inserta fotografías cuando hay imágenes disponibles para la categoría.
- **Entrega resultados trazables:** guarda los reportes individuales, un documento consolidado y un registro de generación con advertencias.
- **Procesa lotes:** puede crear documentos individuales en paralelo antes de unirlos en el reporte final.

Las categorías contempladas en el código son minisplit, cassette, UMA, paquete, chiller, bomba, torre y extractor. La aplicación puede generar un reporte sin fotografías; en ese caso registra la advertencia y conserva los marcadores de imagen de la plantilla.

## Flujo de trabajo

```text
JSON o CSV de equipos
        │
        ▼
Normalización e importación a SQLite
        │
        ▼
Búsqueda y selección en la interfaz
  unidad · servicio · categorías · folios
        │
        ▼
Plantilla DOCX + fotografías por categoría
        │
        ▼
Reportes individuales + reporte consolidado + log
```

## Tecnologías

| Componente | Uso en el proyecto |
| --- | --- |
| Python y Tkinter | Interfaz de escritorio y coordinación del flujo |
| SQLite | Almacenamiento local, índices y consultas de equipos |
| python-docx | Sustitución de campos y edición de la plantilla Word |
| docxcompose | Unión de documentos individuales |
| Pillow | Ajuste y recorte de imágenes para los espacios de la plantilla |
| concurrent.futures | Generación paralela de reportes individuales |

## Puesta en marcha

Necesitas Python 3 con Tkinter disponible y las dependencias de Word e imágenes. El repositorio incluye la plantilla `template_2.docx` y archivos de datos JSON/CSV; **no incluye las carpetas de fotografías**.

```bash
git clone https://github.com/magdaleno88/generador_reportes.git
cd generador_reportes
python -m pip install python-docx docxcompose Pillow
python "generador aires.py"
```

> El nombre del script contiene un espacio; conserva las comillas al ejecutarlo. En algunos sistemas, Tkinter se instala por separado de Python.

Al iniciar, la aplicación carga primero `base_datos_rf_ferems_CORREGIDA.json`. Si no está disponible, busca el CSV correspondiente y después las versiones anteriores definidas en el código. Con esa fuente crea `equipos.db` y vuelve a importar los registros cuando detecta cambios en el archivo de origen.

### Generar un reporte

1. Confirma la **plantilla Word** y elige una **carpeta de salida**.
2. Busca y selecciona la **unidad médica** y el servicio **Preventivo**, **Correctivo** o **Ambos**.
3. Elige entre todos los equipos, categorías concretas o una lista de folios. En el modo de folios puedes escribir valores separados por comas, espacios o punto y coma, así como rangos como `100-103`; pulsa **Validar folios** antes de generar.
4. Si quieres incluir evidencia fotográfica, selecciona las carpetas de imágenes de cada categoría. Para minisplit, el código busca las subcarpetas `completo`, `compresor` y `partes`.
5. Revisa el resumen y pulsa **Generar reportes**.

La aplicación permite cambiar la plantilla y la carpeta de salida desde la interfaz. También ofrece **Reconstruir base local** cuando necesitas forzar una nueva importación.

## Archivos de salida

Para una unidad y servicio seleccionados, la aplicación crea una estructura como esta:

```text
output/
└── nombre_de_la_unidad/
    ├── reportes_individuales/
    │   ├── 0001_folio_servicio.docx
    │   └── 0002_folio_servicio.docx
    ├── reporte_final_nombre_de_la_unidad_servicio.docx
    └── log_generacion.txt
```

Los nombres se normalizan para poder usarse como archivos. El documento final reúne los reportes individuales. El log registra la selección de fotografías, las rutas utilizadas, los recuentos y las advertencias detectadas durante la generación.

## Datos y plantilla

La fuente contiene columnas como `folio`, `servicio`, `region`, `unidad`, `marca`, `modelo`, `serie`, `inventario`, `equipo`, `categoria`, `tipo` y `pagina_pdf`. La aplicación usa esos registros para buscar y completar los reportes.

La plantilla Word puede incluir los marcadores `{{FOLIO}}`, `{{REGION}}`, `{{UNIDAD}}`, `{{MARCA}}`, `{{MODELO}}`, `{{SERIE}}`, `{{INVENTARIO}}`, `{{EQUIPO}}`, `{{SERVICIO}}` y `{{TIPO}}`. Las fotografías sustituyen imágenes insertadas en línea en la plantilla; conviene conservar sus espacios y proporciones al personalizarla.

**Antes de reutilizar los datos o compartir reportes generados**, confirma que tienes autorización para publicar información de unidades, equipos, números de serie y fotografías. Para una demostración de portafolio, utiliza datos e imágenes ficticios o anonimizados.

## Qué demuestra este proyecto

- Transformación de datos de distintas fuentes en una estructura consultable.
- Diseño de filtros y validaciones para evitar reportes de equipos equivocados.
- Automatización documental con plantillas y evidencia fotográfica.
- Interfaz de escritorio para un proceso operativo de varias etapas.
- Generación por lotes con salida individual, consolidada y registro de incidencias.

**Alcance actual:** es una aplicación de escritorio ejecutada desde el código fuente. El repositorio todavía no incluye un instalador ni una suite de pruebas automatizadas. Las fotografías son externas al repositorio y deben seleccionarse en la interfaz.
