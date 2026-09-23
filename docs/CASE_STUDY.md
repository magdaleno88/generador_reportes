# Generador de reportes de servicio | Service Report Generator

[Español](#español) · [English](#english)

## Español

### El problema

Preparar un reporte de servicio por equipo exigía localizar registros, comprobar folios, completar una plantilla Word e insertar fotografías. Al repetirlo para varios equipos, aumentaban las tareas manuales y la posibilidad de seleccionar un registro equivocado.

### Lo que construí

Una aplicación de escritorio en Python que importa registros JSON o CSV a SQLite, permite filtrar por unidad, servicio y categoría, y valida los folios antes de generar documentos. Completa una plantilla Word, inserta fotografías por categoría cuando están disponibles y compone un reporte consolidado a partir de los individuales.

### Cómo comprobarlo

- Ejecutar `python examples/create_demo_report.py` para generar un reporte Word con datos ficticios mediante la función real del proyecto.
- Ejecutar `python -m unittest discover -s tests -v` para comprobar la selección de folios y la sustitución de campos.

El ejemplo muestra **un reporte individual**. El flujo de consolidación requiere `docxcompose` y se utiliza desde la aplicación completa. Todavía no hay una medición publicada del tiempo ahorrado ni un instalador. Los registros operativos y fotografías reales no forman parte de esta demostración.

## English

### The problem

Preparing a service report for each piece of equipment required finding the right record, checking its ID, filling a Word template and adding photos. Repeating those steps for many items created manual work and opportunities to select the wrong record.

### What I built

A Python desktop application that imports JSON or CSV records into SQLite, filters by facility, service and equipment category, and validates record IDs before generating documents. It fills a Word template, inserts category-matched photos when available, and combines individual reports into one consolidated document.

### How to inspect it

- Run `python examples/create_demo_report.py` to generate a sample Word report with fictional data using the project's report function.
- Run `python -m unittest discover -s tests -v` to check record selection and Word field replacement.

The sample demonstrates **one individual report**. Consolidation uses `docxcompose` in the full application. No measured time saving or installer is published yet. Operational records and real photos are excluded from this demo.
