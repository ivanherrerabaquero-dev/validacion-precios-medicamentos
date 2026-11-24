# Validación de precios de medicamentos y gestión de PQR en licitaciones y convenios

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/ivanherrerabaquero-dev/validacion-precios-medicamentos/blob/main/notebooks/proyecto_validacion_precios.ipynb)
[![Portal web](https://img.shields.io/badge/Portal-GitHub%20Pages-1f4e79)](https://ivanherrerabaquero-dev.github.io/validacion-precios-medicamentos/)

Proyecto de portafolio que simula el trabajo del área de **PQR y apoyo a licitaciones y convenios** de **CRM TECNOLOGÍAS EN SALUD**, un gestor farmacéutico ficticio (distribuidor mayorista y dispensador de medicamentos).

> ℹ️ **Datos:** los **precios máximos, CUM y medicamentos son reales**: provienen de la tabla pública de *Precios máximos de venta por presentación comercial* de la Circular 022 de 2026 de la Comisión Nacional de Precios de Medicamentos y Dispositivos Médicos (38.999 presentaciones). La empresa **CRM TECNOLOGÍAS EN SALUD**, los compradores, NIT, tarifas ofertadas, cantidades, precios piso y PQR son **ficticios**.

## El problema

Los compradores institucionales (grande, mediano y pequeño) envían sus bases de medicamentos en Excel con los precios propuestos para cada convenio o licitación. Cada archivo llega con un formato distinto y con errores. Antes de legalizar el acuerdo hay que:

1. **Consolidar** todas las bases en una sola.
2. **Validar** que ningún precio supere el **precio máximo de venta (PMV)** regulado y que no quede por debajo del precio piso interno.
3. **Recibir y hacer seguimiento a las PQR** de usuarios internos (equipos de licitaciones y convenios) y de compradores, dentro del término legal.
4. Entregar **informes ejecutivos** y **documentos gerenciales**.

## Funciones del cargo que cubre

| Función | Entregable |
|---|---|
| Consolidación de datos | Paso 3 del notebook |
| Generación de informes ejecutivos | Excel con formato + portal web |
| Recepción de PQR | Paso 5: base de PQR con días hábiles y oportunidad |
| Generación de documentos gerenciales | Memorando en Word generado automáticamente |
| Atención a usuarios internos de licitaciones y convenios | Segmentación de PQR por solicitante |

## Cómo funciona

```
data/entrada/*.xlsx  ──►  Notebook (Colab)  ──►  data/salida/consolidado_validacion.xlsx
 (3 compradores)          · lectura robusta       reportes/memorando_gerencial.docx
data/referencia/          · validación PMV        docs/data/*.json  ──►  Portal web (GitHub Pages)
 (PMV + catálogo)         · PQR y KPIs
```

**Regla clave de validación:** la Circular 022 de 2026 publica el PMV **por presentación comercial (CUM)** para tres tipos de transacción. Como las licitaciones y convenios son ventas a compradores institucionales, cada línea se cruza por CUM y se compara con el **PMV institucional**:

```
Tarifa ofertada por presentación  >  PMV institucional del CUM   →  línea rechazada
Ej.: IRBESARTÁN 150 mg x 25 tabletas: tarifa $20.143 vs PMV $17.250  →  +16,8 %
```

### Validaciones por línea

- CUM con formato válido (expediente-consecutivo o IUM) y presente en la tabla de la Circular 022.
- Líneas duplicadas.
- Tarifa, IVA y cantidad diligenciados correctamente.
- Tarifa **superior al PMV** → la línea se rechaza (riesgo de sanción).
- Tarifa **inferior al precio piso** interno → riesgo de margen.

### PQR

- Días hábiles calculados con fines de semana y festivos de Colombia 2026.
- Término de respuesta: 15 días hábiles (Ley 1755 de 2015).
- Indicadores: oportunidad, tiempo promedio, abiertas fuera de término, motivos y tipo de solicitante.

## Resultados (corte 20/11/2026)

| Indicador | Valor |
|---|---|
| Líneas validadas | 77 |
| Aprobadas sin observaciones | 71,4 % |
| Líneas sobre el PMV | 6 |
| Excedente estimado sobre PMV | $25,1 millones/mes |
| PQR recibidas | 420 |
| Oportunidad de respuesta | 85,9 % |

![Estado de las líneas](reportes/graficos/estado_lineas.png)

## Estructura

```
├── notebooks/proyecto_validacion_precios.ipynb   Notebook paso a paso (Colab)
├── data/
│   ├── entrada/       Bases de los compradores (Excel simulados)
│   ├── referencia/    Tabla oficial de PMV (Circular 022 de 2026, CSV) y catálogo de la empresa
│   └── salida/        Consolidado validado en Excel
├── reportes/          Memorando gerencial (Word) y gráficos
└── docs/              Portal web (GitHub Pages) + datos JSON + descargas
```

## Cómo ejecutarlo

1. Abre el notebook con el botón **Abrir en Colab**.
2. Ejecuta las celdas en orden (*Entorno de ejecución → Ejecutar todas*).
3. Opcional: cuando la Comisión publique una nueva versión de la tabla, súbela como `data/referencia/circular_oficial_nueva.xlsb` y ejecuta el Paso 2B.
4. Opcional: activa `PUBLICAR = True` en el Paso 9 para subir los resultados y actualizar el portal.

## Herramientas

Python · pandas · NumPy · openpyxl · python-docx · Matplotlib · HTML/JavaScript · Chart.js · GitHub Pages · Google Colab

## Marco normativo de referencia

- Circular 022 de 2026, Comisión Nacional de Precios de Medicamentos y Dispositivos Médicos: precios máximos de venta por presentación comercial, vigentes desde el 19 de agosto de 2026.
- Ley 1755 de 2015 (derecho de petición y términos de respuesta).

## Autor

**Ivan Santiago Herrera Baquero** · Administrador de Empresas · Gestión de PQRS y análisis de datos en salud
[LinkedIn](https://www.linkedin.com/in/ivan-santiago-herrera-baquero-523385204/) · [GitHub](https://github.com/ivanherrerabaquero-dev)
