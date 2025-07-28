# o7

Este repositorio incluye una plantilla de Excel para el control de pagos de una empresa.

## Archivos

- `create_excel.py`: genera un archivo `control_pagos.xls` en formato XML compatible con Excel.
- `control_pagos.xls`: plantilla resultante con dos hojas:
  - **Pagos**: registro de facturas y abonos.
  - **HistorialPorCliente**: resumen por cliente e invoice. Introduzca los nombres de cliente y el número de factura; los montos y saldos se calculan automáticamente con fórmulas.

## Uso

Ejecute el script (requiere Python 3) para volver a generar la plantilla:

```bash
python3 create_excel.py
```

Luego abra `control_pagos.xls` con Microsoft Excel (o software compatible) para llevar el control de pagos.
