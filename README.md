# Forecast 2026 (Provexpress)

Dashboard comercial para seguimiento de forecast 2026. Es un sitio estatico que
se autentica con Microsoft 365 (MSAL) y carga archivos Excel desde SharePoint.

## Funcionalidades
- Carga automatica desde SharePoint (Microsoft Graph).
- Vistas por rol: gerencia, director y ejecutivo.
- Panel de cambio de vista habilitado solo para `especialista.preventa`.
- Modulo gerencial `Rebates, puntos e incentivos por canal`, con unidades
  separadas, canal como filtro principal, KPIs, graficos, ranking, insights,
  tabla maestra y exportacion a Excel.

## Rebates, puntos e incentivos por canal
La vista solo se habilita para los roles `gerencia` y `gerencia_director`.
Incluye una base gerencial validada para Dell, Lenovo, HPE, ASUS, Epson e Intel.
Si se cargan reportes vigentes, cada familia reemplaza su bloque base para evitar
duplicar valores; los canales sin archivo actualizado conservan la base validada.
Puntos, USD y COP siempre se consultan por separado.

Primero consulta esta carpeta dentro del mismo SharePoint de Forecast:

```
COMERCIAL/FORECAST 2026/Fabricantes
```

Alli carga la version mas reciente de estas familias de reportes:

- `Reporte palataformas Juan 2026.xlsx` (tambien acepta `plataformas`).
- `MyRewards Partner Detail Report*.xlsx`.
- `Claim Summary Report*.xlsx`.
- `Program_List_LAS_PA000006347092.xls`.

Si algun archivo no aparece en `Fabricantes`, realiza una busqueda de respaldo
en las bibliotecas accesibles de SharePoint.

Si algun reporte no esta en SharePoint, gerencia puede usar `Cargar Excel` para
procesarlo localmente en memoria. Los archivos y sus datos no se almacenan en el
repositorio. Las hojas resumen/pivote y las hojas vacias se excluyen para evitar
duplicar valores.

## Origen de datos
Ruta esperada en SharePoint (Documentos compartidos):

```
COMERCIAL/FORECAST 2026/Grupo [Director]/[Ejecutivo].xlsx
```

Reglas del archivo:
- Se toma la primera hoja que contenga "Gerencia" o "Comercial".
- La fila de encabezados se detecta por la columna "CLIENTE".
- Los nombres de columnas se normalizan internamente (por ejemplo MONTO VENTA
  CLIENTE, MONEDA 2, FECHA DIA/MES/ANO, TRM REFERENCIA, LINEA DE PRODUCTO).

## Configuracion clave
- `src/scripts/auth.js`: MSAL, roles y mapa correo -> nombre de Excel.
- `src/scripts/main.js`: carga de archivos, render y filtros.

## Ejecucion local
Abrir `index.html` en un navegador con acceso a internet. Requiere cuentas
corporativas para autenticacion en Microsoft 365.

## Notas
- No hay pruebas automatizadas.
- Si un ejecutivo no ve datos, el nombre del Excel debe coincidir con el mapa
  de correos o con el nombre del archivo en SharePoint.
