

# ✉️ Automatizador de Envío de Correos a Transportistas con Detalle

**Con generación de archivos filtrados (Anticipos / Saldos – Seco / Líquidos)**  

---

## 📌 ¿Qué hace este sistema?

Este automatizador está dividido en **dos etapas principales**:

1. **Filtrado inteligente de programaciones de pagos** desde un archivo maestro `Plantilla.xlsx`, dependiendo del tipo de operación y producto.
2. **Envío automatizado de correos electrónicos personalizados a cada transportista**, con detalle en formato tabla HTML y destinatarios cargados desde `remitentes.xlsx`.

---

## 🔍 Etapa 1: Filtrado de planillas

### 🧠 ¿Cómo funciona?

El usuario elige mediante inputs:

- `ANTICIPO` o `SALDO`
- `SECO` o `LÍQUIDOS`

Según la selección:
- Se usa la hoja correspondiente (`Detalles (Seco)` o `Detalles (Liquido)`)
- Se filtra por la **fecha más reciente** en la columna clave (`FECHA PAGO ANTICIPO`, `FECHA SALDO`, etc.)
- Se extraen solo las columnas necesarias
- Se guarda un nuevo Excel filtrado con nombre automático como:

datos_filtrados_anticipo_seco.xlsx
datos_filtrados_saldo_liquido.xlsx



### 📁 Archivos involucrados:

- `Plantilla.xlsx` → Archivo base con los datos brutos  
- `datos_filtrados_*.xlsx` → Salida generada automáticamente  

---

## 📤 Etapa 2: Envío de correos desde Outlook

Una vez generado el Excel filtrado:

- Se cruza con el archivo `remitentes.xlsx` que contiene correos electrónicos por transportadora.
- Se agrupan los datos por transportista.
- Se genera un **correo en formato HTML**, con tabla detallada, y se guarda en **borradores de Outlook**.

### ✉️ Estructura del correo:

- Asunto:  
  `PAGO TRANSPORTISTA 09/10/2025 - ANTICIPO - ALCON LOGISTICS`

- Cuerpo:  
  Incluye saludo, tabla HTML con detalles y despedida estandarizada.

- Para y CC:  
  Automáticamente definidos desde el archivo `remitentes.xlsx`.

---

## 🛠 Requisitos técnicos

- Python 3.11 o superior  
- Librerías:  
  `pandas`, `openpyxl`, `win32com.client` (solo en Windows con Outlook)

---

## 📒 Archivos requeridos

| Archivo              | Descripción                                    |
|----------------------|------------------------------------------------|
| `Plantilla.xlsx`     | Base de datos con todos los movimientos        |
| `remitentes.xlsx`    | Transportadoras y sus correos                  |
| `datos_filtrados_*.xlsx` | Se generan automáticamente por el script  |

---

## ✅ Ventajas

- Evita errores manuales al copiar datos de transportistas
- Asegura que los correos se envíen con formato uniforme
- Permite revisar todo en borradores antes del envío
- Fácil de modificar para nuevos tipos de pagos o columnas

---

## 🧩 Posibles mejoras futuras

- Agregar botón o GUI para evitar inputs manuales  
- Agregar validación cruzada de facturas  
- Permitir enviar directamente (no solo guardar en borradores)

---

