# Manual de Usuario — Sistema de Conciliación Bancaria

**Versión:** 1.0  
**Fecha:** Julio 2026  
**Empresa:** Southern Textil

---

## ¿Qué hace este programa?

El sistema automatiza el proceso de **conciliación bancaria**: cruza los movimientos del estado de cuenta bancario contra los registros del sistema contable (Contanet), identifica las diferencias y genera reportes en Excel y PDF listos para revisión del especialista contable.

El flujo completo tiene **dos etapas**:

1. **Reporte Inicial** → El programa carga los archivos y genera un Excel de trabajo.
2. **Reporte Final** → El especialista completa el Excel y el programa genera el informe definitivo.

---

## Requisitos previos

Antes de ejecutar el programa, asegúrese de tener:

- **Python 3.14** o superior instalado.
- Las dependencias del proyecto instaladas (ver sección de instalación).
- Los archivos Excel de entrada:
  - Estado de cuenta del **banco** (`.xlsx`)
  - Reporte de **contabilidad** exportado de Contanet (`.xlsx`)

### Instalación (primera vez)

Abra una terminal en la carpeta del proyecto y ejecute:

```
uv sync
```

Esto instalará automáticamente todas las librerías necesarias (`pandas`, `openpyxl`, `reportlab`, etc.).

---

## Cómo ejecutar el programa

Abra una terminal en la carpeta del proyecto:

```
cd C:\Users\jarana\repositorio\proyect_conta\prueba_conta
```

Luego ejecute:

```
uv run python main.py
```

---

## Paso a paso — Flujo completo

### PASO 1: Seleccionar empresa y moneda

Al iniciar el programa, aparece el menú principal:

```
============================================================
  CONCILIACION BANCARIA
============================================================

Seleccione la empresa:

  1. Southern Textil
  0. Salir

Opcion:
```

- Ingrese el número de la empresa y presione **Enter**.
- A continuación, seleccione la moneda:

```
Empresa: Southern Textil
Seleccione la moneda:

  1. Dolares (USD)
  2. Soles (PEN)   [no disponible]
  0. Volver

Opcion:
```

> **Nota:** Las opciones marcadas como `[no disponible]` aún no están habilitadas. Solo puede seleccionar las que no tienen esa etiqueta.

---

### PASO 2: Seleccionar acción

Después de elegir empresa y moneda, el programa pregunta qué desea hacer:

```
------------------------------------------------------------
  Que desea hacer?

  1. Generar reporte inicial
     (carga archivos banco + contabilidad y genera el Excel de trabajo)

  2. Generar reporte final
     (carga el Excel ya trabajado por el especialista)

  0. Salir

Opcion:
```

---

## Opción 1 — Generar reporte inicial

Use esta opción la **primera vez** que procesa un período. Necesita los dos archivos de origen.

### 1.1 Ingresar rutas de archivos

El programa solicita la ruta de cada archivo:

```
Ruta del estado de cuenta del BANCO (.xlsx): 
Ruta del reporte de CONTABILIDAD (.xlsx)  : 
```

Puede escribir la ruta directamente o **arrastrar el archivo** desde el Explorador de Windows hacia la terminal. Ejemplos de ruta válida:

```
C:\Users\jarana\Documentos\banco_julio2026.xlsx
C:\Users\jarana\Documentos\contanet_julio2026.xlsx
```

> **Importante:** Los 2 archivos, el del banco y el de contanet, no deben estar manipulados para la revisión inicial, deben agregarse tal como han sido entregados para que el sistema ubique los datos correctamente. 

### 1.2 Archivos generados

Si todo es correcto, el programa genera **dos archivos** en la carpeta del proyecto:

| Archivo | Descripción |
|---|---|
| `Conciliacion_Inicial_AAAAMMDD_HHMMSS.xlsx` | Libro Excel de trabajo con 6 hojas |
| `Conciliacion_Inicial_AAAAMMDD_HHMMSS.pdf` | Resumen ejecutivo en PDF |

El Excel contiene las siguientes hojas:

| Hoja | Color | Contenido |
|---|---|---|
| **BANCO** | Azul oscuro | Todos los movimientos del banco |
| **CONTANET** | Verde oscuro | Todos los registros contables |
| **Tab_Banco** | Azul medio | Tabla simplificada del banco (para cruce) |
| **Tab_Contanet** | Verde medio | Tabla simplificada de contabilidad (para cruce) |
| **Anexar1** | Morado | ⭐ Hoja de trabajo del especialista — banco y conta apilados |
| **Resumen** | Naranja | Totales por tipo/código con diferencias |

---

## PASO INTERMEDIO — Trabajo del especialista en el Excel

Antes de generar el reporte final, el **especialista contable** debe abrir el archivo `Conciliacion_Inicial_...xlsx` y trabajar en la hoja **Anexar1**.

### ¿Qué debe completar el especialista?

En la hoja `Anexar1` hay filas del banco y filas de contabilidad intercaladas. El especialista debe vincular manualmente los pares que corresponden:

| Columna | Qué completar |
|---|---|
| **MAR** | Escribir `X` en la fila del banco Y en la fila de contabilidad que se corresponden |
| **# Operación2** | En la fila de contabilidad, colocar el número de operación del banco correspondiente |
| **DIF COMISON** | Si hay diferencia por comisión bancaria, registrarla aquí |
| **COD_SUB** | Código auxiliar de subcuenta si aplica |

### Criterio de vinculación

Dos filas son un "par conciliado" cuando:
- El **monto** del banco y de contabilidad coinciden (se acepta diferencia de hasta S/. 0.01).
- La **fecha** es igual o difiere como máximo 1 día.
- Se coloca **`X`** en la columna `MAR` de ambas filas.

Las filas que el especialista **no marque** quedarán como partidas abiertas (solo en banco o solo en contabilidad).

---

## Opción 2 — Generar reporte final

Use esta opción **después** de que el especialista haya completado la hoja Anexar1.

### 2.1 Ingresar la ruta del Excel trabajado

```
Cargue el Excel de conciliacion ya trabajado por el especialista.
(Es el archivo 'Conciliacion_Inicial_...' con la hoja Anexar1 completada)

Ruta del Excel trabajado (.xlsx): 
```

Ingrese la ruta del archivo `Conciliacion_Inicial_...xlsx` que el especialista ya completó.

### 2.2 Archivos generados

El programa genera un nuevo archivo:

```
Conciliacion_Final_AAAAMMDD_HHMMSS.xlsx
```

Con las siguientes hojas:

| Hoja | Color | Contenido |
|---|---|---|
| **Conciliados** | Azul oscuro | Pares banco↔conta confirmados por el especialista |
| **Solo_Banco** | Naranja | Movimientos bancarios sin registro contable |
| **Solo_Conta** | Verde oscuro | Registros contables sin movimiento bancario |
| **Resumen_Final** | Morado | Totales por tipo/código con diferencias residuales |
| **Estadisticas** | Gris | Resumen general numérico de la conciliación |

### 2.3 Resumen en pantalla

El programa también muestra en consola un resumen como este:

```
============================================================
  RESULTADO DE CONCILIACION FINAL
============================================================
  Movimientos banco (total)              :              45
  Movimientos conta (total)              :              42
  Pares conciliados por especialista     :              38
  Monto conciliado (banco)               :      125,430.00
  Movimientos solo en banco              :               7
  Monto solo en banco                    :        3,200.50
  Movimientos solo en contabilidad       :               4
  Monto solo en contabilidad             :          890.00
  Total banco                            :      128,630.50
  Total contabilidad                     :      125,540.00
  Diferencia neta                        :        3,090.50

  [OK] Reporte final generado: Conciliacion_Final_20260716_142500.xlsx
```

---

## Mensajes de error frecuentes

| Mensaje | Causa | Solución |
|---|---|---|
| `X No se encontro el archivo: ...` | La ruta ingresada no existe | Verifique la ruta o arrastre el archivo a la terminal |
| `X El archivo debe ser Excel (.xlsx o .xls)` | Se ingresó un archivo de otro tipo | Use solo archivos Excel |
| `X La opcion '...' aun no esta disponible` | Moneda no habilitada | Seleccione una opción sin la etiqueta `[no disponible]` |
| `X Error: El archivo '...' no contiene la hoja 'Anexar1'` | Se cargó el archivo equivocado en el reporte final | Cargue el archivo `Conciliacion_Inicial_...xlsx` (no el final) |

---

## Flujo resumido (diagrama)

```
[Inicio]
    │
    ▼
Seleccionar empresa y moneda
    │
    ▼
┌─────────────────────────────────────────────────┐
│  Opción 1: Reporte Inicial                      │
│                                                 │
│  1. Ingresar archivo banco (.xlsx)              │
│  2. Ingresar archivo contabilidad (.xlsx)       │
│  3. El programa genera:                         │
│     → Conciliacion_Inicial_...xlsx              │
│     → Conciliacion_Inicial_...pdf               │
└─────────────────────────────────────────────────┘
    │
    ▼
[Especialista trabaja en la hoja Anexar1]
Marca con X los pares banco↔conta
    │
    ▼
┌─────────────────────────────────────────────────┐
│  Opción 2: Reporte Final                        │
│                                                 │
│  1. Ingresar el Excel ya completado             │
│  2. El programa genera:                         │
│     → Conciliacion_Final_...xlsx                │
└─────────────────────────────────────────────────┘
    │
    ▼
[Fin]
```

---

## Preguntas frecuentes

**¿Puedo ejecutar el programa varias veces sobre el mismo período?**  
Sí. Cada ejecución genera un archivo nuevo con timestamp distinto. Los archivos anteriores no se sobreescriben.

**¿El PDF del reporte inicial es el mismo que el Excel?**  
No. El PDF es un **resumen ejecutivo** compacto con los totales principales, útil para compartir. El Excel contiene todos los datos de trabajo.

**¿Qué hago si el balance no concilia (diferencia neta ≠ 0)?**  
Revise la hoja `Solo_Banco` y `Solo_Conta` del reporte final. Esas son las partidas pendientes de explicar. Pueden ser comisiones, depósitos en tránsito o errores de registro.

**¿Puedo agregar otra empresa al sistema?**  
Sí. Un desarrollador puede agregar la configuración en el archivo `main.py` dentro del arreglo `EMPRESAS`, especificando el nombre de la hoja del banco y las filas a saltar.
