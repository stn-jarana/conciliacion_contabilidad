# Manual de Usuario — Sistema de Conciliación Bancaria

**Versión:** 2.0  
**Fecha de última actualización:** Agosto 2026  
**Empresas Soportadas:** STN, ITS, CMT del Sur, Dynamitex, DINSURA, Perú Commerce, INFOSUR, Thimble Sourcing (TST), Reforestadora Iñaupari, TECA Peruvian Group, DIONISO.

---

## 1. ¿Qué hace este programa?

El sistema automatiza el proceso de **conciliación bancaria multi-empresa**: cruza los movimientos del estado de cuenta bancario contra los registros exportados del sistema contable (Contanet), aplica reglas inteligentes de coincidencia automática y sugerencias, y genera archivos de trabajo en Excel y PDF para la revisión del especialista contable.

El flujo consta de **dos etapas**:

1. **Reporte Inicial (Generación del Excel de Trabajo)**:
   - Carga el extracto del banco y la exportación de contabilidad.
   - Aplica algoritmos de coincidencia automática de alta certeza (`MAR = X`).
   - Aplica reglas de detección de sugerencias y pre-marca automáticamente con **`X`** en la columna **`Conciliar`** para facilitar el trabajo del especialista.
   - Ordena los registros por grupos (Conciliados, Sugeridos, Solo Banco / Solo Conta ordenados por monto de mayor a menor, ITF y sin monto).
2. **Reporte Final (Consolidación)**:
   - Carga el Excel de trabajo revisado y modificado por el especialista.
   - Aplica las conciliaciones manuales (`# Operación a Conciliar`, `MAR = X` o sugerencias aprobadas).
   - Genera el reporte definitivo descargable en Excel.

---

## 2. Requisitos Previos e Instalación

- **Python 3.14+** o entorno configurado mediante `uv`.
- Archivos de entrada en formato Excel (`.xlsx`):
  - **Estado de Cuenta Bancario**: Archivo consolidado con las pestañas de cada empresa/moneda (ej. `STN DOL`, `STN SOL`, `ITS DOL`, `CMT SOL`, `P.COMMERCE`, `REF.IÑAPARI`, `TECA`, etc.).
  - **Reporte Contable**: Exportación de Contanet sin alterar.

### Ejecución rápida
Abra una consola en la carpeta del proyecto (`C:\Users\jarana\repositorio\proyect_conta\prueba_conta`) y ejecute:

```bash
uv run python main.py
```

---

## 3. Mapeo de Empresas y Monedas Soportadas

Al iniciar el sistema, podrá elegir entre las empresas configuradas y sus pestañas bancarias correspondientes:

| N° | Empresa en Menú | Monedas Disponibles | Pestaña de Banco en Excel (`sheet_bank`) |
|---|---|---|---|
| 1 | **Southern Textil Network (STN)** | Dólares (USD) / Soles (PEN) | `STN DOL` / `STN SOL` |
| 2 | **Integrated Textile Solutions (ITS)** | Dólares (USD) / Soles (PEN) | `ITS DOL` / `ITS SOL` |
| 3 | **CMT del Sur** | Dólares (USD) / Soles (PEN) | `CMT DOL` / `CMT SOL` |
| 4 | **Dynamitex** | Dólares (USD) / Soles (PEN) | `DYNAMITEX DOL` / `DYNAMITEX SOL` |
| 5 | **DINSURA** | Dólares (USD) / Soles (PEN) | `DINSURA DOL` / `DINSURA SOL` |
| 6 | **Perú Commerce** | Soles (PEN) | `P.COMMERCE` |
| 7 | **Inversiones Forestales del Sur (INFOSUR)** | Dólares (USD) / Soles (PEN) | `INFOSUR DOL` / `INFOSUR SOL` |
| 8 | **Thimble Sourcing / TST** | Soles (PEN) | `TST` |
| 9 | **Reforestadora Iñaupari** | Soles (PEN) | `REF.IÑAPARI` |
| 10 | **TECA Peruvian Group** | Soles (PEN) | `TECA` |
| 11 | **DIONISO** | Soles (PEN) | `DIONISO` |

---

## 4. Reglas del Sistema (Automáticas y Sugeridas)

Para optimizar la conciliación, el motor ejecuta las siguientes reglas en secuencia:

### A. Conciliación Automática Directa (`MAR = X`)
Estas filas ya quedan totalmente cerradas y conciliadas desde la generación del reporte inicial:
1. **Código Igual**: Mismo N° de operación y monto idéntico entre Banco y Contabilidad.
2. **N° Operación Contable en Glosa del Banco**: El N° de registro contable aparece dentro de la descripción del banco y el monto coincide.
3. **Monto + Fecha Única**: Mismo monto y misma fecha cuando esa combinación es única en el período.

### B. Sugerencias Inteligentes (Pre-marcadas con `X` en columna `Conciliar`)
Las filas sugeridas quedan agrupadas en el bloque de **Sugeridos** y traen puesta automáticamente una **`X`** en la columna `Conciliar`. Si el asesor no retira la `X`, se conciliarán automáticamente al generar el reporte final:
1. **Sugerencias 1 a N / N a 1**: Sumas de múltiples registros en un lado que coinciden con el total del otro lado.
2. **Sugerencia Factoring**: Si la glosa contable contiene la palabra **"factoring"** y los montos coinciden, se agrupan como sugerencia sin importar la diferencia de fechas.
3. **Sugerencia Cambio de Moneda (COMUS)**: Si la descripción del banco incluye **"COMU" / "COMUS"** y la glosa contable indica **"cambio"**, **"moneda"** o **"tc"**, y los montos coinciden.
4. **Sugerencia Monto Único (diferencia de días)**: Si un monto aparece **una sola vez en todo el extracto bancario y una sola vez en la contabilidad**, pero difiere en la fecha, el sistema los relaciona automáticamente como sugeridos.
5. **Sugerencia por Residuo Único**: Si tras aplicar todas las sugerencias anteriores queda exactamente 1 movimiento en Banco y 1 en Contabilidad sin sugerir y con montos idénticos, se emparejan automáticamente como sugerencia por descarte.
6. **Sugerencia por Diferencia de Comisiones**: Diferencias conocidas de tipo de cambio / comisiones bancarias prefijadas.

---

## 5. Trabajo del Especialista en la Hoja `Anexar1`

El archivo generado `Conciliacion_Inicial_AAAAMMDD_HHMMSS.xlsx` contiene la hoja principal de trabajo **`Anexar1`**.

### ¿Qué DEBE hacer el especialista?
- **Revisar las Sugerencias Pre-marcadas**:
  - Si la sugerencia es correcta: **No hacer nada**. Dejar la **`X`** en la columna **`Conciliar`** para que el sistema la concilie.
  - Si la sugerencia NO es correcta: **Borrar la `X`** de la columna **`Conciliar`**. De este modo, la fila permanecerá como partida abierta (solo en banco / solo en conta).
- **Conciliar Movimientos Manualmente**:
  - **Opción A (Recomendada)**: En la fila del Banco, escribir el N° de registro / operación contable en la columna **`# Operación a Conciliar`**.
  - **Opción B**: Colocar una **`X`** en la columna **`MAR`** tanto en la fila del banco como en la fila de contabilidad correspondiente.
- **Diferencia de Comisión**: Si hay diferencia de comisiones bancarias, anotarla en la columna **`DIF COMISON`**.

### ¿Qué NO debe hacer el especialista?
- **NO alterar el nombre de la hoja `Anexar1`** ni eliminarla.
- **NO modificar los nombres de las columnas ni mover las cabeceras**.
- **NO borrar los identificadores `# Operación2`** creados por el sistema para los sugeridos, ya que estos unen los pares.
- **NO insertar archivos alterados o manipulados** en la fase de carga inicial.

---

## 6. Estructura de Ordenamiento de la Hoja `Anexar1`

Para simplificar la lectura, la hoja `Anexar1` presenta la información organizada en este orden estricto:

1. **Conciliados Directos** (`MAR = X`): Movimientos cruzados automáticamente.
2. **Sugeridos**: Todos los pares y grupos sugeridos pre-marcados con `X`.
3. **Mov. solo en banco y Mov. solo en conta (UNIFICADOS)**:
   - Se muestran juntos en un solo bloque.
   - **Ordenados de MAYOR a MENOR por valor absoluto del monto**, facilitando identificar primero los montos más altos pendientes de justificar.
4. **Movimientos ITF**: Operaciones con impuesto a las transacciones financieras.
5. **Movimientos Sin Monto**: Registros en cero (0.00).

---

## 7. Generación del Reporte Final

Una vez guardados los cambios en `Conciliacion_Inicial_...xlsx`:

1. Vuelva a ejecutar `uv run python main.py`.
2. Elija la **misma empresa y moneda**.
3. Seleccione la **Opción 2 — Generar reporte final**.
4. Arrastre o ingrese la ruta del Excel trabajado.
5. El sistema procesará las confirmaciones y creará el archivo **`Conciliacion_Final_AAAAMMDD_HHMMSS.xlsx`** con las hojas definitivas:
   - `Conciliados`: Movimientos cerrados y conciliados.
   - `Solo_Banco`: Partidas abiertas únicamente en banco.
   - `Solo_Conta`: Partidas abiertas únicamente en contabilidad.
   - `Resumen_Final`: Saldos finales por TIPO/CÓDIGO.
   - `Estadisticas`: Métricas finales y diferencias residuales.

---

## 8. Preguntas Frecuentes y Solución de Problemas

- **¿Qué pasa si olvido borrar una 'X' en una sugerencia errónea?**  
  El sistema la tomará como aprobada y la conciliará. Es importante revisar la hoja `Anexar1` antes de correr el reporte final.
- **¿Por qué un movimiento único de diferente fecha ahora aparece en sugerencias?**  
  Por la regla de **Monto Único**, al ser la única transacción en todo el mes con esa cifra exacta en ambos lados, el sistema asume que corresponden al mismo hecho económico. Si no desea conciliarlo, retire la `X` de la columna `Conciliar`.
- **¿Qué ocurre si cargo un archivo del banco cuya pestaña no coincide?**  
  Verifique la tabla de la Sección 3 de este manual. Si la empresa o moneda seleccionada no encuentra su pestaña (ejemplo: `STN DOL`, `REF.IÑAPARI`), el sistema mostrará un mensaje de error indicando la falta de dicha pestaña.
