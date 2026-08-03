# Sistema de Conciliación Bancaria Multi-Empresa

Sistema automatizado en Python para el procesamiento, cruce, sugerencia inteligente y conciliación de movimientos bancarios frente a registros contables de Contanet para un grupo de 11 empresas.

---

## Características Principales

- **Soporte Multi-Empresa y Moneda**: Manejo automático de pestañas de extractos bancarios en Soles (PEN) y Dólares (USD) para:
  - Southern Textil Network (STN)
  - Integrated Textile Solutions (ITS)
  - CMT del Sur
  - Dynamitex
  - DINSURA
  - Perú Commerce
  - Inversiones Forestales del Sur (INFOSUR)
  - Thimble Sourcing / TST
  - Reforestadora Iñaupari
  - TECA Peruvian Group
  - DIONISO
- **Conciliación Automática de Alta Certidumbre**:
  - Cruce por código de operación igual.
  - Extracción de N° de operación contable dentro de la descripción bancaria.
  - Coincidencia de monto + fecha única.
- **Sugerencias Inteligentes Pre-Marcadas (`X` en columna `Conciliar`)**:
  - Operaciones de **Factoring**.
  - Operaciones de **Cambio de Moneda (COMUS)**.
  - **Montos Únicos** en el período (con diferencia de días).
  - Sumas N a 1 y 1 a N.
  - Parejas resultantes por descarte.
  - Diferencias por comisiones bancarias.
- **Flujo en Dos Etapas**:
  1. `Reporte Inicial`: Genera un Excel interactivo de trabajo (`Conciliacion_Inicial_...xlsx`) y un PDF de resumen ejecutivo.
  2. `Reporte Final`: Procesa la hoja `Anexar1` con las confirmaciones o ediciones del especialista y genera el cierre definitivo (`Conciliacion_Final_...xlsx`).
- **Ordenamiento Inteligente**:
  - Conciliados directos al inicio.
  - Sugeridos pre-marcados a continuación.
  - Bloque unificado de **Partidas Abiertas (Solo Banco / Solo Conta)** ordenados descendentemente **de mayor a menor monto**.

---

## Requisitos e Instalación

Requiere **Python 3.14+** y administrador de proyectos **`uv`**.

```bash
# Clonar o ubicarse en la carpeta del proyecto
cd C:\Users\jarana\repositorio\proyect_conta\prueba_conta

# Sincronizar dependencias con uv
uv sync
```

---

##  Uso del Sistema

Ejecute la aplicación mediante el CLI interactivo:

```bash
uv run python main.py
```

### Flujo de Ejecución:
1. **Seleccionar Empresa y Moneda**: Elija la entidad y la divisa a procesar.
2. **Generar Reporte Inicial (Opción 1)**:
   - Ingrese la ruta del Excel del Banco (ej. extracto consolidado).
   - Ingrese la ruta del Excel de Contabilidad (exportación Contanet).
   - El sistema generará el archivo de trabajo `Conciliacion_Inicial_...xlsx`.
3. **Revisión del Especialista (Excel)**:
   - Abra el archivo `Conciliacion_Inicial_...xlsx` en la hoja **`Anexar1`**.
   - Para sugerencias aceptadas: Deje la **`X`** pre-marcada en la columna `Conciliar`.
   - Para sugerencias rechazadas: Borre la **`X`**.
   - Para vinculaciones manuales: Ingrese el código en `# Operación a Conciliar` o coloque **`X`** en `MAR`.
4. **Generar Reporte Final (Opción 2)**:
   - Seleccione la opción 2 e ingrese la ruta del Excel trabajado.
   - Obtendrá el resultado consolidado en `Conciliacion_Final_...xlsx`.

---

##  Estructura del Proyecto

```
prueba_conta/
├── main.py              # Menú CLI interactivo y flujo de control por empresa/moneda
├── reporte.py           # Algoritmos de coincidencia, sugerencias, ordenamiento y reporte inicial
├── reporte_final.py     # Procesamiento de respuestas del especialista y generación del reporte final
├── conciliacion.py      # Módulos auxiliares de lectura y procesamiento de estructuras
├── pdf_resumen.py       # Generador de reportes en PDF (ReportLab)
├── MANUAL_USUARIO.md    # Manual detallado para el especialista contable
├── pyproject.toml       # Configuración del proyecto y dependencias
└── README.md            # Documentación general técnica y funcional
```

---

## Documentación Completa

Para conocer todos los detalles operativos, reglas de negocio y preguntas frecuentes, consulte el [Manual de Usuario](file:///C:/Users/jarana/repositorio/proyect_conta/prueba_conta/MANUAL_USUARIO.md).
