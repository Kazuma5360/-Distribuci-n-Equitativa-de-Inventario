# 📦 Distribución Equitativa de Inventario

Script de Python que distribuye el inventario disponible de forma **proporcional y justa** entre todos los pedidos pendientes por SKU, generando un reporte detallado en Excel con 4 hojas.

---

## 🚀 ¿Qué hace?

Cuando el inventario no es suficiente para cubrir todos los pedidos, este script:

1. **Lee** un archivo de demandas/pedidos y un archivo de inventario
2. **Combina** ambas fuentes por SKU
3. **Distribuye** el stock disponible proporcionalmente — cada pedido recibe una parte según lo que solicitó respecto al total de la demanda
4. **Maneja los sobrantes** de forma inteligente: las unidades restantes del redondeo se asignan a los pedidos con mayor residuo decimal
5. **Exporta** un archivo Excel limpio con 4 hojas

---

## 📊 Excel de salida — 4 Hojas

| Hoja | Descripción |
|---|---|
| `Distribución Completa` | Todos los pedidos con su cantidad asignada, % atendido y faltante |
| `Resumen por SKU` | Totales agrupados por producto |
| `Pedidos con Faltante` | Solo los pedidos que no se cubrieron al 100% |
| `Estadísticas Generales` | Números resumen globales |

---

## 🧠 Lógica de Distribución

```
ratio = stock_disponible / total_pendiente

asignado[i] = piso(pendiente[i] * ratio)
```

Tras el redondeo hacia abajo, las unidades sobrantes se distribuyen una a una a los pedidos con el **mayor residuo decimal** (asignación justa de sobrantes). Cualquier excedente final va al último pedido sin pasarse de lo que solicitó.

### Ejemplo

| Pedido | Solicitado | Disponible | Distribuido | % Atendido |
|---|---|---|---|---|
| A | 100 | — | 50 | 50% |
| B | 60 | 150 | 30 | 50% |
| C | 140 | — | 70 | 50% |

> Total solicitado: 300 · Disponible: 150 · Ratio: 50%

---

## 🗂️ Estructura del Proyecto

```
📁 distribucion-equitativa/
├── equitable_distribution.py   # Script principal
├── requirements.txt            # Dependencias
└── README.md                   # Estás aquí
```

---

## ⚙️ Configuración

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/distribucion-equitativa.git
cd distribucion-equitativa
```

### 2. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 3. Configurar tus archivos

Abre `equitable_distribution.py` y actualiza la sección de configuración al inicio:

```python
DEMAND_FILE       = "demands.xlsx"    # Ruta a tu archivo de demandas/pedidos
DEMAND_SHEET      = "Sheet1"          # Nombre de la hoja dentro del archivo
DEMAND_HEADER_ROW = 2                 # Fila donde están los encabezados (índice desde 0)

INVENTORY_FILE    = "inventory.xls"   # Ruta a tu archivo de inventario
OUTPUT_FILE       = "Resultado_Distribucion_Equitativa.xlsx"
```

También actualiza `COLUMNS_TO_DROP` con las columnas que quieras eliminar de tu archivo de demandas.

### 4. Ejecutar el script

```bash
python equitable_distribution.py
```

---

## 📋 Requisitos

- Python 3.8+
- pandas
- numpy
- openpyxl
- xlrd (para archivos `.xls`)

Consulta `requirements.txt` para las versiones exactas.

---

## 📁 Formato de los Archivos de Entrada

### Archivo de demandas (`.xlsx`)
Debe contener como mínimo:

| Columna | Descripción |
|---|---|
| `SKU` | Identificador del producto |
| `PENDING_QTY` | Unidades pendientes por surtir |
| `TYPE` | Tipo de pedido (el script filtra por `FRACTIONED`) |

### Archivo de inventario (`.xls`)
Debe contener como mínimo:

| Columna | Descripción |
|---|---|
| `SKU` (o `PRODUCT_ID`) | Identificador del producto |
| `AVAILABLE` | Unidades disponibles |
| `LOT_ID` | Identificador de lote |
| `WAREHOUSE_ID` | Identificador de bodega |

> ⚠️ Renombra tus columnas para que coincidan con estos nombres, o actualiza la sección de mapeo de columnas en el script.

---

## 🤝 Contribuciones

Los pull requests son bienvenidos. Para cambios importantes, abre primero un issue para discutir lo que te gustaría modificar.

---

## 📄 Licencia

[BSD 3-Clause](https://choosealicense.com/licenses/bsd-3-clause/)

Este proyecto está bajo la licencia BSD 3-Clause. Esto significa que puedes usar, modificar y distribuir este código libremente, pero **no puedes usar el nombre del autor para promocionar productos derivados** sin permiso previo por escrito.

---

*Desarrollado con Python · pandas · numpy · openpyxl*
