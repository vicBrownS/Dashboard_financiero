# 📊 **README: Dashboard Financiero en Excel con OpenPyXL & Pandas** 🚀  

Este proyecto genera un **Dashboard Financiero Automatizado** en Excel utilizando **Python, OpenPyXL y Pandas**. Analiza y visualiza datos financieros de diferentes países, generando **resúmenes estadísticos, gráficos avanzados y formatos personalizados** para facilitar la interpretación de la información.

Además del script principal, el proyecto incluye un **notebook de ejecución**, que permite ejecutar el código paso a paso con **comentarios más detallados** sobre la elección de los datos y su tratamiento con Pandas.

---

## **📌 Características del Proyecto**
✅ **Carga de Datos**: Importa un archivo Excel con información financiera.  
✅ **Análisis con Pandas**: Calcula métricas clave como Ventas, Ganancias, Crecimiento Anual y Márgenes de Beneficio.  
✅ **Automatización de Reportes**: Organiza los datos en una hoja de resúmenes dentro del mismo archivo Excel.  
✅ **Formato Profesional en Excel**: Ajuste de columnas, bordes, formato condicional con escalas de colores.  
✅ **Gráficos Dinámicos con OpenPyXL**: Generación de gráficos de Barras, Pie y Scatter dentro de Excel.  
✅ **Notebook de ejecución**: Ejecuta el código poco a poco con explicaciones más detalladas sobre la manipulación de datos.  
✅ **Salida Personalizada**: Genera un archivo final con todas las métricas calculadas y visualizaciones.  

---

## **📂 Estructura del Proyecto**
```
📂 Proyecto-Dashboard
│── 📂 data
│   ├── Financial Sample.xlsx  # Archivo Excel de entrada y salida
│── main.py                    # Script principal (automatizado)
│── Notebook_ejecucion.ipynb    # Notebook con ejecución paso a paso
│── README.md                   # Archivo de documentación
```

El script genera un archivo Excel con tres hojas:
1️⃣ **Datos Originales**: Datos financieros originales cargados desde el archivo.  
2️⃣ **Resumenes**: Cálculos financieros (Ventas Brutas, Ganancias, Crecimiento Anual, etc.).  
3️⃣ **Graficos**: Visualizaciones automáticas con OpenPyXL.

---

## **📌 Uso de OpenPyXL en el Proyecto**
El proyecto **integra OpenPyXL** para manipular y mejorar los archivos Excel de la siguiente manera:

### 📊 **1️⃣ Generación de Reportes y Estadísticas**
- Creación de una nueva hoja llamada **"Resumenes"** con los cálculos financieros.  
- Escritura de datos en celdas con formato numérico redondeado.  
- Ajuste automático del ancho de las columnas para mejorar la visibilidad.  

### 🎨 **2️⃣ Aplicación de Formatos y Estilos**
- **Formato Condicional** → Resalta valores bajos y altos en colores rojo, amarillo y verde.  
- **Bordes** → Se aplican bordes a todas las celdas para mejorar la estética.  

### 📈 **3️⃣ Creación de Gráficos dentro de Excel**
Se generan automáticamente los siguientes gráficos dentro del archivo Excel:
- **Gráfico de Barras Apiladas** → Comparación de **Ventas Brutas vs Ganancias Totales** por país.  
- **Gráfico Circular** → Distribución del **Crecimiento Anual** por país.  
- **Gráfico de Barras Individual** → **Promedio de Ventas** por país.  
- **Gráfico de Dispersión (Scatter Chart)** → Relación entre **Ventas Brutas y Ganancias Totales**.

Los gráficos son generados con:
```python
from openpyxl.chart import BarChart, PieChart, ScatterChart, Reference
```

---

## **📌 Notebook de Ejecución**
Además del script principal **`main.py`**, este proyecto incluye un **notebook en Jupyter** (`execution_notebook.ipynb`), donde el código se ejecuta de manera **progresiva y con explicaciones más detalladas**.  

El notebook es útil para:
🔹 Comprender paso a paso cómo se procesan los datos con **Pandas**.  
🔹 Analizar decisiones como la limpieza y transformación de los datos antes de exportarlos.  
🔹 Visualizar los resultados intermedios antes de generarlos en Excel.  
🔹 Personalizar los cálculos y gráficos según las necesidades del usuario.  

---

## **📌 Instalación y Uso**
### **1️⃣ Instalar Dependencias**
Ejecuta el siguiente comando para instalar las librerías necesarias:
```bash
pip install openpyxl pandas jupyter
```

### **2️⃣ Ejecutar el Script Automático**
```bash
python main.py
```
Esto generará un nuevo archivo Excel con el Dashboard completo.

### **3️⃣ Ejecutar el Notebook Paso a Paso**
Para ejecutar el notebook de forma interactiva:
```bash
jupyter notebook
```
Luego, abre **`Notebook_ejecucion.ipynb`** y ejecuta las celdas una por una para explorar el proceso de análisis y generación del dashboard.

---

## **📌 Posibles Mejoras**
🔹 Modularizar el código en funciones (`calcular_estadisticas()`, `formatear_excel()`, `crear_graficos()`).  
🔹 Agregar gráficos adicionales como **líneas de tendencia y diagramas de dispersión avanzados**.  
🔹 Mejorar la interacción con el usuario, permitiendo **cambiar parámetros sin modificar el código**.

---

## **📌 Conclusión**
Este proyecto **automatiza la creación de reportes financieros en Excel**, combinando el poder de **Pandas para el análisis de datos** y **OpenPyXL para la manipulación y visualización en Excel**. 🚀  

Además, el **notebook de ejecución** permite entender cada paso del proceso de manera más detallada, facilitando la personalización y exploración de los datos.  

---

📌 **Autor:** Victor Brown Sogorb   

¡Si te ha gustado este proyecto, no dudes en darle una ⭐ en GitHub! 😊🚀
