# 🔍 Buscador de Ceses

Busca nombres de personas (columna **C**) dentro del texto de publicaciones (columna **B**) y genera un archivo Excel con los identificadores de publicación donde hubo coincidencias (columna **A**).

> La búsqueda es inteligente: ignora mayúsculas, minúsculas, acentos, puntuación y diferencias de espacios.

---

## 📋 Formato del Excel de Entrada (`.xlsx`)

El archivo debe contener las siguientes 3 columnas (los datos inician en la fila 2):

| Columna | Campo                               | Descripción                                                                                                                                       |
| :-----: | :---------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------ |
|  **A**  | **Identificador de la publicación** | ID de la publicación (ej. exportación de Facebook). El script le antepone automáticamente el prefijo de la página: `1013585111843916_`           |
|  **B**  | **Título**                          | Texto de la publicación donde se buscará                                                                                                          |
|  **C**  | **Lista de ceses**                  | Nombre de la persona a rastrear (ej. del reporteador)                                                                                             |

---

## 🚀 Instalación

Solo requiere **Python 3.8+** y la librería `openpyxl`.

1. **Crear el entorno virtual**:

   ```bash
   python -m venv venv
   ```

2. **Instalar la dependencia**:
   - **Windows:**
     ```cmd
     .\venv\Scripts\python.exe -m pip install openpyxl
     ```
   - **macOS / Linux:**
     ```bash
     ./venv/bin/python -m pip install openpyxl
     ```

---

## 💻 Uso

### 1. Búsqueda directa (genera salida automática)

Crea un archivo terminado en `_coincidencias.xlsx` en la misma carpeta:

```bash
./venv/bin/python buscador-ceses.py "archivo.xlsx"
```

### 2. Especificar nombre o ruta del archivo de salida

```bash
./venv/bin/python buscador-ceses.py "archivo.xlsx" -o "coincidencias.xlsx"
```

### 3. Indicar una hoja específica

```bash
./venv/bin/python buscador-ceses.py "archivo.xlsx" --sheet "Hoja1"
```

---

## 📊 Resultado

El archivo generado contendrá la hoja **"Coincidencias"** con formato listo para filtrar:

| Columna A                           | Columna B             |
| :---------------------------------- | :-------------------- |
| **Identificador de la publicación** | **Nombre encontrado** |
