# Handoff — Biblioteca Yugioh

## Estado actual

App de escritorio funcional en producción. Último exe: `v1.1` (release en GitHub).

Repositorio: https://github.com/XxbazingaxX/APP_yugioh_DB  
Rama principal: `master`  
Último commit: `f0afd93`

---

## Archivos clave

| Archivo | Función |
|---|---|
| `app.py` | UI principal (CustomTkinter). Toda la lógica de pantallas y eventos. |
| `db.py` | Capa SQLite. Todas las queries. DB path resuelto por `sys.executable` en exe. |
| `importer.py` | Lector de Excel → BD. Lógica de bloques L/R independientes. |
| `exporter.py` | Escritor BD → Excel. Reproduce estructura dual-bloque. |
| `biblioteca.db` | BD SQLite. Va junto al exe. NO se sube a GitHub (`.gitignore`). |
| `Cartas YUGIOH.xlsx` | Excel fuente. NO se sube a GitHub. |

---

## Estructura de la BD

```sql
cartas      -- id, tipo, cant, nombre, ubicacion, ubicacion_copia, cant2, repetidos, vacio, sort_order
decks       -- id, cant, nombre, copia_extra, repetidos, deck_nombre
pendientes  -- id, nombre, tipo, notas
albums      -- id, nombre, paginas
config      -- clave, valor  (total_cartas = O26 del Excel de Estadistica)
```

---

## Lógica de importación (importer.py)

El Excel tiene hojas por tipo de carta con **dos bloques por fila**:
- **Bloque izquierdo**: columnas A–F (índices 0–5)
- **Bloque derecho**: columnas H–M (índices 7–12)
- **Excepción**: `Rituales` es bloque único (solo columna izquierda)

### Algoritmo clave
1. Abre el workbook sin `read_only` para detectar rangos de tablas Excel (`ws.tables`) → `_get_table_bounds()` devuelve el `min_row` de cabecera (+1 para saltarla).
2. Lee el sheet completo desde `min_row` **sin límite de `max_row`** — las tablas Excel no siempre cubren todos los datos.
3. Escanea todas las filas para calcular `first_l/last_l` (rango real del bloque izquierdo) y `first_r/last_r` (rango real del bloque derecho) **por separado**.
4. Para cada fila, solo inserta el bloque izquierdo si `first_l <= i <= last_l`, ídem derecho. Esto evita huecos espurios cuando una columna termina antes que la otra.
5. Filas completamente vacías dentro del rango → `vacio=1` (hueco físico en el álbum).
6. `_trim_trailing_vacios()` elimina huecos al final de cada tipo (por si acaso).

### Hoja Decks
Cada fila de cabecera (`row[1] == "Nombre"`) incrementa `deck_section`. Hay 3 bloques por fila → `deck_nombre = f"Deck {(section-1)*3 + block_idx + 1}"`.

### Hoja Estadistica
Lee celda `O26` → guarda en `config` como `total_cartas`.

---

## Comportamiento del exe (PyInstaller --onefile)

- `sys.frozen = True` cuando corre como exe.
- `db.py` y `app.py` resuelven paths relativos a `os.path.dirname(sys.executable)`.
- La BD (`biblioteca.db`) debe estar en la misma carpeta que el exe.
- El Excel (`Cartas YUGIOH.xlsx`) también puede ir junto al exe para que Reimportar lo encuentre automáticamente.
- `sys.exit(0)` al cerrar la ventana para matar todos los procesos de PyInstaller.

---

## Compilar exe

```bash
pyinstaller --onefile --windowed --name "BibliotecaYugioh" \
  --add-data "PATH_TO/customtkinter;customtkinter" \
  --add-data "PATH_TO/matplotlib/mpl-data;matplotlib/mpl-data" \
  --hidden-import customtkinter \
  --hidden-import matplotlib.backends.backend_tkagg \
  --hidden-import matplotlib.backends._backend_tk \
  app.py
```

Ajustar `PATH_TO` al directorio de site-packages de Python en la máquina de compilación.

---

## Reimportar

- Borra y recrea `cartas` y `decks`.
- **Conserva**: `albums`, `pendientes`, `config`.
- Muestra aviso previo al usuario.
- Ediciones manuales de cartas (ubicacion, cant2…) no guardadas en Excel **se pierden**.

---

## Pendiente / posibles mejoras

- El spec de PyInstaller (`BibliotecaYugioh.spec`) tiene rutas hardcoded a la máquina de compilación — regenerar con `--add-data` en cada máquina nueva o parametrizar.
- El exportador no regenera las tablas Excel formales, solo escribe celdas — si se reimporta el Excel exportado, las tablas no estarán definidas y `_get_table_bounds()` devolverá vacío (el fallback `min_row=2` lo cubre correctamente).
- El buscador global (sidebar) fue añadido en el pull del GitHub pero puede tener comportamiento no verificado en todos los flujos.
