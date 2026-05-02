# Biblioteca Yugioh

App de escritorio para gestionar una colección física de cartas Yu-Gi-Oh almacenadas en álbumes.

## Características

- **Cartas por tipo** — Monstruos, Rituales, Fusión, Sincronía, Enlace, XYZ, Péndulo, Magia, Trampas
- **Huecos vacíos** — los espacios sin carta del álbum se conservan visualmente
- **Repetidos** — detecta automáticamente si una carta ya existe en el álbum al añadirla
- **Decks** — mueve cartas del álbum a un deck; marca si queda copia en el álbum
- **Pendientes** — lista de cartas deseadas; aparecen en la búsqueda general
- **Álbumes** — gestiona álbumes y número de páginas; genera ubicaciones automáticamente
- **Estadísticas** — total de colección, cartas tenidas, cartas que faltan
- **Gráficas** — distribución por tipo (barras + tarta)
- **Importar / Exportar Excel** — lee y escribe el Excel original con backup automático

## Requisitos

```
Python 3.10+
customtkinter
openpyxl
matplotlib
```

Instalar dependencias:

```bash
pip install customtkinter openpyxl matplotlib
```

## Uso

```bash
python app.py
```

Al arrancar, si la base de datos está vacía busca automáticamente `Cartas YUGIOH.xlsx` en la misma carpeta para importar.

## Estructura del Excel esperado

| Hoja | Contenido |
|------|-----------|
| Monstruos, Fusion, Sincronia… | Dos bloques por fila (izquierda cols 1-6, derecha cols 8-13) |
| Rituales | Un único bloque por fila |
| Decks | Grupos de 3 decks separados por filas de cabecera |
| Estadistica | Celda O26 = total de cartas de la colección completa |

## Base de datos

SQLite local (`biblioteca.db`) — se crea automáticamente.  
No se sube al repositorio.
