import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "biblioteca.db")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS cartas (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo             TEXT    NOT NULL,
            cant             INTEGER DEFAULT 0,
            nombre           TEXT,
            ubicacion        TEXT,
            ubicacion_copia  TEXT,
            cant2            REAL,
            repetidos        INTEGER DEFAULT 0,
            vacio            INTEGER DEFAULT 0,
            sort_order       INTEGER DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS decks (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            cant        INTEGER DEFAULT 1,
            nombre      TEXT    NOT NULL,
            copia_extra REAL,
            repetidos   INTEGER DEFAULT 0,
            deck_nombre TEXT    DEFAULT 'Deck Principal'
        );
        CREATE TABLE IF NOT EXISTS pendientes (
            id     INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            tipo   TEXT,
            notas  TEXT
        );
        CREATE TABLE IF NOT EXISTS config (
            clave TEXT PRIMARY KEY,
            valor TEXT
        );
        CREATE TABLE IF NOT EXISTS albums (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre  TEXT    NOT NULL UNIQUE,
            paginas INTEGER NOT NULL DEFAULT 1
        );
    """)
    # Migraciones
    for migration in [
        "ALTER TABLE decks ADD COLUMN deck_nombre TEXT DEFAULT 'Deck Principal'",
        "ALTER TABLE cartas ADD COLUMN sort_order INTEGER DEFAULT 0",
    ]:
        try:
            conn.execute(migration)
            conn.commit()
        except Exception:
            pass
    conn.close()


def get_cartas(tipo=None, search=None, mostrar_vacios=True, include_pendientes=True):
    conn = get_conn()
    conditions, params = [], []

    if tipo:
        conditions.append("tipo = ?")
        params.append(tipo)
    if not mostrar_vacios:
        conditions.append("vacio = 0")
    if search:
        conditions.append("nombre LIKE ?")
        params.append(f"%{search}%")

    where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    rows = list(conn.execute(
        f"SELECT *, 0 as pendiente FROM cartas {where} ORDER BY sort_order, id", params
    ).fetchall())

    if include_pendientes and search:
        pend = conn.execute(
            "SELECT id, nombre, tipo, notas, 0 as cant, NULL as ubicacion, "
            "NULL as ubicacion_copia, NULL as cant2, 0 as repetidos, 0 as vacio, "
            "1 as pendiente FROM pendientes WHERE nombre LIKE ?",
            (f"%{search}%",)
        ).fetchall()
        rows += list(pend)

    conn.close()
    return rows


def get_stats():
    conn = get_conn()
    cfg = conn.execute("SELECT valor FROM config WHERE clave='total_cartas'").fetchone()
    total = int(cfg["valor"]) if cfg else 0

    owned = conn.execute(
        "SELECT COALESCE(SUM(cant), 0) FROM cartas WHERE vacio=0"
    ).fetchone()[0]

    por_tipo = conn.execute(
        "SELECT tipo, COUNT(*) as distintas, SUM(cant) as total, "
        "SUM(CASE WHEN vacio=1 THEN 1 ELSE 0 END) as vacios "
        "FROM cartas GROUP BY tipo ORDER BY tipo"
    ).fetchall()

    conn.close()
    return {
        "total": total,
        "owned": owned,
        "restantes": max(0, total - owned),
        "por_tipo": [(r["tipo"], r["distintas"], r["total"], r["vacios"]) for r in por_tipo],
    }


def add_carta(tipo, cant, nombre, ubicacion, ubi_copia, cant2, repetidos):
    conn = get_conn()
    conn.execute(
        "INSERT INTO cartas (tipo,cant,nombre,ubicacion,ubicacion_copia,cant2,repetidos,vacio) "
        "VALUES (?,?,?,?,?,?,?,0)",
        (tipo, cant, nombre, ubicacion, ubi_copia, cant2, repetidos),
    )
    conn.commit()
    conn.close()


def update_carta(carta_id, cant, nombre, ubicacion, ubi_copia, cant2, repetidos):
    conn = get_conn()
    conn.execute(
        "UPDATE cartas SET cant=?,nombre=?,ubicacion=?,ubicacion_copia=?,"
        "cant2=?,repetidos=?,vacio=0 WHERE id=?",
        (cant, nombre, ubicacion, ubi_copia, cant2, repetidos, carta_id),
    )
    conn.commit()
    conn.close()


def delete_carta(carta_id):
    conn = get_conn()
    conn.execute("DELETE FROM cartas WHERE id=?", (carta_id,))
    conn.commit()
    conn.close()


def clear_carta(carta_id):
    """Convierte una carta en hueco vacío (no la elimina)."""
    conn = get_conn()
    conn.execute(
        "UPDATE cartas SET cant=0,nombre=NULL,ubicacion=NULL,ubicacion_copia=NULL,"
        "cant2=NULL,repetidos=0,vacio=1 WHERE id=?",
        (carta_id,),
    )
    conn.commit()
    conn.close()


def check_repetido(nombre):
    conn = get_conn()
    r = conn.execute(
        "SELECT COUNT(*) FROM cartas WHERE LOWER(nombre)=LOWER(?) AND vacio=0",
        (nombre,),
    ).fetchone()[0]
    conn.close()
    return r > 0


def get_deck_names():
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT deck_nombre FROM decks ORDER BY deck_nombre"
    ).fetchall()
    conn.close()
    return [r[0] for r in rows if r[0]]


def get_decks(search=None, deck_nombre=None):
    conn = get_conn()
    conditions, params = [], []
    if search:
        conditions.append("nombre LIKE ?")
        params.append(f"%{search}%")
    if deck_nombre and deck_nombre != "Todos":
        conditions.append("deck_nombre = ?")
        params.append(deck_nombre)
    where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    rows = conn.execute(
        f"SELECT * FROM decks {where} ORDER BY deck_nombre, nombre", params
    ).fetchall()
    conn.close()
    return rows


def move_carta_to_deck(carta_id, deck_nombre, dejar_hueco=True):
    """Mueve una carta del álbum a un deck. Opcionalmente deja el hueco."""
    conn = get_conn()
    carta = conn.execute("SELECT * FROM cartas WHERE id=?", (carta_id,)).fetchone()
    if not carta:
        conn.close()
        return
    carta = dict(carta)

    # ¿Hay otra copia en el álbum?
    copies_in_album = conn.execute(
        "SELECT COUNT(*) FROM cartas WHERE LOWER(nombre)=LOWER(?) AND vacio=0 AND id!=?",
        (carta["nombre"], carta_id),
    ).fetchone()[0]

    conn.execute(
        "INSERT INTO decks (cant, nombre, copia_extra, repetidos, deck_nombre) "
        "VALUES (?,?,?,?,?)",
        (carta["cant"], carta["nombre"], None,
         1 if copies_in_album > 0 else 0,
         deck_nombre),
    )

    if dejar_hueco:
        conn.execute(
            "UPDATE cartas SET cant=0,nombre=NULL,ubicacion=NULL,ubicacion_copia=NULL,"
            "cant2=NULL,repetidos=0,vacio=1 WHERE id=?",
            (carta_id,),
        )
    else:
        conn.execute("DELETE FROM cartas WHERE id=?", (carta_id,))

    conn.commit()
    conn.close()


def delete_deck_carta(deck_id):
    conn = get_conn()
    conn.execute("DELETE FROM decks WHERE id=?", (deck_id,))
    conn.commit()
    conn.close()


def delete_deck(deck_nombre):
    conn = get_conn()
    conn.execute("DELETE FROM decks WHERE deck_nombre=?", (deck_nombre,))
    conn.commit()
    conn.close()


def rename_deck(old_nombre, new_nombre):
    conn = get_conn()
    conn.execute("UPDATE decks SET deck_nombre=? WHERE deck_nombre=?", (new_nombre, old_nombre))
    conn.commit()
    conn.close()


# ── Álbumes ──────────────────────────────────────────────────────────────────

def get_albums():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM albums ORDER BY nombre").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_album(nombre, paginas):
    conn = get_conn()
    conn.execute("INSERT OR IGNORE INTO albums (nombre, paginas) VALUES (?,?)",
                 (nombre, int(paginas)))
    conn.commit()
    conn.close()


def update_album(album_id, nombre, paginas):
    conn = get_conn()
    conn.execute("UPDATE albums SET nombre=?, paginas=? WHERE id=?",
                 (nombre, int(paginas), album_id))
    conn.commit()
    conn.close()


def delete_album(album_id):
    conn = get_conn()
    conn.execute("DELETE FROM albums WHERE id=?", (album_id,))
    conn.commit()
    conn.close()


def get_ubicaciones():
    """Genera lista de ubicaciones desde álbumes + las existentes en cartas."""
    conn = get_conn()
    albums = conn.execute("SELECT nombre, paginas FROM albums ORDER BY nombre").fetchall()
    ubicaciones = []
    for a in albums:
        for p in range(1, int(a["paginas"]) + 1):
            ubicaciones.append(f"{a['nombre']}.pag.{p}")

    # Añadir ubicaciones existentes no generadas por álbumes
    existing = conn.execute(
        "SELECT DISTINCT ubicacion FROM cartas WHERE ubicacion IS NOT NULL ORDER BY ubicacion"
    ).fetchall()
    gen_set = set(ubicaciones)
    for r in existing:
        if r[0] and r[0] not in gen_set:
            ubicaciones.append(r[0])

    conn.close()
    return ubicaciones


def add_pendiente(nombre, tipo, notas):
    conn = get_conn()
    conn.execute(
        "INSERT INTO pendientes (nombre,tipo,notas) VALUES (?,?,?)",
        (nombre, tipo, notas),
    )
    conn.commit()
    conn.close()


def get_pendientes():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM pendientes ORDER BY nombre").fetchall()
    conn.close()
    return rows


def delete_pendiente(pend_id):
    conn = get_conn()
    conn.execute("DELETE FROM pendientes WHERE id=?", (pend_id,))
    conn.commit()
    conn.close()
