import openpyxl
from db import get_conn, init_db

TIPOS = ["Monstruos", "Rituales", "Fusion", "Sincronia", "Enlace", "XYZ", "Pendulo", "Magia", "Trampas"]
SINGLE_BLOCK = {"Rituales"}
MAX_CONSECUTIVE_BLANKS = 30


def _clean(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _clean_int(v):
    """Devuelve entero o None; descarta texto no numérico."""
    if v is None:
        return None
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return None


def _to_bool(v):
    if v is None:
        return 0
    return 1 if str(v).strip().lower() in ("true", "1", "verdadero") else 0


def _safe(row, idx, default=None):
    try:
        v = row[idx]
        return v if v is not None else default
    except (IndexError, TypeError):
        return default


def _row_has_data(row):
    return any(
        v is not None and str(v).strip() not in ("", "None")
        for v in (row or [])
    )


def import_excel(path, progress_cb=None):
    init_db()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    conn = get_conn()

    conn.execute("DELETE FROM cartas")
    conn.execute("DELETE FROM decks")

    total_steps = len(TIPOS) + 1  # +1 for decks

    sort_counter = 0

    for step, tipo in enumerate(TIPOS):
        if tipo not in wb.sheetnames:
            if progress_cb:
                progress_cb(step + 1, total_steps)
            continue

        ws = wb[tipo]
        is_single = tipo in SINGLE_BLOCK
        consecutive_blanks = 0
        seen_data = False

        for row in ws.iter_rows(min_row=2, values_only=True):
            # Left block values
            nombre_l = _clean(_safe(row, 1))
            # Right block values (only for dual-block sheets)
            nombre_r = None if is_single else _clean(_safe(row, 8))

            # Saltar filas plantilla (cabecera repetida)
            if nombre_l == "Nombre":
                nombre_l = None
            if nombre_r == "Nombre":
                nombre_r = None

            left_has = nombre_l is not None
            right_has = nombre_r is not None
            row_has = left_has or right_has or _row_has_data(row[:6]) or (not is_single and _row_has_data(_safe(row, slice(7, 13), [])))

            if not row_has:
                consecutive_blanks += 1
                if consecutive_blanks >= MAX_CONSECUTIVE_BLANKS:
                    break
                # Preserve vacant slots within a data range
                if seen_data:
                    conn.execute(
                        "INSERT INTO cartas (tipo,cant,nombre,vacio,sort_order) VALUES (?,0,NULL,1,?)",
                        (tipo, sort_counter),
                    )
                    sort_counter += 1
                    if not is_single:
                        conn.execute(
                            "INSERT INTO cartas (tipo,cant,nombre,vacio,sort_order) VALUES (?,0,NULL,1,?)",
                            (tipo, sort_counter),
                        )
                        sort_counter += 1
                continue

            consecutive_blanks = 0
            seen_data = True

            # Left block
            conn.execute(
                "INSERT INTO cartas (tipo,cant,nombre,ubicacion,ubicacion_copia,cant2,repetidos,vacio,sort_order) "
                "VALUES (?,?,?,?,?,?,?,?,?)",
                (
                    tipo,
                    _safe(row, 0, 0) or 0,
                    nombre_l,
                    _clean(_safe(row, 2)),
                    _clean(_safe(row, 3)),
                    _safe(row, 4),
                    _to_bool(_safe(row, 5)),
                    0 if left_has else 1,
                    sort_counter,
                ),
            )
            sort_counter += 1

            # Right block
            if not is_single:
                conn.execute(
                    "INSERT INTO cartas (tipo,cant,nombre,ubicacion,ubicacion_copia,cant2,repetidos,vacio,sort_order) "
                    "VALUES (?,?,?,?,?,?,?,?,?)",
                    (
                        tipo,
                        _safe(row, 7, 0) or 0,
                        nombre_r,
                        _clean(_safe(row, 9)),
                        _clean(_safe(row, 10)),
                        _safe(row, 11),
                        _to_bool(_safe(row, 12)),
                        0 if right_has else 1,
                        sort_counter,
                    ),
                )
                sort_counter += 1

        if progress_cb:
            progress_cb(step + 1, total_steps)

    # Decks sheet — cada cabecera (Nombre/Cant) marca inicio de nueva sección de decks
    if "Decks" in wb.sheetnames:
        ws = wb["Decks"]
        deck_section = 0  # se incrementa con cada fila cabecera
        for row in ws.iter_rows(min_row=1, values_only=True):
            if not row:
                continue
            # Detectar fila cabecera: segunda celda es "Nombre"
            if _clean(_safe(row, 1)) == "Nombre":
                deck_section += 1
                continue
            for block_idx, start in enumerate((0, 4, 8)):
                nombre = _clean(_safe(row, start + 1))
                if nombre:
                    deck_nombre = f"Deck {(deck_section - 1) * 3 + block_idx + 1}"
                    conn.execute(
                        "INSERT INTO decks (cant,nombre,copia_extra,repetidos,deck_nombre) "
                        "VALUES (?,?,?,?,?)",
                        (
                            _safe(row, start, 1) or 1,
                            nombre,
                            _clean_int(_safe(row, start + 2)),
                            _to_bool(_safe(row, start + 3)),
                            deck_nombre,
                        ),
                    )

    # Estadistica: total from O26 (col 15, row 26)
    if "Estadistica" in wb.sheetnames:
        ws_stat = wb["Estadistica"]
        total_val = ws_stat["O26"].value
        if total_val:
            conn.execute(
                "INSERT OR REPLACE INTO config VALUES ('total_cartas',?)",
                (str(int(total_val)),),
            )

    if progress_cb:
        progress_cb(total_steps, total_steps)

    conn.commit()
    conn.close()
    wb.close()
