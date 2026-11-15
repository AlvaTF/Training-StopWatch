import asyncio
import io
import json
import os
import tempfile
import time
from datetime import date

import flet as ft
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

KEY = "cronoclase_v1"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

DATA = {"seq": 0, "turnos": [], "students": [], "exercises": [], "results": []}


# ---------------------------------------------------------------- datos
def next_id():
    DATA["seq"] += 1
    return DATA["seq"]


def students(turno=None):
    return sorted(
        (s for s in DATA["students"] if turno is None or s["turno"] == turno),
        key=lambda s: (s["turno"], s["id"]),
    )


def exercises(day):
    return [e for e in DATA["exercises"] if e["day"] == day]


def times_index():
    """(ejercicio, alumno) -> {pasada: segundos}"""
    idx = {}
    for r in DATA["results"]:
        idx.setdefault((r["ex"], r["st"]), {})[r["p"]] = r["sec"]
    return idx


def fmt(sec):
    m, s = divmod(sec, 60)
    return f"{int(m):02d}:{s:05.2f}"


# ---------------------------------------------------------------- exportacion
def write_day_blocks(ws, day):
    """Un bloque por ejercicio; todos los turnos comparten la planilla y la
    columna 'Turno' indica a cual pertenece cada alumno."""
    idx = times_index()
    row = 1
    for ex in exercises(day):
        passes = ex["passes"]
        ws.cell(row, 1, f"{ex['name']}  ({day})").font = Font(bold=True, size=12)
        row += 1
        headers = (
            ["Turno", "Alumno"]
            + [f"Pasada {i}" for i in range(1, passes + 1)]
            + ["Mejor", "Promedio"]
        )
        for c, h in enumerate(headers, 1):
            ws.cell(row, c, h).font = Font(bold=True)
        row += 1
        for s in students():
            t = idx.get((ex["id"], s["id"]))
            if not t:
                continue
            ws.cell(row, 1, s["turno"])
            ws.cell(row, 2, s["name"])
            for p in range(1, passes + 1):
                if p in t:
                    ws.cell(row, 2 + p, round(t[p], 2))
            ws.cell(row, passes + 3, round(min(t.values()), 2))
            ws.cell(row, passes + 4, round(sum(t.values()) / len(t), 2))
            row += 1
        row += 1
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 28


def build_day_wb(day):
    wb = Workbook()
    ws = wb.active
    ws.title = day
    write_day_blocks(ws, day)
    return wb


def build_all_wb():
    """Una hoja: alumnos en filas (col. A); por cada dia un titulo arriba y
    los ejercicios como columnas angostas (una por pasada)."""
    yellow = PatternFill("solid", fgColor="FFFF00")
    palette = ["00B050", "FFC000", "FF0000", "5B9BD5", "B084CC", "F4B183"]
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "Todos los dias"
    studs = students()
    idx = times_index()

    for col, text in ((1, "Nombre"), (2, "Turno")):
        c = ws.cell(2, col, text)
        c.fill, c.alignment, c.border = yellow, center, box
        c.font = Font(bold=True)
        ws.cell(1, col).fill = yellow
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 8

    col = 3
    for day in sorted({e["day"] for e in DATA["exercises"]}):
        day_start = col
        for i, ex in enumerate(exercises(day)):
            passes = ex["passes"]
            color = PatternFill("solid", fgColor=palette[i % len(palette)])
            first, last = col, col + passes - 1
            if last > first:
                ws.merge_cells(start_row=2, start_column=first, end_row=2, end_column=last)
            h = ws.cell(2, first, ex["name"])
            h.alignment, h.font = center, Font(bold=True)
            for cc in range(first, last + 1):
                ws.cell(2, cc).fill = color
                ws.cell(2, cc).border = box
                ws.column_dimensions[get_column_letter(cc)].width = 7
            for r, s in enumerate(studs, 3):
                t = idx.get((ex["id"], s["id"]), {})
                for p in range(1, passes + 1):
                    cell = ws.cell(r, first + p - 1)
                    cell.border = box
                    if p in t:
                        cell.value = round(t[p], 2)
            col = last + 1
        if col > day_start:
            if col - 1 > day_start:
                ws.merge_cells(start_row=1, start_column=day_start, end_row=1, end_column=col - 1)
            d = ws.cell(1, day_start, day)
            d.alignment, d.font = center, Font(bold=True, size=12)
            for cc in range(day_start, col):
                ws.cell(1, cc).fill = yellow
                ws.cell(1, cc).border = box

    for r, s in enumerate(studs, 3):
        ws.cell(r, 1, s["name"]).border = box
        c = ws.cell(r, 2, s["turno"])
        c.border, c.alignment = box, Alignment(horizontal="center")

    ws.freeze_panes = "C3"
    return wb


def wb_bytes(wb):
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def wb_tsv(wb):
    """Texto separado por tabs para pegar en Excel / Google Sheets.
    Los decimales llevan coma (configuracion regional de Argentina)."""

    def cell(v):
        if v is None:
            return ""
        if isinstance(v, float):
            return str(v).replace(".", ",")
        return str(v)

    ws = wb.active
    return "\n".join("\t".join(cell(v) for v in row) for row in ws.iter_rows(values_only=True))


def deliver(page, name, data):
    """Descarga el archivo en el navegador (Pyodide). Fuera del navegador
    (prueba en la compu) lo guarda en la carpeta temporal."""
    try:
        import js
        from pyodide.ffi import to_js
    except ImportError:
        path = os.path.join(tempfile.gettempdir(), name)
        with open(path, "wb") as f:
            f.write(data)
        return f"Guardado en: {path}"

    arr = js.Uint8Array.new(len(data))
    arr.assign(data)
    blob = js.Blob.new(to_js([arr]), to_js({"type": XLSX_MIME}, dict_converter=js.Object.fromEntries))
    url = js.URL.createObjectURL(blob)
    doc = getattr(js, "document", None)
    if doc is not None:  # Python corre en el hilo principal
        a = doc.createElement("a")
        a.href = url
        a.download = name
        doc.body.appendChild(a)
        a.click()
        a.remove()
        return f"Descarga iniciada: {name}"
    page.launch_url(url)  # Python corre en un web worker
    return "Se abrio el archivo en otra pestana. Si no se descarga, usa 'Copiar para pegar'."


# ---------------------------------------------------------------- interfaz
async def main(page: ft.Page):
    page.title = "CronoClase"
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 16

    # cargar lo guardado en el navegador
    raw = await page.client_storage.get_async(KEY)
    if raw:
        try:
            DATA.update(json.loads(raw))
        except ValueError:
            pass

    async def save_data():
        await page.client_storage.set_async(KEY, json.dumps(DATA))

    def commit():
        page.run_task(save_data)

    settings = {
        "round_robin": False,
        "turno_alta": None,
        "turno_filtro": None,
        "elegido": {},  # ejercicio -> alumno elegido a mano
    }

    def today():
        return date.today().isoformat()

    def render(*controls):
        page.clean()
        page.add(*controls)
        page.update()

    def back_btn(target):
        return ft.TextButton("< Volver", on_click=lambda e: target())

    # ---------------- inicio
    def show_home():
        render(
            ft.Text("CronoClase", size=30, weight=ft.FontWeight.BOLD),
            ft.Text(f"Hoy: {today()}"),
            ft.ElevatedButton("Turnos y alumnos", on_click=lambda e: show_students(), width=300),
            ft.ElevatedButton("Entrenamiento de hoy", on_click=lambda e: show_today(), width=300),
            ft.ElevatedButton("Exportar", on_click=lambda e: show_export(), width=300),
        )

    # ---------------- turnos y alumnos
    def show_students():
        turnos = sorted(DATA["turnos"])
        alta = settings["turno_alta"]
        if alta not in turnos:
            alta = turnos[0] if turnos else None
            settings["turno_alta"] = alta

        turno_in = ft.TextField(
            label="Nuevo turno (numero)", width=200, keyboard_type=ft.KeyboardType.NUMBER
        )

        def add_turno(e):
            v = turno_in.value.strip()
            if v.lstrip("-").isdigit():
                n = int(v)
                if n not in DATA["turnos"]:
                    DATA["turnos"].append(n)
                settings["turno_alta"] = n
                commit()
                show_students()

        def delete_turno(t):
            DATA["turnos"].remove(t)
            commit()
            show_students()

        turno_dd = ft.Dropdown(
            label="Turno del alumno",
            width=200,
            options=[ft.dropdown.Option(str(t), f"Turno {t}") for t in turnos],
            value=str(alta) if alta is not None else None,
            on_change=lambda e: settings.update(turno_alta=int(e.control.value)),
        )
        name_in = ft.TextField(label="Nombre del alumno", expand=True)

        def add_student(e):
            if name_in.value.strip() and turno_dd.value is not None:
                DATA["students"].append(
                    {"id": next_id(), "name": name_in.value.strip(), "turno": int(turno_dd.value)}
                )
                commit()
                show_students()

        def delete_student(sid):
            DATA["students"][:] = [s for s in DATA["students"] if s["id"] != sid]
            DATA["results"][:] = [r for r in DATA["results"] if r["st"] != sid]
            commit()
            show_students()

        listing = []
        for t in turnos:
            studs = students(t)
            header = [ft.Text(f"Turno {t}", size=18, weight=ft.FontWeight.BOLD)]
            if not studs:
                header.append(ft.TextButton("Borrar turno", on_click=lambda e, t=t: delete_turno(t)))
            listing.append(ft.Row(header))
            for s in studs:
                listing.append(
                    ft.Row(
                        [
                            ft.Text(s["name"], expand=True),
                            ft.TextButton("Borrar", on_click=lambda e, i=s["id"]: delete_student(i)),
                        ]
                    )
                )

        render(
            back_btn(show_home),
            ft.Text("Turnos y alumnos", size=24, weight=ft.FontWeight.BOLD),
            ft.Row([turno_in, ft.ElevatedButton("Agregar turno", on_click=add_turno)]),
            ft.Divider(),
            turno_dd,
            ft.Row([name_in, ft.ElevatedButton("Agregar", on_click=add_student)]),
            ft.Divider(),
            *listing,
        )

    # ---------------- ejercicios del dia
    def show_today():
        name_in = ft.TextField(label="Ejercicio", expand=True)
        passes_in = ft.TextField(
            label="Pasadas", value="3", width=100, keyboard_type=ft.KeyboardType.NUMBER
        )

        def add(e):
            if name_in.value.strip() and passes_in.value.strip().isdigit():
                DATA["exercises"].append(
                    {
                        "id": next_id(),
                        "day": today(),
                        "name": name_in.value.strip(),
                        "passes": int(passes_in.value),
                    }
                )
                commit()
                show_today()

        def delete(eid):
            DATA["exercises"][:] = [x for x in DATA["exercises"] if x["id"] != eid]
            DATA["results"][:] = [r for r in DATA["results"] if r["ex"] != eid]
            commit()
            show_today()

        rows = [
            ft.Row(
                [
                    ft.ElevatedButton(
                        f"{x['name']} ({x['passes']} pasadas)",
                        expand=True,
                        on_click=lambda e, i=x["id"]: show_run(i),
                    ),
                    ft.TextButton("Borrar", on_click=lambda e, i=x["id"]: delete(i)),
                ]
            )
            for x in exercises(today())
        ]
        render(
            back_btn(show_home),
            ft.Text("Ejercicios de hoy", size=24, weight=ft.FontWeight.BOLD),
            ft.Row([name_in, passes_in]),
            ft.ElevatedButton("Agregar ejercicio", on_click=add),
            ft.Divider(),
            ft.Text("Toca un ejercicio para empezar a cronometrar:"),
            *rows,
        )

    # ---------------- cronometro por ejercicio
    def show_run(ex_id):
        ex = next(x for x in DATA["exercises"] if x["id"] == ex_id)
        passes = ex["passes"]
        turnos = sorted(DATA["turnos"])
        tf = settings["turno_filtro"]
        if tf not in turnos:
            tf = None
            settings["turno_filtro"] = None

        studs = students(tf)
        idx = times_index()
        counts = {s["id"]: len(idx.get((ex_id, s["id"]), {})) for s in studs}
        by_id = {s["id"]: s for s in studs}

        # quien sigue: el alumno elegido a mano o, si no, el orden automatico
        current = None  # (alumno, numero de pasada)
        chosen = settings["elegido"].get(ex_id)
        if chosen in by_id and counts[chosen] < passes:
            current = (by_id[chosen], counts[chosen] + 1)
        elif studs:
            if settings["round_robin"]:
                lowest = min(counts.values())
                if lowest < passes:
                    s = next(s for s in studs if counts[s["id"]] == lowest)
                    current = (s, lowest + 1)
            else:
                for s in studs:
                    if counts[s["id"]] < passes:
                        current = (s, counts[s["id"]] + 1)
                        break

        state = {"running": False, "t0": 0.0}
        label = ft.Text("00:00.00", size=64, weight=ft.FontWeight.BOLD)
        btn = ft.ElevatedButton("INICIAR", width=300, height=70)

        async def tick():
            while state["running"]:
                label.value = fmt(time.perf_counter() - state["t0"])
                label.update()
                await asyncio.sleep(0.03)

        def toggle(e):
            if not state["running"]:
                state["running"] = True
                state["t0"] = time.perf_counter()
                btn.text = "PARAR"
                btn.update()
                page.run_task(tick)
            else:
                elapsed = time.perf_counter() - state["t0"]
                state["running"] = False
                DATA["results"].append(
                    {
                        "id": next_id(),
                        "ex": ex_id,
                        "st": current[0]["id"],
                        "p": current[1],
                        "sec": elapsed,
                    }
                )
                commit()
                show_run(ex_id)

        btn.on_click = toggle

        def undo(e):
            for i in range(len(DATA["results"]) - 1, -1, -1):
                if DATA["results"][i]["ex"] == ex_id:
                    del DATA["results"][i]
                    break
            commit()
            show_run(ex_id)

        def set_mode(e):
            settings["round_robin"] = e.control.value
            show_run(ex_id)

        def set_turno(e):
            v = e.control.value
            settings["turno_filtro"] = None if v == "todos" else int(v)
            settings["elegido"].pop(ex_id, None)
            show_run(ex_id)

        def set_student(e):
            v = e.control.value
            if v == "auto":
                settings["elegido"].pop(ex_id, None)
            else:
                settings["elegido"][ex_id] = int(v)
            show_run(ex_id)

        def leave(e):
            state["running"] = False
            show_today()

        turno_dd = ft.Dropdown(
            label="Turno",
            width=160,
            options=[ft.dropdown.Option("todos", "Todos")]
            + [ft.dropdown.Option(str(t), f"Turno {t}") for t in turnos],
            value="todos" if tf is None else str(tf),
            on_change=set_turno,
        )
        student_dd = ft.Dropdown(
            label="Alumno",
            expand=True,
            options=[ft.dropdown.Option("auto", "Automatico (siguiente)")]
            + [
                ft.dropdown.Option(
                    str(s["id"]), f"{s['name']} (T{s['turno']}) {counts[s['id']]}/{passes}"
                )
                for s in studs
            ],
            value=str(chosen) if chosen in by_id else "auto",
            on_change=set_student,
        )

        if current:
            head = [
                ft.Text(
                    f"{current[0]['name']}  (Turno {current[0]['turno']})",
                    size=26,
                    weight=ft.FontWeight.BOLD,
                ),
                ft.Text(f"Pasada {current[1]} de {passes}", size=18),
                label,
                btn,
            ]
        elif not studs:
            head = [ft.Text("No hay alumnos en esta seleccion.")]
        else:
            head = [ft.Text("Ejercicio completo", size=24, weight=ft.FontWeight.BOLD)]

        summary = []
        for s in studs:
            t = idx.get((ex_id, s["id"]), {})
            summary.append(
                ft.Text(f"[T{s['turno']}] {s['name']}: " + "  |  ".join(fmt(t[p]) for p in sorted(t)))
            )

        render(
            ft.TextButton("< Volver", on_click=leave),
            ft.Text(ex["name"], size=22, weight=ft.FontWeight.BOLD),
            ft.Row([turno_dd, student_dd]),
            ft.Switch(
                label="Rotar alumnos en cada pasada",
                value=settings["round_robin"],
                on_change=set_mode,
            ),
            *head,
            ft.TextButton("Deshacer ultimo tiempo", on_click=undo),
            ft.Divider(),
            *summary,
        )

    # ---------------- exportar
    def show_export():
        out = ft.Text("")

        def say(msg):
            out.value = msg
            out.update()

        def dl_day(e):
            say(deliver(page, f"clase_{today()}.xlsx", wb_bytes(build_day_wb(today()))))

        def dl_all(e):
            say(deliver(page, "clase_todos_los_dias.xlsx", wb_bytes(build_all_wb())))

        def copy_day(e):
            page.set_clipboard(wb_tsv(build_day_wb(today())))
            say("Copiado. Pegalo en Excel o Google Sheets (celda A1).")

        def copy_all(e):
            page.set_clipboard(wb_tsv(build_all_wb()))
            say("Copiado. Pegalo en Excel o Google Sheets (celda A1).")

        render(
            back_btn(show_home),
            ft.Text("Exportar", size=24, weight=ft.FontWeight.BOLD),
            ft.Text("Descargar Excel (.xlsx)", weight=ft.FontWeight.BOLD),
            ft.ElevatedButton("El dia de hoy", on_click=dl_day, width=300),
            ft.ElevatedButton("Todos los dias", on_click=dl_all, width=300),
            ft.Divider(),
            ft.Text("Copiar para pegar en Excel / Sheets", weight=ft.FontWeight.BOLD),
            ft.ElevatedButton("Copiar el dia de hoy", on_click=copy_day, width=300),
            ft.ElevatedButton("Copiar todos los dias", on_click=copy_all, width=300),
            out,
            ft.Divider(),
            ft.Text(
                "Los datos se guardan en este navegador. Si borras los datos del sitio "
                "se pierden: exporta seguido.",
                size=12,
            ),
        )

    show_home()


ft.app(main)