import asyncio
import os
import sqlite3
import time
from datetime import date

import flet as ft
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ---------------------------------------------------------------- almacenamiento
DATA_DIR = os.getenv("FLET_APP_STORAGE_DATA", os.path.abspath("."))
EXPORT_DIR = os.path.join(DATA_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)

db = sqlite3.connect(os.path.join(DATA_DIR, "cronoclase.db"), check_same_thread=False)
db.executescript(
    """
    CREATE TABLE IF NOT EXISTS turnos(
        id INTEGER PRIMARY KEY);
    CREATE TABLE IF NOT EXISTS students(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS exercises(
        id INTEGER PRIMARY KEY AUTOINCREMENT, day TEXT NOT NULL,
        name TEXT NOT NULL, passes INTEGER NOT NULL);
    CREATE TABLE IF NOT EXISTS results(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        exercise_id INTEGER NOT NULL, student_id INTEGER NOT NULL,
        pass_no INTEGER NOT NULL, seconds REAL NOT NULL);
    """
)
# migracion para bases creadas con la version anterior (sin turno)
try:
    db.execute("ALTER TABLE students ADD COLUMN turno INTEGER NOT NULL DEFAULT 0")
except sqlite3.OperationalError:
    pass
db.execute("INSERT OR IGNORE INTO turnos(id) SELECT DISTINCT turno FROM students")
db.commit()


def q(sql, args=()):
    return db.execute(sql, args).fetchall()


def run(sql, args=()):
    db.execute(sql, args)
    db.commit()


def fmt(sec):
    m, s = divmod(sec, 60)
    return f"{int(m):02d}:{s:05.2f}"


# ---------------------------------------------------------------- exportacion
def write_day_blocks(ws, day):
    """Un bloque por ejercicio. Todos los turnos comparten la planilla;
    la columna 'Turno' indica a cual pertenece cada alumno."""
    row = 1
    for ex_id, name, passes in q(
        "SELECT id, name, passes FROM exercises WHERE day=? ORDER BY id", (day,)
    ):
        ws.cell(row, 1, f"{name}  ({day})").font = Font(bold=True, size=12)
        row += 1
        headers = (
            ["Turno", "Alumno"]
            + [f"Pasada {i}" for i in range(1, passes + 1)]
            + ["Mejor", "Promedio"]
        )
        for c, h in enumerate(headers, 1):
            ws.cell(row, c, h).font = Font(bold=True)
        row += 1
        for sid, sname, turno in q(
            "SELECT DISTINCT s.id, s.name, s.turno FROM students s "
            "JOIN results r ON r.student_id=s.id WHERE r.exercise_id=? "
            "ORDER BY s.turno, s.id",
            (ex_id,),
        ):
            times = {
                p: sec
                for p, sec in q(
                    "SELECT pass_no, seconds FROM results WHERE exercise_id=? AND student_id=?",
                    (ex_id, sid),
                )
            }
            ws.cell(row, 1, turno)
            ws.cell(row, 2, sname)
            for p in range(1, passes + 1):
                if p in times:
                    ws.cell(row, 2 + p, round(times[p], 2))
            if times:
                ws.cell(row, passes + 3, round(min(times.values()), 2))
                ws.cell(row, passes + 4, round(sum(times.values()) / len(times), 2))
            row += 1
        row += 1
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 28


def export_day(day):
    wb = Workbook()
    ws = wb.active
    ws.title = day
    write_day_blocks(ws, day)
    path = os.path.join(EXPORT_DIR, f"clase_{day}.xlsx")
    wb.save(path)
    return path


def export_all():
    """Una sola hoja: alumnos en filas (col. A) y, por cada dia, un titulo
    arriba con los ejercicios como columnas angostas (una por pasada)."""
    from openpyxl.utils import get_column_letter

    yellow = PatternFill("solid", fgColor="FFFF00")
    palette = ["00B050", "FFC000", "FF0000", "5B9BD5", "B084CC", "F4B183"]
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "Todos los dias"

    students = q("SELECT id, name, turno FROM students ORDER BY turno, id")

    # columnas fijas
    for col, text in ((1, "Nombre"), (2, "Turno")):
        c = ws.cell(2, col, text)
        c.fill, c.alignment, c.border = yellow, center, box
        c.font = Font(bold=True)
        ws.cell(1, col).fill = yellow
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 8

    # tiempos: (ejercicio, alumno) -> {pasada: segundos}
    times = {}
    for ex_id, sid, p, sec in q("SELECT exercise_id, student_id, pass_no, seconds FROM results"):
        times.setdefault((ex_id, sid), {})[p] = sec

    col = 3
    for (day,) in q("SELECT DISTINCT day FROM exercises ORDER BY day"):
        day_start = col
        for i, (ex_id, ex_name, passes) in enumerate(
            q("SELECT id, name, passes FROM exercises WHERE day=? ORDER BY id", (day,))
        ):
            color = PatternFill("solid", fgColor=palette[i % len(palette)])
            first, last = col, col + passes - 1
            # titulo del ejercicio (fila 2), combinado sobre sus pasadas
            if last > first:
                ws.merge_cells(start_row=2, start_column=first, end_row=2, end_column=last)
            h = ws.cell(2, first, ex_name)
            h.fill, h.alignment, h.font = color, center, Font(bold=True)
            for cc in range(first, last + 1):
                ws.cell(2, cc).fill = color
                ws.cell(2, cc).border = box
                ws.column_dimensions[get_column_letter(cc)].width = 7
            # tiempos de cada alumno
            for r, (sid, _, _) in enumerate(students, 3):
                t = times.get((ex_id, sid), {})
                for p in range(1, passes + 1):
                    cell = ws.cell(r, first + p - 1)
                    cell.border = box
                    if p in t:
                        cell.value = round(t[p], 2)
            col = last + 1
        # titulo del dia (fila 1), combinado sobre todos sus ejercicios
        if col > day_start:
            if col - 1 > day_start:
                ws.merge_cells(start_row=1, start_column=day_start, end_row=1, end_column=col - 1)
            d = ws.cell(1, day_start, day)
            d.alignment, d.font = center, Font(bold=True, size=12)
            for cc in range(day_start, col):
                ws.cell(1, cc).fill = yellow
                ws.cell(1, cc).border = box

    # nombre y turno de cada alumno
    for r, (_, sname, turno) in enumerate(students, 3):
        ws.cell(r, 1, sname).border = box
        c = ws.cell(r, 2, turno)
        c.border, c.alignment = box, Alignment(horizontal="center")

    ws.freeze_panes = "C3"
    path = os.path.join(EXPORT_DIR, "clase_todos_los_dias.xlsx")
    wb.save(path)
    return path


# ---------------------------------------------------------------- interfaz
def main(page: ft.Page):
    page.title = "CronoClase"
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 16
    settings = {
        "round_robin": False,
        "turno_alta": None,  # turno elegido al dar de alta alumnos
        "turno_filtro": None,  # turno elegido en la pantalla de cronometraje
        "elegido": {},  # ejercicio -> alumno elegido manualmente
    }

    def today():
        return date.today().isoformat()

    def render(*controls):
        page.clean()
        page.add(*controls)
        page.update()

    def back_btn(target):
        return ft.TextButton("< Volver", on_click=lambda e: target())

    def all_turnos():
        return [t for (t,) in q("SELECT id FROM turnos ORDER BY id")]

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
        turnos = all_turnos()
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
                run("INSERT OR IGNORE INTO turnos(id) VALUES (?)", (int(v),))
                settings["turno_alta"] = int(v)
                show_students()

        def delete_turno(t):
            run("DELETE FROM turnos WHERE id=?", (t,))
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
                run(
                    "INSERT INTO students(name, turno) VALUES (?, ?)",
                    (name_in.value.strip(), int(turno_dd.value)),
                )
                show_students()

        def delete_student(sid):
            run("DELETE FROM results WHERE student_id=?", (sid,))
            run("DELETE FROM students WHERE id=?", (sid,))
            show_students()

        listing = []
        for t in turnos:
            studs = q("SELECT id, name FROM students WHERE turno=? ORDER BY id", (t,))
            header = [ft.Text(f"Turno {t}", size=18, weight=ft.FontWeight.BOLD)]
            if not studs:
                header.append(ft.TextButton("Borrar turno", on_click=lambda e, t=t: delete_turno(t)))
            listing.append(ft.Row(header))
            for sid, sname in studs:
                listing.append(
                    ft.Row(
                        [
                            ft.Text(sname, expand=True),
                            ft.TextButton("Borrar", on_click=lambda e, i=sid: delete_student(i)),
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
                run(
                    "INSERT INTO exercises(day, name, passes) VALUES (?,?,?)",
                    (today(), name_in.value.strip(), int(passes_in.value)),
                )
                show_today()

        def delete(eid):
            run("DELETE FROM results WHERE exercise_id=?", (eid,))
            run("DELETE FROM exercises WHERE id=?", (eid,))
            show_today()

        rows = [
            ft.Row(
                [
                    ft.ElevatedButton(
                        f"{n} ({p} pasadas)",
                        expand=True,
                        on_click=lambda e, i=i: show_run(i),
                    ),
                    ft.TextButton("Borrar", on_click=lambda e, i=i: delete(i)),
                ]
            )
            for i, n, p in q(
                "SELECT id, name, passes FROM exercises WHERE day=? ORDER BY id", (today(),)
            )
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
        ex_name, passes = q("SELECT name, passes FROM exercises WHERE id=?", (ex_id,))[0]
        turnos = all_turnos()
        tf = settings["turno_filtro"]
        if tf not in turnos:
            tf = None
            settings["turno_filtro"] = None

        if tf is None:
            students = q("SELECT id, name, turno FROM students ORDER BY turno, id")
        else:
            students = q(
                "SELECT id, name, turno FROM students WHERE turno=? ORDER BY id", (tf,)
            )
        counts = {
            sid: q(
                "SELECT COUNT(*) FROM results WHERE exercise_id=? AND student_id=?",
                (ex_id, sid),
            )[0][0]
            for sid, _, _ in students
        }
        by_id = {sid: (sname, t) for sid, sname, t in students}

        # quien sigue: primero el alumno elegido a mano, si no el orden automatico
        current = None
        chosen = settings["elegido"].get(ex_id)
        if chosen in by_id and counts[chosen] < passes:
            current = (chosen, by_id[chosen][0], counts[chosen] + 1)
        elif students:
            if settings["round_robin"]:
                lowest = min(counts.values())
                if lowest < passes:
                    for sid, sname, _ in students:
                        if counts[sid] == lowest:
                            current = (sid, sname, lowest + 1)
                            break
            else:
                for sid, sname, _ in students:
                    if counts[sid] < passes:
                        current = (sid, sname, counts[sid] + 1)
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
                run(
                    "INSERT INTO results(exercise_id, student_id, pass_no, seconds) VALUES (?,?,?,?)",
                    (ex_id, current[0], current[2], elapsed),
                )
                show_run(ex_id)

        btn.on_click = toggle

        def undo(e):
            last = q(
                "SELECT id FROM results WHERE exercise_id=? ORDER BY id DESC LIMIT 1", (ex_id,)
            )
            if last:
                run("DELETE FROM results WHERE id=?", (last[0][0],))
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
                    str(sid), f"{sname} (T{t}) {counts[sid]}/{passes}"
                )
                for sid, sname, t in students
            ],
            value=str(chosen) if chosen in by_id else "auto",
            on_change=set_student,
        )

        if current:
            sname, st = by_id[current[0]][0], by_id[current[0]][1]
            head = [
                ft.Text(f"{sname}  (Turno {st})", size=26, weight=ft.FontWeight.BOLD),
                ft.Text(f"Pasada {current[2]} de {passes}", size=18),
                label,
                btn,
            ]
        elif not students:
            head = [ft.Text("No hay alumnos en esta seleccion.")]
        else:
            head = [ft.Text("Ejercicio completo", size=24, weight=ft.FontWeight.BOLD)]

        summary = []
        for sid, sname, t in students:
            times = q(
                "SELECT seconds FROM results WHERE exercise_id=? AND student_id=? ORDER BY pass_no",
                (ex_id, sid),
            )
            summary.append(
                ft.Text(f"[T{t}] {sname}: " + "  |  ".join(fmt(x[0]) for x in times))
            )

        render(
            ft.TextButton("< Volver", on_click=leave),
            ft.Text(ex_name, size=22, weight=ft.FontWeight.BOLD),
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

        def do_day(e):
            out.value = f"Guardado en:\n{export_day(today())}"
            out.update()

        def do_all(e):
            out.value = f"Guardado en:\n{export_all()}"
            out.update()

        render(
            back_btn(show_home),
            ft.Text("Exportar a Excel", size=24, weight=ft.FontWeight.BOLD),
            ft.ElevatedButton("Exportar el dia de hoy", on_click=do_day, width=300),
            ft.ElevatedButton("Exportar todos los dias", on_click=do_all, width=300),
            out,
        )

    show_home()


ft.app(main)