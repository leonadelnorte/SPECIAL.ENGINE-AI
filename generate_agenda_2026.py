from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook


def build_agenda(output_path: Path) -> Path:
    days_in_month = {
        "Enero": 31,
        "Febrero": 28,
        "Marzo": 31,
        "Abril": 30,
        "Mayo": 31,
        "Junio": 30,
        "Julio": 31,
        "Agosto": 31,
        "Septiembre": 30,
        "Octubre": 31,
        "Noviembre": 30,
        "Diciembre": 31,
    }

    full_moons = {
        "2026-01-03",
        "2026-02-01",
        "2026-03-03",
        "2026-04-01",
        "2026-04-30",
        "2026-05-30",
        "2026-06-28",
        "2026-07-28",
        "2026-08-26",
        "2026-09-24",
        "2026-10-24",
        "2026-11-22",
        "2026-12-22",
    }

    months = list(days_in_month.keys())

    workbook = Workbook()
    workbook.remove(workbook.active)

    monthly_sheet = workbook.create_sheet("Mensual")
    monthly_sheet.append(
        [
            "Mes",
            "Día",
            "Luna llena",
            "Empresa",
            "Asociación",
            "Trabajo físico",
            "Web/Tienda",
            "IA/Formación",
            "Ejercicio (plancha/comba)",
            "Idea / Motivación / Desarrollo",
        ]
    )

    for month_index, month_name in enumerate(months, start=1):
        for day in range(1, days_in_month[month_name] + 1):
            iso_date = date(2026, month_index, day).isoformat()
            monthly_sheet.append(
                [
                    month_name,
                    day,
                    "Sí" if iso_date in full_moons else "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                ]
            )

    daily_sheet = workbook.create_sheet("Diario")
    daily_sheet.append(
        [
            "Fecha",
            "Empresa",
            "Asociación",
            "Trabajo físico",
            "Web/Tienda",
            "IA/Formación",
            "Ejercicio (plancha/comba)",
            "Idea / Motivación / Desarrollo",
        ]
    )

    for _ in range(1, 367):
        daily_sheet.append(["", "", "", "", "", "", "", ""])

    trimesters_sheet = workbook.create_sheet("Trimestres")
    trimesters_sheet.append(["Trimestre", "Meses incluidos", "Notas clave"])
    trimesters_sheet.append(["T1", "Enero - Febrero - Marzo", "Modelos 130 / 303 abril"])
    trimesters_sheet.append(["T2", "Abril - Mayo - Junio", "Evento 20-21-22 mayo"])
    trimesters_sheet.append(["T3", "Julio - Agosto - Septiembre", "Revisión anual progresiva"])
    trimesters_sheet.append(["T4", "Octubre - Noviembre - Diciembre", "Cierre fiscal"])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return output_path


if __name__ == "__main__":
    output = build_agenda(Path("AGENDA_EMPRESARIAL_2026.xlsx"))
    print(output)
