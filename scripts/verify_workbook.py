#!/usr/bin/env python3
"""Rellena una copia de prueba del libro con respuestas conocidas, la
recalcula con LibreOffice en modo headless, y verifica que cada celda
calculada coincide con el valor esperado. Borra la copia de prueba al
finalizar (con éxito o con error).

Uso:
    python3 scripts/verify_workbook.py
"""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parent.parent
SRC = REPO_ROOT / "dist" / "Autoevaluacion-LPDP.xlsx"

# Orden de las 19 preguntas en la hoja Autoevaluación (según build_workbook.py)
KEYS_AUTOEVAL = [
    "org.politica", "org.delegado", "org.lineamientos_arcopol", "org.registro",
    "org.eipd", "org.capacitacion", "der.aviso", "der.formularios",
    "der.plazos", "tec.antivirus", "tec.cifrado", "tec.mfa",
    "tec.respaldos", "tec.vulnerabilidades", "tec.monitoreo", "fis.acceso",
    "tra.contratos", "tra.internacional", "tra.notificacion72",
]

PRIORIDAD = {
    "org.politica": "Alta", "org.delegado": "Alta", "org.lineamientos_arcopol": "Alta",
    "org.registro": "Alta", "org.eipd": "Media", "org.capacitacion": "Media",
    "der.aviso": "Alta", "der.formularios": "Media", "der.plazos": "Alta",
    "tec.antivirus": "Alta", "tec.cifrado": "Alta", "tec.mfa": "Alta",
    "tec.respaldos": "Media", "tec.vulnerabilidades": "Media",
    "tec.monitoreo": "Alta", "fis.acceso": "Media", "tra.contratos": "Alta",
    "tra.internacional": "Media", "tra.notificacion72": "Alta",
}

CATEGORIA = {
    "org.politica": "Organización", "org.delegado": "Organización",
    "org.lineamientos_arcopol": "Organización", "org.registro": "Organización",
    "org.eipd": "Organización", "org.capacitacion": "Organización",
    "der.aviso": "Derechos", "der.formularios": "Derechos",
    "der.plazos": "Derechos", "tec.antivirus": "Técnicas",
    "tec.cifrado": "Técnicas", "tec.mfa": "Técnicas",
    "tec.respaldos": "Técnicas", "tec.vulnerabilidades": "Técnicas",
    "tec.monitoreo": "Técnicas", "fis.acceso": "Entorno y terceros",
    "tra.contratos": "Entorno y terceros",
    "tra.internacional": "Entorno y terceros",
    "tra.notificacion72": "Entorno y terceros",
}

# Mezcla de respuestas conocida, cubriendo Sí / Parcialmente / No / No sé
# y dejando dos preguntas sin responder para probar el caso en blanco.
RESPUESTAS = {
    "org.politica": "Sí",
    "org.delegado": "Sí",
    "org.lineamientos_arcopol": "No",
    "org.registro": "Parcialmente",
    "org.eipd": "No sé",
    "org.capacitacion": "No",
    "der.aviso": "Sí",
    "der.formularios": "Parcialmente",
    "der.plazos": "No",
    "tec.antivirus": "Sí",
    "tec.cifrado": "No",
    "tec.mfa": "Parcialmente",
    "tec.respaldos": "Sí",
    "tec.vulnerabilidades": "No sé",
    "tec.monitoreo": "No",
    "fis.acceso": "Sí",
    "tra.contratos": "No",
    "tra.internacional": "",  # sin responder
    "tra.notificacion72": "",  # sin responder
}

SCORE = {"Sí": 1, "Parcialmente": 0.5, "No": 0, "No sé": 0, "": None}


def expected_values():
    total = sum(v for k in KEYS_AUTOEVAL for v in [SCORE[RESPUESTAS[k]]] if v is not None)
    overall_pct = total / len(KEYS_AUTOEVAL)

    cat_pct = {}
    for cat in ["Organización", "Derechos", "Técnicas", "Entorno y terceros"]:
        keys = [k for k in KEYS_AUTOEVAL if CATEGORIA[k] == cat]
        scores = [SCORE[RESPUESTAS[k]] or 0 for k in keys]
        cat_pct[cat] = sum(scores) / len(keys)

    gaps_alta = sum(
        1 for k in KEYS_AUTOEVAL
        if PRIORIDAD[k] == "Alta" and RESPUESTAS[k] != "Sí"
    )
    partial_count = sum(1 for k in KEYS_AUTOEVAL if RESPUESTAS[k] == "Parcialmente")

    estado_map = {"Sí": "Cumplido", "Parcialmente": "Parcial", "No": "Pendiente",
                  "No sé": "Por confirmar", "": "Sin responder"}
    estados = {k: estado_map[RESPUESTAS[k]] for k in KEYS_AUTOEVAL}

    return dict(overall_pct=overall_pct, cat_pct=cat_pct, gaps_alta=gaps_alta,
                partial_count=partial_count, estados=estados)


def recalc_with_libreoffice(path: Path, workdir: Path):
    outdir = workdir / "recalced"
    outdir.mkdir()
    # Usamos el filtro de conversión de LibreOffice, que recalcula todas las
    # fórmulas del documento al cargarlo (las fórmulas escritas por openpyxl
    # no llevan valor cacheado, por lo que LibreOffice las evalúa desde cero).
    cmd = [
        "soffice", "--headless", "--norestore", "--calc",
        "--convert-to", "xlsx",
        "--outdir", str(outdir),
        str(path),
    ]
    env = {"HOME": str(workdir)}
    result = subprocess.run(cmd, capture_output=True, text=True, env=env, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice falló: {result.stdout}\n{result.stderr}")
    out_path = outdir / path.name
    if not out_path.exists():
        raise RuntimeError(f"LibreOffice no generó {out_path}")
    return out_path


def main():
    if not SRC.exists():
        print(f"No existe {SRC}. Ejecute primero build_workbook.py", file=sys.stderr)
        sys.exit(1)

    tmpdir = Path(tempfile.mkdtemp(prefix="lpdp-verify-"))
    test_path = tmpdir / "test_copy.xlsx"
    shutil.copy(SRC, test_path)

    errors = []
    try:
        wb = load_workbook(test_path)
        ws = wb["Autoevaluación"]
        key_to_row = {}
        for row in range(2, 2 + len(KEYS_AUTOEVAL)):
            key = ws.cell(row=row, column=9).value  # columna I = Clave
            key_to_row[key] = row
        assert set(key_to_row) == set(KEYS_AUTOEVAL), (
            f"Claves inesperadas en Autoevaluación: {set(key_to_row) ^ set(KEYS_AUTOEVAL)}")

        for key, resp in RESPUESTAS.items():
            row = key_to_row[key]
            if resp:
                ws.cell(row=row, column=5, value=resp)
        wb.save(test_path)

        recalced_path = recalc_with_libreoffice(test_path, tmpdir)

        wb2 = load_workbook(recalced_path, data_only=True)
        ws_auto = wb2["Autoevaluación"]
        ws_res = wb2["Resultados"]
        ws_plan = wb2["Plan de acción"]

        exp = expected_values()

        # --- Puntaje por pregunta en Autoevaluación ---
        key_to_row2 = {}
        for row in range(2, 2 + len(KEYS_AUTOEVAL)):
            key = ws_auto.cell(row=row, column=9).value
            key_to_row2[key] = row
        for key, resp in RESPUESTAS.items():
            row = key_to_row2[key]
            got = ws_auto.cell(row=row, column=6).value
            want = SCORE[resp]
            if want is None:
                if got not in (None, ""):
                    errors.append(f"Puntaje {key}: esperado en blanco, obtuvo {got!r}")
            else:
                if got != want:
                    errors.append(f"Puntaje {key}: esperado {want}, obtuvo {got!r}")

            que_implica = ws_auto.cell(row=row, column=7).value
            if resp == "Sí":
                if que_implica not in (None, ""):
                    errors.append(f"Qué implica {key}: esperado vacío para 'Sí', obtuvo {que_implica!r}")
            elif resp == "":
                if que_implica not in (None, ""):
                    errors.append(f"Qué implica {key}: esperado vacío sin respuesta, obtuvo {que_implica!r}")
            else:
                if not que_implica:
                    errors.append(f"Qué implica {key}: esperado texto no vacío para '{resp}', obtuvo {que_implica!r}")

        # --- Resultados ---
        overall = ws_res.cell(row=4, column=2).value
        if overall is None or abs(overall - exp["overall_pct"]) > 1e-9:
            errors.append(f"% general: esperado {exp['overall_pct']}, obtuvo {overall!r}")

        cat_row_start = 8
        cats = ["Organización", "Derechos", "Técnicas", "Entorno y terceros"]
        for i, cat in enumerate(cats):
            r = cat_row_start + i
            name = ws_res.cell(row=r, column=1).value
            val = ws_res.cell(row=r, column=2).value
            if name != cat:
                errors.append(f"Fila de categoría {r}: esperado {cat!r}, obtuvo {name!r}")
                continue
            if val is None or abs(val - exp["cat_pct"][cat]) > 1e-9:
                errors.append(f"% categoría {cat}: esperado {exp['cat_pct'][cat]}, obtuvo {val!r}")

        # localizar filas de brechas / parcialmente por etiqueta, por robustez
        gaps_val = partial_val = None
        for r in range(1, 30):
            label = ws_res.cell(row=r, column=1).value
            if isinstance(label, str) and "prioridad Alta" in label:
                gaps_val = ws_res.cell(row=r, column=2).value
            if isinstance(label, str) and "Parcialmente" in label:
                partial_val = ws_res.cell(row=r, column=2).value
        if gaps_val != exp["gaps_alta"]:
            errors.append(f"Brechas Alta: esperado {exp['gaps_alta']}, obtuvo {gaps_val!r}")
        if partial_val != exp["partial_count"]:
            errors.append(f"Conteo Parcialmente: esperado {exp['partial_count']}, obtuvo {partial_val!r}")

        # --- Plan de acción: Estado por clave ---
        for row in range(2, 2 + len(KEYS_AUTOEVAL)):
            key = ws_plan.cell(row=row, column=8).value
            estado = ws_plan.cell(row=row, column=7).value
            want = exp["estados"].get(key)
            if want is None:
                errors.append(f"Plan de acción: clave inesperada {key!r} en fila {row}")
                continue
            if estado != want:
                errors.append(f"Plan de acción estado {key}: esperado {want!r}, obtuvo {estado!r}")

        plan_keys = [ws_plan.cell(row=r, column=8).value for r in range(2, 2 + len(KEYS_AUTOEVAL))]
        if set(plan_keys) != set(KEYS_AUTOEVAL):
            errors.append("Plan de acción no contiene exactamente las 19 claves esperadas")
        plan_prioridades = [ws_plan.cell(row=r, column=4).value for r in range(2, 2 + len(KEYS_AUTOEVAL))]
        if plan_prioridades != sorted(plan_prioridades, key=lambda p: 0 if p == "Alta" else 1):
            errors.append("Plan de acción no está ordenado por prioridad (Alta primero)")

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    if errors:
        print("VERIFICACIÓN FALLIDA:\n" + "\n".join(f"  - {e}" for e in errors), file=sys.stderr)
        sys.exit(1)
    print(f"OK: todas las celdas calculadas coinciden con lo esperado "
          f"({len(KEYS_AUTOEVAL)} preguntas, % general {expected_values()['overall_pct']:.1%}).")


if __name__ == "__main__":
    main()
