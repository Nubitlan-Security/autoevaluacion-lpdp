#!/usr/bin/env python3
"""Genera dist/Autoevaluacion-LPDP.xlsx a partir de los controles definidos en este script.

El libro de Excel es un artefacto de compilación: la fuente de verdad es este
script. Para regenerarlo:

    python3 scripts/build_workbook.py

Requiere openpyxl (pip install openpyxl).
"""
import math
import os

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Datos de los 19 controles
# ---------------------------------------------------------------------------
# Cada "base_legal" ha sido verificado contra los documentos en referencias/:
#   - Políticas N.º 001-0309025-DPDP (Agencia de Ciberseguridad del Estado)
#   - Diario Oficial, Tomo 452, Nº 146, 11 de agosto de 2026:
#       "Lineamientos para el Delegado de Protección de Datos Personales"
#       "Normativa para el Desarrollo del Procedimiento Administrativo
#        Sancionador en el Ámbito de la Ley para la Protección de Datos
#        Personales"
#   - "Reformas a la Ley para la Protección de Datos Personales" (iniciativa
#     del Presidente de la República presentada el 17 de septiembre de 2026,
#     aprobada por la Asamblea Legislativa ese mismo día con 57 votos).
#
# IMPORTANTE — efecto de la reforma de 2026:
#   Art. 13 de la reforma DEROGA los artículos 15 y 17 de la LPDP, que eran
#   la base de la obligación de nombrar un Delegado de Protección de Datos
#   Personales. Tras la reforma, esa obligación subsiste únicamente para las
#   instituciones públicas (Art. 47 LPDP reformado). Para las entidades
#   privadas, las funciones que correspondían al Delegado pasan directamente
#   al sujeto obligado (Art. 16 LPDP reformado).
#
#   En consecuencia, los "Lineamientos para el Delegado" de la ACE dejan de
#   ser base legal aplicable a una empresa privada, y los controles que antes
#   los citaban se reanclaron a los artículos reformados de la LPDP.
#
#   La reforma entra en vigencia ocho días después de su publicación en el
#   Diario Oficial (Art. 14 de la reforma). Este libro ya está redactado
#   conforme al texto reformado; vea REFORMA_NOTA.

CONTROLES = [
    dict(
        key="org.politica",
        categoria="Organización",
        prioridad="Alta",
        pregunta="¿Su empresa cuenta con una política interna escrita de "
        "protección de datos personales, conocida por el personal?",
        base_legal="Políticas ACE N.º 001-0309025-DPDP, Art. 4 (Medidas "
        "Organizativas, lit. a) y Art. 6 lit. a)",
        que_implica="Sin una política escrita, no existe una referencia común "
        "sobre cómo debe tratarse un dato personal dentro de la empresa, y "
        "ante una inspección de la ACE no podrá demostrar que el tema se "
        "gestiona de forma ordenada.",
        accion="Redactar y aprobar una política interna de protección de "
        "datos personales, y divulgarla a todo el personal.",
        responsable="Entidad",
    ),
    dict(
        key="org.delegado",
        categoria="Organización",
        prioridad="Alta",
        pregunta="¿Tiene designada internamente a una persona o área "
        "responsable de atender las solicitudes ARCO-POL y de orientar a sus "
        "dependencias y proveedores en la materia?",
        base_legal="LPDP Art. 16 lit. b) (texto reformado en 2026). Nota: la "
        "reforma derogó el Art. 15 LPDP, por lo que una empresa privada ya "
        "NO está obligada a nombrar un Delegado formal; la obligación "
        "subsiste solo para instituciones públicas (Art. 47 LPDP reformado)",
        que_implica="La empresa ya no está obligada a nombrar un Delegado "
        "formal, pero las solicitudes ARCO-POL siguen llegándole a ella y "
        "deben responderse en veinte días hábiles. Si nadie tiene el encargo "
        "de recibirlas, el plazo corre igual y normalmente se vence.",
        accion="Designar internamente por escrito a una persona o área "
        "responsable de recibir y resolver las solicitudes ARCO-POL y de "
        "orientar a las demás áreas y a los proveedores. Ya no se requiere "
        "el nombramiento formal de un Delegado ni su certificación ante la "
        "ACE, salvo para instituciones públicas.",
        responsable="Entidad",
    ),
    dict(
        key="org.lineamientos_arcopol",
        categoria="Organización",
        prioridad="Alta",
        pregunta="¿Ha establecido lineamientos internos escritos que definan "
        "cómo se reciben, tramitan y resuelven las solicitudes ARCO-POL de "
        "los titulares?",
        base_legal="LPDP Art. 16 lit. d) (texto reformado en 2026)",
        que_implica="Esta es una obligación nueva y expresa para toda "
        "empresa: tener por escrito cómo se gestionan estas solicitudes. Sin "
        "esos lineamientos, cada solicitud se resuelve de forma improvisada y "
        "distinta, y la empresa no puede demostrar que cuenta con un "
        "procedimiento.",
        accion="Redactar y aprobar lineamientos internos que definan quién "
        "recibe cada solicitud ARCO-POL, cómo se registra, quién la resuelve, "
        "en qué plazos y cómo se documenta la respuesta al titular.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="org.registro",
        categoria="Organización",
        prioridad="Alta",
        pregunta="¿Mantiene un Registro de Actividades de Tratamiento que "
        "documente qué datos personales recopila, para qué los usa y dónde "
        "los almacena?",
        base_legal="Políticas ACE Art. 4 (Medidas Organizativas, lit. d) y "
        "Art. 6 lit. a)",
        que_implica="Sin este registro, la empresa no puede demostrar de "
        "forma ordenada qué tratamientos de datos realiza, lo que dificulta "
        "responder solicitudes ARCO-POL y atender cualquier requerimiento "
        "de la ACE.",
        accion="Elaborar un inventario de los tratamientos de datos "
        "personales (finalidad, tipo de dato, base de tratamiento, plazo de "
        "conservación, a quién se transfiere) y mantenerlo actualizado.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="org.eipd",
        categoria="Organización",
        prioridad="Media",
        pregunta="Para los tratamientos de mayor riesgo (por ejemplo, datos "
        "sensibles o de menores), ¿ha realizado una Evaluación de Impacto "
        "en la Privacidad (EIPD)?",
        base_legal="Políticas ACE Art. 4 (Medidas Organizativas, lit. e)",
        que_implica="Sin una EIPD, la empresa no ha identificado formalmente "
        "los riesgos de los tratamientos más sensibles que realiza, ni las "
        "medidas para mitigarlos.",
        accion="Identificar los tratamientos de mayor riesgo y realizar una "
        "evaluación de impacto que documente los riesgos y las medidas de "
        "mitigación adoptadas.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="org.capacitacion",
        categoria="Organización",
        prioridad="Media",
        pregunta="¿Su personal recibe, al menos una vez al año, "
        "capacitación sobre protección de datos personales y el ejercicio "
        "de los derechos ARCO-POL?",
        base_legal="Políticas ACE Art. 4 (Medidas Organizativas, lit. c — "
        "Capacitación del Personal)",
        que_implica="Un personal sin capacitación es más propenso a cometer "
        "errores que generan una vulneración de datos, por ejemplo enviar "
        "información a un destinatario equivocado.",
        accion="Programar una capacitación anual sobre protección de datos "
        "y derechos ARCO-POL para todo el personal, incluyendo al personal "
        "de nuevo ingreso, y conservar la evidencia de la capacitación.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="der.aviso",
        categoria="Derechos",
        prioridad="Alta",
        pregunta="¿Cuenta con un Aviso de Privacidad publicado en su sitio "
        "web y/o en sus establecimientos físicos, autorizado y documentado "
        "internamente?",
        base_legal="LPDP Art. 24, en especial el lit. f) — indicación de los "
        "medios o mecanismos para presentar las solicitudes de derechos "
        "ARCO-POL (texto reformado en 2026)",
        que_implica="Sin un Aviso de Privacidad visible, sus clientes y "
        "empleados no están siendo informados de cómo se usan sus datos, lo "
        "cual es una de las obligaciones más básicas frente a los "
        "titulares. Además, tras la reforma el aviso debe decir "
        "expresamente por qué medio se presentan las solicitudes ARCO-POL.",
        accion="Redactar y publicar el Aviso de Privacidad en el sitio web "
        "y/o el establecimiento, verificando que incluya todos los "
        "contenidos del Art. 24 de la LPDP y, en particular, la indicación "
        "de los medios para presentar solicitudes ARCO-POL. Conserve la "
        "documentación de su publicación.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="der.formularios",
        categoria="Derechos",
        prioridad="Media",
        pregunta="¿Tiene un medio claro y disponible (formulario, correo o "
        "ventanilla) para que un titular presente su solicitud ARCO-POL "
        "directamente ante su empresa?",
        base_legal="LPDP Art. 18 (la solicitud se presenta ante el sujeto "
        "obligado) y Art. 24 lit. f) (indicación de los medios para "
        "presentarla), ambos con texto reformado en 2026",
        que_implica="Tras la reforma, las solicitudes ya no se canalizan a "
        "través de un Delegado: llegan directamente a su empresa. Si el "
        "titular no encuentra un medio claro para presentarlas, esto puede "
        "tomarse como un obstáculo a su ejercicio, aunque la empresa esté "
        "dispuesta a atenderlo de forma informal.",
        accion="Habilitar y publicar al menos un medio concreto para recibir "
        "solicitudes ARCO-POL (formulario en el sitio web, correo "
        "electrónico o ventanilla física), e indicarlo expresamente en el "
        "Aviso de Privacidad.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="der.plazos",
        categoria="Derechos",
        prioridad="Alta",
        pregunta="¿Cuenta con un procedimiento interno que controle los "
        "plazos para resolver las solicitudes ARCO-POL de sus clientes o "
        "empleados?",
        base_legal="LPDP Art. 20 (veinte días hábiles, prorrogables por otros "
        "veinte), Art. 18 (prevención, diez días hábiles), Art. 19 "
        "(incompetencia, cinco días hábiles), Art. 21 (rectificación, cinco "
        "días hábiles), Art. 22 (denegatoria, tres días hábiles) y Art. 30 "
        "(revocación, cinco días hábiles); todos con texto reformado en 2026",
        que_implica="Tras la reforma, estos plazos corren contra su empresa "
        "directamente y no contra un Delegado. Una solicitud sin respuesta "
        "dentro del plazo legal es en sí misma un incumplimiento, aunque "
        "después se resuelva de forma favorable.",
        accion="Implementar una bitácora o sistema de seguimiento que "
        "registre la fecha de cada solicitud ARCO-POL, el plazo que le "
        "aplica y una alerta antes de su vencimiento. Como referencia: "
        "veinte días hábiles para la respuesta, tres para una denegatoria y "
        "cinco para los avisos de rectificación, revocación o "
        "incompetencia.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="tec.antivirus",
        categoria="Técnicas",
        prioridad="Alta",
        pregunta="¿Todos los equipos de la empresa cuentan con protección "
        "contra programas maliciosos administrada desde una consola "
        "central?",
        base_legal="Políticas ACE Art. 4 (Medidas Técnicas, lit. e — "
        "Protección de Infraestructura)",
        que_implica="Un equipo sin protección centralizada puede ser "
        "infectado sin que nadie se entere hasta que ya ocurrió una fuga de "
        "datos.",
        accion="Instalar una solución antivirus/antimalware con consola "
        "central en todos los equipos, y verificar periódicamente que esté "
        "activa y actualizada.",
        responsable="Apoyo técnico",
    ),
    dict(
        key="tec.cifrado",
        categoria="Técnicas",
        prioridad="Alta",
        pregunta="¿Los equipos (laptops, discos) y los respaldos que "
        "contienen datos personales están cifrados?",
        base_legal="Políticas ACE Art. 4 (Medidas Técnicas, lit. b — "
        "Cifrado de Datos)",
        que_implica="Una laptop robada sin cifrado es una vulneración "
        "notificable. Con cifrado, normalmente no lo es.",
        accion="Activar el cifrado de disco completo en todos los equipos "
        "que almacenen datos personales y cifrar los respaldos, incluyendo "
        "los que se guarden fuera de las instalaciones.",
        responsable="Apoyo técnico",
    ),
    dict(
        key="tec.mfa",
        categoria="Técnicas",
        prioridad="Alta",
        pregunta="¿El correo institucional y los sistemas críticos de la "
        "empresa requieren doble factor de autenticación (2FA/MFA)?",
        base_legal="Políticas ACE Art. 4 (Medidas Técnicas, lit. a — "
        "Control de Acceso)",
        que_implica="Una contraseña filtrada o adivinada, sin un segundo "
        "factor, es suficiente para que un tercero acceda al correo o a los "
        "sistemas y extraiga datos personales.",
        accion="Activar la autenticación de doble factor en el correo "
        "institucional y en todos los sistemas que almacenen o den acceso a "
        "datos personales.",
        responsable="Apoyo técnico",
    ),
    dict(
        key="tec.respaldos",
        categoria="Técnicas",
        prioridad="Media",
        pregunta="¿Realiza respaldos periódicos de la información y ha "
        "probado que estos se pueden restaurar correctamente?",
        base_legal="Políticas ACE Art. 4 (Medidas Técnicas, lit. d — Copias "
        "de Seguridad y Recuperación)",
        que_implica="Un respaldo que nunca se ha probado puede fallar justo "
        "cuando se necesita, por ejemplo tras un ataque de ransomware, "
        "dejando a la empresa sin forma de garantizar la disponibilidad de "
        "los datos.",
        accion="Definir una frecuencia de respaldo acorde al volumen de "
        "información y realizar al menos una prueba de restauración al "
        "año, documentando el resultado.",
        responsable="Apoyo técnico",
    ),
    dict(
        key="tec.vulnerabilidades",
        categoria="Técnicas",
        prioridad="Media",
        pregunta="¿Realiza análisis periódicos de vulnerabilidades sobre "
        "los sistemas que almacenan o procesan datos personales?",
        base_legal="Políticas ACE Art. 4 (Medidas Técnicas, lit. f — "
        "Análisis de Vulnerabilidades y Pruebas de Penetración)",
        que_implica="Sin análisis periódicos, la empresa no sabe qué "
        "puertas quedaron abiertas hasta que alguien más las encuentra "
        "primero.",
        accion="Contratar o programar análisis de vulnerabilidades con una "
        "periodicidad definida (por ejemplo, anual) sobre los sistemas que "
        "tratan datos personales, y dar seguimiento a los hallazgos.",
        responsable="Apoyo técnico",
    ),
    dict(
        key="tec.monitoreo",
        categoria="Técnicas",
        prioridad="Alta",
        pregunta="¿Cuenta con algún mecanismo para detectar incidentes de "
        "seguridad fuera del horario laboral (fines de semana, noches)?",
        base_legal="Políticas ACE Art. 4 (Medidas Técnicas, lit. e — "
        "Protección de Infraestructura; y Medidas de Seguridad en "
        "Transferencias de Datos, lit. d — Notificación de Brechas de "
        "Seguridad)",
        que_implica="La mayoría de los incidentes de seguridad ocurren o se "
        "descubren fuera de horario. Sin monitoreo, el plazo de 72 horas "
        "para notificar una vulneración a la ACE puede empezar a correr sin "
        "que la empresa se entere.",
        accion="Implementar alertas automáticas o un servicio de monitoreo "
        "que permita detectar incidentes de seguridad en cualquier horario "
        "y activar el protocolo de respuesta correspondiente.",
        responsable="Apoyo técnico",
    ),
    dict(
        key="fis.acceso",
        categoria="Entorno y terceros",
        prioridad="Media",
        pregunta="¿El acceso físico a los servidores y a los archivos "
        "físicos con datos personales está restringido solo al personal "
        "autorizado?",
        base_legal="Políticas ACE Art. 4 (Medidas Físicas, lit. a — Control "
        "de Acceso a Instalaciones)",
        que_implica="Cualquier persona que pueda entrar sin control a donde "
        "están los servidores o los archivos físicos puede copiar, alterar "
        "o sustraer datos personales sin que quede registro.",
        accion="Restringir el acceso a las áreas donde se ubican servidores "
        "o archivos físicos con datos personales, mediante llave, tarjeta u "
        "otro control, y llevar un registro de ingreso.",
        responsable="Entidad",
    ),
    dict(
        key="tra.contratos",
        categoria="Entorno y terceros",
        prioridad="Alta",
        pregunta="¿Los contratos con sus proveedores que tratan datos "
        "personales por cuenta suya (nómina, hosting, mercadeo, etc.) "
        "incluyen cláusulas de tratamiento de datos y confidencialidad?",
        base_legal="Políticas ACE Art. 4 (Medidas de Seguridad en "
        "Transferencias de Datos, lit. b — Contratos de Confidencialidad y "
        "Contratos para la Transferencia de Datos)",
        que_implica="Sin cláusulas contractuales, la empresa sigue siendo "
        "responsable ante la ACE por lo que haga el proveedor con los "
        "datos, pero no tiene forma de exigirle nada ni de repartir esa "
        "responsabilidad.",
        accion="Revisar los contratos vigentes con proveedores que traten "
        "datos personales por cuenta de la empresa e incorporar cláusulas "
        "de confidencialidad y de tratamiento de datos; incluirlas también "
        "en los contratos nuevos.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="tra.internacional",
        categoria="Entorno y terceros",
        prioridad="Media",
        pregunta="¿Sabe en qué país o países se almacenan físicamente los "
        "datos personales que su empresa trata (por ejemplo, el país donde "
        "está el centro de datos de su proveedor de nube o correo)?",
        base_legal="Políticas ACE Art. 4 (Medidas de Seguridad en "
        "Transferencias de Datos, lit. c — Transferencias Internacionales "
        "Seguras)",
        que_implica="Si no sabe dónde están sus datos, no puede verificar "
        "que el país de destino ofrezca un nivel de protección "
        "equivalente, un requisito para cualquier transferencia "
        "internacional de datos.",
        accion="Identificar en qué país(es) se almacenan los datos "
        "personales con cada proveedor (nube, correo, respaldo) y verificar "
        "las condiciones de protección que ofrece ese país o el contrato "
        "con el proveedor.",
        responsable="Responsable interno o asesor legal",
    ),
    dict(
        key="tra.notificacion72",
        categoria="Entorno y terceros",
        prioridad="Alta",
        pregunta="¿Cuenta con un procedimiento escrito para notificar una "
        "vulneración de datos personales a la ACE, a la Fiscalía y a los "
        "titulares afectados dentro de un plazo máximo de 72 horas?",
        base_legal="Políticas ACE Art. 4 (Medidas de Seguridad en "
        "Transferencias de Datos, lit. d — Notificación de Brechas de "
        "Seguridad)",
        que_implica="Sin un procedimiento previamente definido, las 72 "
        "horas para notificar se consumen decidiendo qué hacer, en lugar de "
        "actuando, lo que aumenta el riesgo de notificar tarde o de forma "
        "incompleta.",
        accion="Documentar un procedimiento de respuesta a incidentes que "
        "defina quién decide, a quién se notifica (ACE, Fiscalía General de "
        "la República, titulares) y en qué plazo, de modo que pueda "
        "activarse de inmediato.",
        responsable="Responsable interno o asesor legal",
    ),
]

assert len(CONTROLES) == 19
assert {c["key"] for c in CONTROLES} == {
    "org.politica", "org.delegado", "org.lineamientos_arcopol", "org.registro",
    "org.eipd", "org.capacitacion", "der.aviso", "der.formularios",
    "der.plazos", "tec.antivirus", "tec.cifrado", "tec.mfa",
    "tec.respaldos", "tec.vulnerabilidades", "tec.monitoreo", "fis.acceso",
    "tra.contratos", "tra.internacional", "tra.notificacion72",
}

CATEGORIAS = ["Organización", "Derechos", "Técnicas", "Entorno y terceros"]

DISCLAIMER = ("No constituye asesoría legal ni certificación oficial de la "
              "ACE.")

# Nota sobre el estado de la reforma de 2026. Se muestra en la hoja
# Instrucciones y se repite en el README del repositorio.
REFORMA_NOTA = (
    "Este libro está redactado conforme a las Reformas a la Ley para la "
    "Protección de Datos Personales aprobadas por la Asamblea Legislativa el "
    "17 de septiembre de 2026. Dichas reformas derogaron los artículos 15 y "
    "17 de la LPDP y, con ello, la obligación de las empresas privadas de "
    "nombrar un Delegado de Protección de Datos Personales; esa obligación "
    "subsiste únicamente para las instituciones públicas. Las funciones que "
    "antes correspondían al Delegado pasan directamente a la empresa. "
    "Tenga presente que la reforma entra en vigencia ocho días después de su "
    "publicación en el Diario Oficial, por lo que durante ese intervalo el "
    "texto anterior sigue siendo el vigente. Verifique la fecha de "
    "publicación antes de tomar decisiones basadas en este documento."
)

# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
FONT_NAME = "Arial"
TITLE_FONT = Font(name=FONT_NAME, size=16, bold=True, color="1F3864")
SUBTITLE_FONT = Font(name=FONT_NAME, size=11, bold=True, color="1F3864")
BODY_FONT = Font(name=FONT_NAME, size=10)
BODY_FONT_ITALIC = Font(name=FONT_NAME, size=9, italic=True, color="595959")
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
ALT_FILL = PatternFill("solid", fgColor="F2F2F2")
ALTA_FILL = PatternFill("solid", fgColor="FCE4D6")
MEDIA_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
WRAP_CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")

# ---------------------------------------------------------------------------
# Paleta semántica para el formato condicional
# ---------------------------------------------------------------------------
# Se usan los tonos clásicos de Excel (Bueno / Neutral / Incorrecto) para que
# resulten familiares. El color siempre acompaña a un texto o a un número que
# ya dice lo mismo por sí solo: nunca es el único portador del significado,
# de modo que el libro sigue siendo legible al imprimirse en blanco y negro o
# para una persona con daltonismo.
CF_VERDE = dict(fill=PatternFill("solid", fgColor="C6EFCE"),
                font=Font(name=FONT_NAME, size=10, color="006100"))
CF_AMBAR = dict(fill=PatternFill("solid", fgColor="FFEB9C"),
                font=Font(name=FONT_NAME, size=10, color="9C6500"))
CF_ROJO = dict(fill=PatternFill("solid", fgColor="FFC7CE"),
               font=Font(name=FONT_NAME, size=10, color="9C0006"))
CF_NARANJA = dict(fill=PatternFill("solid", fgColor="FCE4D6"),
                  font=Font(name=FONT_NAME, size=10, color="833C0C"))
CF_GRIS = dict(fill=PatternFill("solid", fgColor="E7E6E6"),
               font=Font(name=FONT_NAME, size=10, color="7F7F7F"))

# Estado del Plan de acción -> color. El orden importa: Excel aplica la
# primera regla que coincide.
CF_ESTADOS = [
    ("Cumplido", CF_VERDE),
    ("Parcial", CF_AMBAR),
    ("Pendiente", CF_ROJO),
    ("Por confirmar", CF_NARANJA),
    ("Sin responder", CF_GRIS),
]

# Umbrales de los porcentajes de cumplimiento, iguales en toda la hoja
# Resultados para que un color signifique siempre lo mismo.
UMBRAL_ALTO = 0.8
UMBRAL_MEDIO = 0.5


def aplicar_cf_porcentaje(ws, rango):
    """Colorea un rango de porcentajes en tres bandas: >=80%, 50-79%, <50%."""
    ws.conditional_formatting.add(rango, CellIsRule(
        operator="greaterThanOrEqual", formula=[str(UMBRAL_ALTO)],
        stopIfTrue=True, **CF_VERDE))
    ws.conditional_formatting.add(rango, CellIsRule(
        operator="greaterThanOrEqual", formula=[str(UMBRAL_MEDIO)],
        stopIfTrue=True, **CF_AMBAR))
    ws.conditional_formatting.add(rango, CellIsRule(
        operator="lessThan", formula=[str(UMBRAL_MEDIO)],
        stopIfTrue=True, **CF_ROJO))


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def estimar_alto_fila(celdas, alto_minimo=30):
    """Estima el alto de fila necesario para que no se corte el texto.

    openpyxl no puede autoajustar el alto, y dejarlo fijo recorta las
    celdas con más texto al imprimir. Se estima el número de líneas que
    ocupa cada celda envuelta, a partir de la longitud del texto y del
    ancho de su columna, y se toma la más alta.

    `celdas` es una secuencia de pares (texto, ancho_de_columna).
    """
    lineas = 1
    for texto, ancho in celdas:
        if not texto or ancho <= 0:
            continue
        # A 10 pt Arial entra algo más de un carácter por unidad de ancho.
        por_linea = max(ancho * 1.05, 1)
        lineas = max(lineas, math.ceil(len(texto) / por_linea))
    return max(alto_minimo, lineas * 12.5 + 8)


def configurar_impresion(ws, ancho_paginas=1):
    """Ajusta la hoja para que quepa a lo ancho al imprimirla o exportarla."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = ancho_paginas
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = WRAP_CENTER
        cell.border = BORDER


def footer(ws, row, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=DISCLAIMER)
    cell.font = BODY_FONT_ITALIC


# ---------------------------------------------------------------------------
# Construcción del libro
# ---------------------------------------------------------------------------

def build():
    wb = Workbook()

    ws_instr = wb.active
    ws_instr.title = "Instrucciones"
    build_instrucciones(ws_instr)

    ws_auto = wb.create_sheet("Autoevaluación")
    build_autoevaluacion(ws_auto)

    ws_res = wb.create_sheet("Resultados")
    build_resultados(ws_res)

    ws_plan = wb.create_sheet("Plan de acción")
    build_plan(ws_plan)

    wb.active = 0
    return wb


def build_instrucciones(ws):
    set_col_widths(ws, [95])
    row = 1
    ws.cell(row=row, column=1, value="Autoevaluación LPDP para empresas salvadoreñas").font = TITLE_FONT
    row += 2

    parrafos_que_es = [
        ("¿Qué es esto?", SUBTITLE_FONT),
        ("Esta es una herramienta gratuita y de código abierto que le "
         "permite a su empresa estimar, en menos de una hora, qué tan cerca "
         "está de cumplir con las medidas de seguridad exigidas por la Ley "
         "para la Protección de Datos Personales (LPDP) de El Salvador y "
         "por la normativa que ha emitido la Agencia de Ciberseguridad del "
         "Estado (ACE) para desarrollarla.", BODY_FONT),
        ("No necesita conocimientos técnicos ni jurídicos para usarla: cada "
         "pregunta está escrita en lenguaje sencillo y ordinario, pensada "
         "para quien administra o dirige un negocio, no para un "
         "especialista en seguridad de la información.", BODY_FONT),
        ("Reforma de septiembre de 2026: ya no se exige Delegado a las "
         "empresas privadas", SUBTITLE_FONT),
        (REFORMA_NOTA, BODY_FONT),
        ("En la práctica, esto significa que su empresa ya no necesita "
         "nombrar ni certificar a un Delegado, pero sí debe atender ella "
         "misma las solicitudes ARCO-POL (acceso, rectificación, "
         "cancelación, oposición, portabilidad, olvido y limitación), "
         "establecer lineamientos internos por escrito para gestionarlas, y "
         "cumplir los plazos de respuesta. Las preguntas de este libro ya "
         "reflejan ese reparto de obligaciones.", BODY_FONT),
        ("¿Cómo se usa?", SUBTITLE_FONT),
        ("1. Vaya a la hoja \"Autoevaluación\". 2. En la columna "
         "\"Respuesta\", de color amarillo, seleccione para cada pregunta "
         "una opción del listado desplegable: Sí, Parcialmente, No, o No "
         "sé. 3. Conforme responda, la columna \"Qué implica\" le explicará "
         "en lenguaje simple la consecuencia de no tener ese control "
         "resuelto. 4. Revise la hoja \"Resultados\" para ver su porcentaje "
         "general y por categoría. 5. Revise la hoja \"Plan de acción\" "
         "para ver, ordenadas por prioridad, las acciones pendientes y "
         "quién debería encargarse de cada una.", BODY_FONT),
        ("¿Cómo se califica?", SUBTITLE_FONT),
        ("Cada pregunta otorga: Sí = 1 punto, Parcialmente = 0.5 puntos, "
         "No = 0 puntos, No sé = 0 puntos. \"No sé\" se califica igual que "
         "\"No\", porque un control que nadie en la empresa puede confirmar "
         "que existe no es un control del que se pueda dar cuenta ante la "
         "ACE, un cliente o una auditoría: si nadie sabe que existe, en la "
         "práctica es como si no existiera.", BODY_FONT),
        ("Advertencia importante", SUBTITLE_FONT),
        ("Esta autoevaluación es declarativa, no evidenciada: los "
         "resultados dependen enteramente de lo que usted responda, y la "
         "herramienta no verifica ni solicita evidencia documental de "
         "ningún tipo. Por lo tanto, este documento NO constituye un "
         "informe de auditoría, NO certifica el cumplimiento de la LPDP, y "
         "NO vincula ni compromete a la Agencia de Ciberseguridad del "
         "Estado (ACE) de ninguna forma. Es únicamente una herramienta de "
         "orientación inicial para que la empresa identifique, por sí "
         "misma, sus brechas más evidentes.", BODY_FONT),
        ("Aviso legal", SUBTITLE_FONT),
        ("El contenido de este documento tiene fines informativos y de "
         "autodiagnóstico. No constituye asesoría legal ni sustituye la "
         "opinión de un profesional del derecho o de un Delegado de "
         "Protección de Datos Personales debidamente cualificado. Antes de "
         "tomar decisiones con impacto legal, contable o económico, "
         "consulte con un asesor calificado.", BODY_FONT),
        ("Licencia", SUBTITLE_FONT),
        ("Esta obra se publica bajo la licencia Creative Commons "
         "Atribución-CompartirIgual 4.0 Internacional (CC BY-SA 4.0). Puede "
         "copiarla, modificarla y redistribuirla, incluso con fines "
         "comerciales, siempre que dé crédito a Nubitlan y distribuya "
         "cualquier obra derivada bajo esta misma licencia. Vea el archivo "
         "LICENSE para el texto completo.", BODY_FONT),
        ("Fuentes verificadas", SUBTITLE_FONT),
        ("Cada referencia de \"Base legal\" en la hoja \"Autoevaluación\" "
         "fue verificada contra: la Ley para la Protección de Datos "
         "Personales (Decreto Legislativo N.º 144); las Reformas a dicha ley "
         "aprobadas el 17 de septiembre de 2026, que son la fuente de los "
         "artículos identificados como \"texto reformado en 2026\"; las "
         "Políticas de Actuación y Manejo de Datos Personales de la ACE "
         "(Políticas N.º 001-0309025-DPDP); y la Normativa para el "
         "Desarrollo del Procedimiento Administrativo Sancionador en el "
         "Ámbito de la LPDP (Diario Oficial, Tomo 452, N.º 146, 11 de agosto "
         "de 2026). Los Lineamientos para el Delegado de Protección de Datos "
         "Personales, publicados en ese mismo Diario Oficial, ya no se citan "
         "como base legal aplicable a una empresa privada, porque "
         "desarrollaban artículos que la reforma derogó o reescribió; siguen "
         "siendo aplicables a las instituciones públicas. Si encuentra una "
         "cita que le parece incorrecta, repórtela siguiendo las "
         "instrucciones del README.", BODY_FONT),
    ]

    for text, font in parrafos_que_es:
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = font
        cell.alignment = WRAP
        ws.row_dimensions[row].height = 18 if font is SUBTITLE_FONT else max(
            30, 15 * (len(text) // 100 + 1))
        row += 1

    row += 1
    configurar_impresion(ws)
    ws.page_setup.orientation = "portrait"
    footer(ws, row, 1)


def build_autoevaluacion(ws):
    headers = ["N°", "Categoría", "Pregunta", "Base legal", "Respuesta",
               "Puntaje", "Qué implica", "Prioridad"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    ws.row_dimensions[1].height = 30

    anchos_auto = [5, 14, 46, 34, 14, 9, 46, 10, 16, 46]
    set_col_widths(ws, anchos_auto)
    ws.column_dimensions["I"].hidden = True
    ws.column_dimensions["J"].hidden = True
    ws.cell(row=1, column=9, value="Clave").font = HEADER_FONT
    ws.cell(row=1, column=10, value="Qué implica (texto)").font = HEADER_FONT

    first_row = 2
    for i, c in enumerate(CONTROLES):
        r = first_row + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=c["categoria"])
        ws.cell(row=r, column=3, value=c["pregunta"])
        ws.cell(row=r, column=4, value=c["base_legal"])
        # columna 5 (Respuesta) queda en blanco: la llena el usuario
        ws.cell(row=r, column=6,
                value=(f'=IF(E{r}="Sí",1,IF(E{r}="Parcialmente",0.5,'
                       f'IF(OR(E{r}="No",E{r}="No sé"),0,"")))'))
        ws.cell(row=r, column=7,
                value=f'=IF(E{r}="","",IF(E{r}="Sí","",J{r}))')
        ws.cell(row=r, column=8, value=c["prioridad"])
        ws.cell(row=r, column=9, value=c["key"])
        ws.cell(row=r, column=10, value=c["que_implica"])

        for col in range(1, 11):
            cell = ws.cell(row=r, column=col)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = WRAP if col in (3, 4, 7) else Alignment(
                vertical="center", horizontal="center" if col in (1, 5, 6, 8) else "left")
            if r % 2 == 0:
                cell.fill = ALT_FILL

        ws.cell(row=r, column=5).fill = INPUT_FILL
        # "Qué implica" solo se muestra cuando la respuesta no es "Sí", pero
        # la fila debe poder contenerlo cuando aparezca.
        ws.row_dimensions[r].height = estimar_alto_fila([
            (c["pregunta"], anchos_auto[2]),
            (c["base_legal"], anchos_auto[3]),
            (c["que_implica"], anchos_auto[6]),
        ], alto_minimo=45)

    last_row = first_row + len(CONTROLES) - 1

    dv = DataValidation(
        type="list",
        formula1='"Sí,Parcialmente,No,No sé"',
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Respuesta no válida",
        error="Seleccione una opción de la lista: Sí, Parcialmente, No o No sé.",
    )
    dv.add(f"E{first_row}:E{last_row}")
    ws.add_data_validation(dv)

    ws.freeze_panes = "A2"
    configurar_impresion(ws)
    footer(ws, last_row + 2, 8)


def build_resultados(ws):
    set_col_widths(ws, [40, 20, 45])
    row = 1
    ws.cell(row=row, column=1, value="Resultados").font = TITLE_FONT
    row += 2

    a_first, a_last = 2, 2 + len(CONTROLES) - 1
    rng_resp = f"Autoevaluación!$E${a_first}:$E${a_last}"
    rng_score = f"Autoevaluación!$F${a_first}:$F${a_last}"
    rng_cat = f"Autoevaluación!$B${a_first}:$B${a_last}"
    rng_prio = f"Autoevaluación!$H${a_first}:$H${a_last}"

    ws.cell(row=row, column=1, value="Cumplimiento general").font = SUBTITLE_FONT
    row += 1
    ws.cell(row=row, column=1, value="Porcentaje de cumplimiento general").font = BODY_FONT
    overall_row = row
    c = ws.cell(row=row, column=2, value=f"=SUM({rng_score})/{len(CONTROLES)}")
    c.number_format = "0%"
    c.font = Font(name=FONT_NAME, size=12, bold=True)
    row += 2

    ws.cell(row=row, column=1, value="Cumplimiento por categoría").font = SUBTITLE_FONT
    row += 1
    header_row = row
    ws.cell(row=row, column=1, value="Categoría").font = HEADER_FONT
    ws.cell(row=row, column=2, value="% cumplimiento").font = HEADER_FONT
    for col in (1, 2):
        ws.cell(row=row, column=col).fill = HEADER_FILL
        ws.cell(row=row, column=col).border = BORDER
    row += 1
    cat_first_row = row
    for cat in CATEGORIAS:
        ws.cell(row=row, column=1, value=cat).font = BODY_FONT
        cell = ws.cell(
            row=row, column=2,
            value=f'=SUMIFS({rng_score},{rng_cat},A{row})/COUNTIFS({rng_cat},A{row})')
        cell.number_format = "0%"
        cell.font = BODY_FONT
        for col in (1, 2):
            ws.cell(row=row, column=col).border = BORDER
        row += 1
    cat_last_row = row - 1
    row += 1

    ws.cell(row=row, column=1, value="Brechas pendientes").font = SUBTITLE_FONT
    row += 1
    ws.cell(row=row, column=1, value="Controles de prioridad Alta aún no resueltos (respuesta distinta de \"Sí\")").font = BODY_FONT
    ws.cell(row=row, column=1).alignment = WRAP
    gaps_row = row
    ws.cell(row=row, column=2,
            value=f'=COUNTIFS({rng_prio},"Alta",{rng_resp},"<>Sí")').font = BODY_FONT
    row += 1
    ws.cell(row=row, column=1, value="Controles respondidos como \"Parcialmente\"").font = BODY_FONT
    partial_row = row
    ws.cell(row=row, column=2,
            value=f'=COUNTIF({rng_resp},"Parcialmente")').font = BODY_FONT
    row += 2

    ws.cell(row=row, column=1, value="Lectura del resultado").font = SUBTITLE_FONT
    row += 1
    lectura_cell = ws.cell(
        row=row, column=1,
        value=(f'=IF(B{overall_row}>={UMBRAL_ALTO},'
               f'"Nivel de cumplimiento alto: su empresa ya tiene resuelta la '
               f'mayoría de los controles evaluados. Enfóquese en cerrar las '
               f'brechas de prioridad Alta que aún queden pendientes en el '
               f'Plan de acción.",'
               f'IF(B{overall_row}>={UMBRAL_MEDIO},'
               f'"Nivel de cumplimiento intermedio: existen avances, pero '
               f'todavía hay controles importantes sin resolver. Revise el '
               f'Plan de acción y priorice los controles de prioridad Alta.",'
               f'"Nivel de cumplimiento bajo: la mayoría de los controles '
               f'evaluados no están resueltos. Se recomienda iniciar cuanto '
               f'antes por los controles de prioridad Alta del Plan de '
               f'acción, comenzando por el Aviso de Privacidad y los '
               f'lineamientos internos para atender solicitudes '
               f'ARCO-POL."))'))
    lectura_cell.font = BODY_FONT
    lectura_cell.alignment = WRAP
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.row_dimensions[row].height = 80
    row += 2

    # --- Formato condicional -------------------------------------------
    # Los porcentajes usan las mismas tres bandas en toda la hoja, para que
    # un color signifique siempre lo mismo.
    aplicar_cf_porcentaje(ws, f"B{overall_row}")
    aplicar_cf_porcentaje(ws, f"B{cat_first_row}:B{cat_last_row}")

    # Brechas de prioridad Alta: ninguna es buena señal; tres o más, mala.
    ws.conditional_formatting.add(f"B{gaps_row}", CellIsRule(
        operator="equal", formula=["0"], stopIfTrue=True, **CF_VERDE))
    ws.conditional_formatting.add(f"B{gaps_row}", CellIsRule(
        operator="lessThanOrEqual", formula=["2"], stopIfTrue=True,
        **CF_AMBAR))
    ws.conditional_formatting.add(f"B{gaps_row}", CellIsRule(
        operator="greaterThan", formula=["2"], stopIfTrue=True, **CF_ROJO))

    # "Parcialmente" no es un incumplimiento, pero sí trabajo a medio
    # terminar: se marca en ámbar solo cuando hay alguno.
    ws.conditional_formatting.add(f"B{partial_row}", CellIsRule(
        operator="greaterThan", formula=["0"], stopIfTrue=True, **CF_AMBAR))

    for r in (overall_row, gaps_row, partial_row):
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2).border = BORDER
    for r in range(cat_first_row, cat_last_row + 1):
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")

    configurar_impresion(ws)
    ws.page_setup.orientation = "portrait"
    footer(ws, row, 3)
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 18


def build_plan(ws):
    headers = ["N°", "Categoría", "Tema", "Prioridad", "Acción recomendada",
               "Responsable sugerido", "Estado"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    ws.row_dimensions[1].height = 30
    anchos_plan = [5, 14, 34, 10, 50, 34, 14, 16, 14]
    set_col_widths(ws, anchos_plan)
    ws.column_dimensions["H"].hidden = True
    ws.column_dimensions["I"].hidden = True
    ws.cell(row=1, column=8, value="Clave").font = HEADER_FONT
    ws.cell(row=1, column=9, value="Respuesta (lookup)").font = HEADER_FONT

    ordenados = sorted(
        CONTROLES, key=lambda c: 0 if c["prioridad"] == "Alta" else 1)

    a_first = 2
    a_last = 2 + len(CONTROLES) - 1

    first_row = 2
    for i, c in enumerate(ordenados):
        r = first_row + i
        # El "Tema" es la pregunta sin el signo de apertura. Si hay que
        # recortarla, se corta en el último espacio para no partir una
        # palabra por la mitad.
        tema = c["pregunta"].split("¿", 1)[-1]
        if len(tema) > 90:
            corte = tema[:87].rsplit(" ", 1)[0].rstrip(",;:")
            tema = corte + "…"
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=c["categoria"])
        ws.cell(row=r, column=3, value=tema)
        ws.cell(row=r, column=4, value=c["prioridad"])
        ws.cell(row=r, column=5, value=c["accion"])
        ws.cell(row=r, column=6, value=c["responsable"])
        ws.cell(row=r, column=8, value=c["key"])
        ws.cell(row=r, column=9,
                value=(f'=INDEX(Autoevaluación!$E${a_first}:$E${a_last},'
                       f'MATCH(H{r},Autoevaluación!$I${a_first}:$I${a_last},0))'))
        ws.cell(row=r, column=7,
                value=(f'=IFERROR(IF(I{r}="","Sin responder",'
                       f'IF(I{r}="Sí","Cumplido",'
                       f'IF(I{r}="Parcialmente","Parcial",'
                       f'IF(I{r}="No","Pendiente","Por confirmar")))),'
                       f'"Sin responder")'))

        for col in range(1, 10):
            cell = ws.cell(row=r, column=col)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = WRAP if col in (3, 5) else Alignment(
                vertical="center", horizontal="center" if col in (1, 4, 7) else "left")
            if r % 2 == 0:
                cell.fill = ALT_FILL
        if c["prioridad"] == "Alta":
            ws.cell(row=r, column=4).fill = ALTA_FILL
        else:
            ws.cell(row=r, column=4).fill = MEDIA_FILL
        ws.row_dimensions[r].height = estimar_alto_fila([
            (tema, anchos_plan[2]),
            (c["accion"], anchos_plan[4]),
            (c["responsable"], anchos_plan[5]),
        ], alto_minimo=40)

    last_row = first_row + len(CONTROLES) - 1

    # Formato condicional: la columna Estado se colorea según el valor que
    # devuelve su fórmula, de modo que cambia sola al responder la hoja
    # Autoevaluación.
    rango_estado = f"G{first_row}:G{last_row}"
    for texto, estilo in CF_ESTADOS:
        ws.conditional_formatting.add(rango_estado, CellIsRule(
            operator="equal", formula=[f'"{texto}"'], stopIfTrue=True,
            **estilo))

    # Resalta la celda de Prioridad solo cuando el control es de prioridad
    # Alta y además sigue sin resolverse: es lo que conviene atender primero.
    # Un control Alta ya cumplido conserva su relleno estático suave.
    ws.conditional_formatting.add(f"D{first_row}:D{last_row}", FormulaRule(
        formula=[f'AND($D{first_row}="Alta",$G{first_row}<>"Cumplido")'],
        stopIfTrue=True,
        fill=PatternFill("solid", fgColor="F8CBAD"),
        font=Font(name=FONT_NAME, size=10, bold=True, color="833C0C")))

    ws.freeze_panes = "A2"
    configurar_impresion(ws)
    footer(ws, last_row + 2, 7)


def main():
    wb = build()
    out_dir = os.path.join(os.path.dirname(__file__), "..", "dist")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "Autoevaluacion-LPDP.xlsx")
    wb.save(out_path)
    print(f"Generado: {out_path}")


if __name__ == "__main__":
    main()
