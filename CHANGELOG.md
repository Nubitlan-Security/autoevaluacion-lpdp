# Changelog

Todas las versiones notables de este proyecto se documentan en este archivo.

El formato sigue aproximadamente [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/).

## [1.0.0] - 2026-09-05

### Agregado

- Primera versión pública de la Autoevaluación LPDP: libro de Excel con 19
  controles agrupados en 4 categorías (Organización, Derechos, Técnicas,
  Entorno y terceros).
- Hoja **Instrucciones**: qué es la herramienta, cómo responder, criterio de
  calificación, advertencia de que es una autoevaluación declarativa y no
  evidenciada, aviso legal y licencia.
- Hoja **Autoevaluación**: las 19 preguntas con su base legal verificada,
  respuesta mediante lista desplegable (Sí / Parcialmente / No / No sé),
  puntaje calculado por fórmula y explicación de "qué implica" cada
  respuesta distinta de "Sí".
- Hoja **Resultados**: porcentaje de cumplimiento general y por categoría,
  conteo de brechas de prioridad Alta sin resolver, conteo de respuestas
  "Parcialmente" y una lectura en lenguaje simple del resultado.
- Hoja **Plan de acción**: las 19 acciones recomendadas, ordenadas por
  prioridad, con responsable sugerido y estado calculado en vivo a partir
  de las respuestas.
- Script `scripts/build_workbook.py` (fuente de verdad del libro, usando
  openpyxl) y `scripts/verify_workbook.py` (verificación automática de las
  fórmulas contra una combinación de respuestas conocida, usando
  LibreOffice en modo headless para recalcular).
- Citas de "Base legal" verificadas contra la Ley para la Protección de
  Datos Personales, las Políticas de la Agencia de Ciberseguridad del
  Estado (Políticas N.º 001-0309025-DPDP) y la normativa publicada en el
  Diario Oficial del 11 de agosto de 2026 (Lineamientos para el Delegado de
  Protección de Datos Personales y Normativa para el Desarrollo del
  Procedimiento Administrativo Sancionador).
- `README.md`, `LICENSE` (CC BY-SA 4.0) y este `CHANGELOG.md`.
