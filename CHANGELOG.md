# Changelog

Todas las versiones notables de este proyecto se documentan en este archivo.

El formato sigue aproximadamente [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/).

## [1.1.0] - 2026-09-20

Actualización por las **Reformas a la Ley para la Protección de Datos
Personales** aprobadas por la Asamblea Legislativa el 17 de septiembre de
2026 (57 votos a favor, 1 en contra), a iniciativa del Presidente de la
República. La reforma entra en vigencia ocho días después de su publicación
en el Diario Oficial; esta versión se adelanta a ese plazo de forma
deliberada, para que ninguna empresa invierta en nombrar un Delegado que ya
no se le exigirá.

### Cambiado

- `org.delegado` ya no pregunta por el nombramiento formal de un Delegado.
  Ahora pregunta si hay una persona o área designada internamente para
  atender las solicitudes ARCO-POL y orientar a dependencias y proveedores,
  con base en el Art. 16 lit. b) reformado. El texto aclara expresamente
  que una empresa privada ya no está obligada a nombrar Delegado.
- `der.aviso` se reancla al Art. 24 de la LPDP, en especial su lit. f)
  reformado, que exige indicar en el Aviso de Privacidad los medios para
  presentar solicitudes ARCO-POL. Antes citaba los Lineamientos, Art. 31.
- `der.formularios` deja de exigir formularios físicos en cada sede
  (obligación que venía de los Lineamientos, Art. 32) y ahora pregunta por
  la existencia de un medio claro para presentar solicitudes, con base en
  los Arts. 18 y 24 lit. f) reformados.
- `der.plazos` se reancla a los Arts. 18, 19, 20, 21, 22 y 30 reformados, y
  su texto incorpora los plazos concretos, que tras la reforma corren
  contra la empresa y no contra un Delegado.
- `org.capacitacion` pierde la cita a los Lineamientos, Art. 22, y conserva
  únicamente las Políticas ACE, Art. 4 (Medidas Organizativas, lit. c).
- El responsable sugerido "Delegado o asesor legal" pasa a llamarse
  "Responsable interno o asesor legal", por coherencia con la reforma.
- La hoja Instrucciones y el `README.md` incorporan una sección sobre la
  reforma, su alcance y su fecha de entrada en vigencia.

### Agregado

- Control `org.lineamientos_arcopol` (prioridad Alta), que reemplaza a
  `org.delegado_ace` y cubre la obligación **nueva y expresa** del Art. 16
  lit. d) reformado: establecer lineamientos internos escritos para
  gestionar las solicitudes ARCO-POL.
- `referencias/README.md` documenta el proyecto de reforma, con el detalle
  artículo por artículo de su efecto sobre la LPDP.

### Eliminado

- Control `org.delegado_ace` (comunicación del nombramiento a la ACE e
  inscripción en el Registro Público de Delegados). La reforma dejó sin
  objeto ese trámite para las empresas privadas.
- Toda cita a los **Lineamientos para el Delegado de Protección de Datos
  Personales** como base legal. Ese instrumento desarrollaba los artículos
  15, 16 y 17 de la LPDP; derogados el 15 y el 17 y reescrito el 16, deja
  de ser base citable frente a una empresa privada, aunque sigue siendo
  pertinente para las instituciones públicas.

### Pendiente

- Sustituir la referencia al proyecto de reforma por el Diario Oficial en
  que se publique, e incorporar a las citas el número y la fecha del
  decreto, que en el proyecto aparecen en blanco.

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
