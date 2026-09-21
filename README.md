# Autoevaluación LPDP

Herramienta gratuita y de código abierto para que una empresa salvadoreña
estime, en menos de una hora, qué tan cerca está de cumplir con las medidas
de seguridad exigidas por la **Ley para la Protección de Datos Personales
(LPDP)** de El Salvador y por la normativa que la desarrolla, emitida por la
**Agencia de Ciberseguridad del Estado (ACE)**.

Está pensada para la persona que administra o dirige un negocio, **no** para
un especialista en seguridad de la información ni en derecho. El resultado
es un libro de Excel (`.xlsx`) que se descarga, se llena y se queda en la
computadora de quien lo usa: no envía ni recopila ninguna información.

## Descarga

El archivo listo para usar está en [`dist/Autoevaluacion-LPDP.xlsx`](dist/Autoevaluacion-LPDP.xlsx).

## Aviso: reforma de septiembre de 2026

Este libro está redactado conforme a las **Reformas a la Ley para la
Protección de Datos Personales** aprobadas por la Asamblea Legislativa el
**17 de septiembre de 2026** (57 votos a favor, 1 en contra), sobre
iniciativa del Presidente de la República.

Lo esencial para una empresa privada:

- **Ya no está obligada a nombrar un Delegado de Protección de Datos
  Personales.** La reforma derogó los artículos 15 y 17 de la LPDP. Esa
  obligación subsiste solo para las **instituciones públicas** (Art. 47
  reformado), donde además puede recaer en el Oficial de Información.
- **Las obligaciones no desaparecieron: se trasladaron a la empresa.** El
  Art. 16 reformado ahora lista obligaciones del *sujeto obligado*, entre
  ellas orientar a sus dependencias y proveedores (lit. b) y **establecer
  lineamientos internos escritos** para gestionar las solicitudes ARCO-POL
  (lit. d).
- **Las solicitudes ARCO-POL se presentan directamente ante la empresa**
  (Art. 18 reformado), con plazos que corren contra ella: 20 días hábiles
  para responder, prorrogables por otros 20 (Art. 20); 3 días para una
  denegatoria (Art. 22); 5 días para avisos de incompetencia, rectificación
  o revocación (Arts. 19, 21 y 30).
- **El Aviso de Privacidad debe indicar los medios para presentar esas
  solicitudes** (Art. 24 lit. f reformado).

> **Sobre la vigencia.** La reforma entra en vigencia **ocho días después de
> su publicación en el Diario Oficial**. Al momento de publicar esta versión
> no se había confirmado esa publicación, por lo que durante ese intervalo
> el texto anterior sigue siendo el vigente. El libro se adelanta a la
> reforma de forma deliberada, para que una empresa no invierta en nombrar
> un Delegado que ya no se le exigirá. Verifique la fecha de publicación
> antes de tomar decisiones basadas en este documento.

Los **Lineamientos para el Delegado** emitidos por la ACE desarrollaban los
artículos 15, 16 y 17. Como la reforma derogó dos de ellos y reescribió el
tercero, ese instrumento ya no se cita como base legal frente a una empresa
privada; sigue siendo pertinente para las instituciones públicas.

## Cómo usarlo

1. Abra `dist/Autoevaluacion-LPDP.xlsx` en Excel, LibreOffice Calc o Google
   Sheets.
2. Lea la hoja **Instrucciones** antes de responder.
3. En la hoja **Autoevaluación**, responda cada una de las 19 preguntas en
   la columna amarilla "Respuesta", eligiendo una opción del listado
   desplegable: `Sí`, `Parcialmente`, `No` o `No sé`.
4. Revise la hoja **Resultados** para ver su porcentaje general de
   cumplimiento, el porcentaje por categoría, y una lectura en lenguaje
   simple de lo que ese resultado significa.
5. Revise la hoja **Plan de acción** para ver, ordenadas por prioridad, las
   acciones recomendadas para cada brecha y quién dentro de la empresa
   debería encargarse de cada una.

Cada respuesta que no sea "Sí" muestra automáticamente, en la columna "Qué
implica", una explicación breve en lenguaje llano de la consecuencia
práctica de no tener ese control resuelto.

## Qué NO es esta herramienta

- **No es una auditoría.** Las respuestas son declarativas: la herramienta
  no verifica evidencia, documentos ni configuraciones reales.
- **No es una certificación.** Ningún resultado obtenido aquí constituye
  una certificación de cumplimiento de la LPDP.
- **No vincula a la ACE.** La Agencia de Ciberseguridad del Estado no
  participa en ni respalda esta herramienta.
- **No es asesoría legal.** Para decisiones con impacto legal, contable o
  económico, consulte con un profesional calificado o con su Delegado de
  Protección de Datos Personales.

## Cómo reportar una cita incorrecta

Cada referencia de la columna "Base legal" fue verificada contra el texto
de la Ley para la Protección de Datos Personales, las Reformas a dicha ley
aprobadas el 17 de septiembre de 2026 (fuente de los artículos marcados como
"texto reformado en 2026"), las Políticas de la ACE (Políticas N.º
001-0309025-DPDP) y la Normativa para el Desarrollo del Procedimiento
Administrativo Sancionador publicada en el Diario Oficial del 11 de agosto
de 2026.

Si encuentra una referencia que le parece incorrecta, desactualizada o mal
citada, repórtela abriendo un *issue* en este repositorio con:

- el número de pregunta o el nombre de la hoja donde la encontró,
- la cita que considera incorrecta, y
- la fuente que respalda la corrección propuesta (artículo, norma y fecha).

Si no tiene acceso a GitHub, puede escribir a **hola@nubitlan.com** con la
misma información.

## Cómo regenerar el libro de Excel

El archivo `.xlsx` es un **artefacto de compilación**: la fuente de verdad
es el script [`scripts/build_workbook.py`](scripts/build_workbook.py), que
usa [openpyxl](https://openpyxl.readthedocs.io/) para generar el libro con
fórmulas reales (no valores precalculados), de modo que la hoja se
recalcule automáticamente cuando usted cambia una respuesta.

```bash
python3 -m pip install openpyxl
python3 scripts/build_workbook.py
```

Después de cualquier cambio al script, verifique el resultado con
[`scripts/verify_workbook.py`](scripts/verify_workbook.py), que llena una
copia de prueba con una combinación conocida de respuestas, la recalcula
con LibreOffice en modo headless y compara cada celda calculada contra el
valor esperado (requiere `libreoffice` instalado):

```bash
python3 scripts/verify_workbook.py
```

## Licencia

Esta obra se publica bajo la licencia **Creative Commons
Atribución-CompartirIgual 4.0 Internacional (CC BY-SA 4.0)**. Puede
copiarla, modificarla, adaptarla y redistribuirla, incluso con fines
comerciales, siempre que:

1. dé crédito a **Nubitlan** como autor original, y
2. distribuya cualquier obra derivada bajo esta misma licencia.

Vea el texto completo en [`LICENSE`](LICENSE).

## Créditos

Publicado por **Nubitlan** como aporte a la comunidad empresarial
salvadoreña. No constituye asesoría legal ni certificación oficial de la
ACE.
