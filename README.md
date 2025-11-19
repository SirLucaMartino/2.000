# Planilla "Base Causas" – Estudio Jurídico Ley 20.009

Este documento describe, paso a paso, cómo debe configurarse el archivo Excel que usa el estudio para seguir causas de fraude bancario tramitadas ante los Juzgados de Policía Local (JPL). Sirve tanto como especificación funcional como prompt listo para pegar en herramientas tipo Codex, VBA, Apps Script o similares.

---

## 1. Contexto general

- La planilla registra causas de fraude bancario (Ley 20.009) que se litigan en distintos JPL.
- El objetivo es lograr trazabilidad, KPIs confiables y datos estandarizados.
- Se debe evitar la compra de pasajes para audiencias que no han sido notificadas.
- No se busca medir rendimiento individual (ganadas vs. perdidas), solo número de causas activas por tramitador.
- **Las primeras 13 columnas de la hoja `Base Causas` son intocables**: nombre, orden y contenido se mantienen.

Columnas base:

1. `RUT`
2. `Nombre Demandado`
3. `Cuantía`
4. `JPL`
5. `ROL`
6. `INGRESO ICA/CS`
7. `TRAMITADOR/A`
8. `GESTIONES`
9. `PENDIENTE`
10. `COMENTARIOS`
11. `LINK CARPETA DRIVE`
12. `ESTADO`
13. `PODER`

---

## 2. Columnas auxiliares (desde la columna 14 hacia la derecha)

### 🟨 Bloque Audiencias

| Columna | Tipo | Detalle |
| --- | --- | --- |
| `Fecha Audiencia` | Fecha | Se completa manualmente. |
| `Tipo Audiencia` | Lista | Valores: `CCCP`, `Primera absolución`, `Segunda absolución`, `Indagatoria`, `Otro`. |
| `Notificada (SI/NO)` | Lista | Valores: `SI`, `NO`. |
| `Requiere viaje (SI/NO)` | Lista | Valores: `SI`, `NO`. |
| `OK para comprar pasaje (SI/NO)` | Fórmula | `=SI(Y(Q2="SI";R2="SI");"SI";"NO")`. Sólo “SI” si está notificada y requiere viaje. |

### 🟧 Bloque Medida Prejudicial (MP)

| Columna | Tipo | Detalle |
| --- | --- | --- |
| `MP Interpuesta (SI/NO)` | Lista | Valores: `SI`, `NO`. |
| `Fecha MP` | Fecha | Día en que se presenta la MP. |
| `Resultado MP` | Lista | Valores: `Aprobada`, `Rechazada`, `Pendiente`. |

### 🟩 Bloque Tiempos y Estado Global

| Columna | Tipo | Fórmula / Lógica |
| --- | --- | --- |
| `Días desde ingreso` | Fórmula | `=SI(ESBLANCO(F2);"";HOY()-F2)` |
| `Es activa (SI/NO)` | Fórmula | Devuelve `"SI"` si `ESTADO` ∈ {`En tramitación`, `Con audiencia`, `Con MP`, `Notificación`, `Otro estado activo`}; en caso contrario `"NO"`. |
| `Audiencia próxima semana (SI/NO)` | Fórmula | `=SI(Y(NO(ESBLANCO(N2));N2>=HOY()+1;N2<=HOY()+7);"SI";"NO")` |
| `Días ingreso-audiencia` | Fórmula | `=SI(O(ESBLANCO(F2);ESBLANCO(N2));"";N2-F2)` |

---

## 3. Hoja `LISTAS` y validaciones de datos

1. Crear una hoja llamada **`LISTAS`**.
2. Definir listas verticales (un valor por fila) y rangos con nombre:
   - `TRAMITADORES`
   - `ESTADOS`
   - `PODER_LISTA`
   - `TIPOS_AUDIENCIA`
   - Listas básicas `SI/NO` para notificación, viajes y MP.
   - `RESULTADO_MP` con `Aprobada`, `Rechazada`, `Pendiente`.
3. Aplicar validación de datos “Lista” en:
   - `TRAMITADOR/A` → `=TRAMITADORES`
   - `ESTADO` → `=ESTADOS`
   - `PODER` → `=PODER_LISTA`
   - `Tipo Audiencia` → `=TIPOS_AUDIENCIA`
   - Columnas con respuesta `SI/NO` → lista fija `SI;NO`
   - `Resultado MP` → `Aprobada;Rechazada;Pendiente`

---

## 4. Automatizaciones y reportes

1. **Control de compra de pasajes**
   - La columna `OK para comprar pasaje` se calcula automáticamente.
   - Opcional: macro/botón que filtre filas donde `Requiere viaje="SI"` y `Notificada="NO"` o `OK="NO"`, mostrando alerta para evitar errores.

2. **Listado semanal de audiencias**
   - Crear hoja `Listado Semanal`.
   - Macro que copie desde `Base Causas` todas las filas con `Audiencia próxima semana="SI"` y pegue columnas clave (Fecha, Tipo, Demandado, ROL, JPL, Notificada, TRAMITADOR/A), ordenadas por fecha y JPL.

3. **KPIs en hoja `KPIs`**
   - Tabla dinámica 1: filas = `JPL`, valores = conteo de `Resultado MP` (Aprobada vs Rechazada) + % aprobación.
   - Tabla dinámica 2: filas = `TRAMITADOR/A`, filtro = `Es activa="SI"`, valores = conteo de `ROL` (o `RUT`).

4. **Control de antigüedad**
   - Utilizar `Días desde ingreso` para formato condicional (p.ej., rojo si supera umbral) y potenciales tablas dinámicas por JPL o tramitador.

5. **Formato condicional en `PENDIENTE` (opcional)**
   - Si hay audiencia fijada y `Notificada="NO"` → color amarillo.
   - Si hay audiencia fijada y `Notificada="SI"` → color naranjo.

---

## 5. Prompt listo para herramientas como Codex

```
Eres un asistente experto en Excel y VBA.
Trabajo con la hoja "Base Causas" del estudio jurídico (fraude bancario Ley 20.009). Las primeras 13 columnas son intocables: RUT, Nombre Demandado, Cuantía, JPL, ROL, INGRESO ICA/CS, TRAMITADOR/A, GESTIONES, PENDIENTE, COMENTARIOS, LINK CARPETA DRIVE, ESTADO, PODER.
Añadí columnas auxiliares para audiencias, medidas prejudiciales y control de tiempos, con las siguientes reglas:
- Fecha Audiencia (fecha manual)
- Tipo Audiencia (lista: CCCP, Primera absolución, Segunda absolución, Indagatoria, Otro)
- Notificada (SI/NO)
- Requiere viaje (SI/NO)
- OK para comprar pasaje = SI(Y(Notificada="SI"; Requiere viaje="SI"); "SI"; "NO")
- MP Interpuesta (SI/NO)
- Fecha MP
- Resultado MP (Aprobada, Rechazada, Pendiente)
- Días desde ingreso = SI(ESBLANCO(INGRESO);"";HOY()-INGRESO)
- Es activa = SI(ESTADO está en {En tramitación, Con audiencia, Con MP, Notificación, Otro estado activo}; "SI"; "NO")
- Audiencia próxima semana = SI(Fecha Audiencia está entre mañana y 7 días más; "SI"; "NO")
- Días ingreso-audiencia = SI(faltan fechas;""; Fecha Audiencia - INGRESO)
Necesito:
1. Crear hoja LISTAS con rangos para TRAMITADOR/A, ESTADO, PODER, Tipo Audiencia, respuestas SI/NO, y Resultado MP. Aplicar validación de datos usando esos rangos.
2. Generar hoja "Listado Semanal" que copie filas con Audiencia próxima semana = "SI" (mostrar Fecha, Tipo, Demandado, ROL, JPL, Notificada, TRAMITADOR/A) ordenadas por fecha y JPL.
3. Crear hoja "KPIs" con tablas dinámicas: (a) Resultado MP por JPL (aprobadas vs. rechazadas y %), (b) Causas activas por TRAMITADOR/A usando la columna Es activa.
4. Proponer formato condicional en PENDIENTE según estado de Notificada y Fecha Audiencia.
Devuélveme el VBA o las instrucciones detalladas para implementar todo sin alterar las 13 columnas iniciales.
```

Este README reemplaza la solicitud original y puede copiarse directamente en la herramienta elegida para automatizar la planilla.
