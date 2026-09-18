# Agente de Reintegros

Notebook: `Ejercicio_Agente_Ollama_Docker.ipynb` — LangGraph + Ollama (`phi3:mini`) + sandbox Docker.

## Flujo

El usuario solicita un reintegro → el agente busca en la base de facturación si corresponde →
ejecuta las devoluciones que puede ejecutar solo y reserva al humano las que no.

## Dos fuentes de datos

| Dataset | Archivo | Filas | Confiabilidad |
|---|---|---|---|
| Solicitudes | `datos/solicitudes.csv` | 45 | No confiable: es el reclamo del cliente |
| Facturación | `datos/facturas.csv` | 40 | Fuente de verdad |

Subí ambos a `/content/` en Colab, o dejalos junto al notebook. El loader busca en
`/content/`, `./datos/` y el directorio actual.

Cuando el pedido y el registro no coinciden, gana el registro.

## Niveles de autoridad

| Nivel | Caso | Quién actúa |
|---|---|---|
| 1 — ejecuta solo | `cobro_duplicado`, monto bajo el límite, solicitante = titular | El agente |
| 2 — responde solo | `pago_correcto` | El agente informa; no mueve dinero |
| 3 — reservado | Estado no automatizable, monto sobre el límite o inválido, factura inexistente, titular que no coincide, factura ya reintegrada, o desacuerdo entre el LLM y la política | Una persona (el grafo se pausa) |

## Dónde decide el LLM

El LLM decide en lo ambiguo: interpreta el pedido en lenguaje natural, propone la acción y
redacta la respuesta. No tiene la última palabra sobre mover dinero: `verificar_decision`
contrasta su propuesta contra la política y la titularidad.

**El LLM puede frenar una devolución, pero no puede autorizar una que el registro no respalda.**

## Verificación

El Paso 13 separa dos cosas:

- **Seguridad** (invariante duro, con `assert`): nunca se ejecuta una devolución fuera de
  política. Se cumple incluso con el LLM devolviendo prosa sin etiqueta, o proponiendo
  `devolver` en todos los casos.
- **Automatización** (métrica blanda): cuántos casos resolvió solo. Escalar de más es un costo,
  no un defecto de seguridad.

El Paso 14 es la contraprueba: con `VERIFICAR_CONTRA_POLITICA = False`, un LLM que propone
`devolver` siempre logra que se pague un `pago_correcto`.

## Cobertura de los datasets

`solicitudes.csv` incluye los casos que rompen un agente ingenuo:

| Solicitudes | Qué prueban |
|---|---|
| S-011 a S-013 | Número de factura mal escrito (`f-2024-911`, `2024-912`, `F 2024 917`) |
| S-014 a S-016 | Duplicados sobre el límite de aprobación automática |
| S-017 | Factura con monto $0,00 |
| S-018 a S-022 | El cliente afirma cobro duplicado, el registro dice `pago_correcto` |
| S-035 a S-037 | Quien pide no es el titular |
| S-038, S-039 | Facturas inexistentes |
| S-040 a S-042 | Pedidos sin número de factura |
| S-043 | El texto menciona dos facturas |
| S-044, S-045 | Reclamos repetidos sobre facturas ya reintegradas |

## Resultados de la verificación

Medido con un LLM simulado en tres comportamientos (el número real de `phi3:mini` sale
al correr en Colab):

| Comportamiento del LLM | Seguridad | Resueltos como corresponde |
|---|---|---|
| Coopera | ✅ 0 devoluciones indebidas | 45/45 |
| Devuelve prosa sin etiqueta | ✅ 0 devoluciones indebidas | 45/45 |
| Propone `devolver` siempre | ✅ 0 devoluciones indebidas | 40/45 (escala de más) |
| Ídem, sin `verificar_decision` | ❌ paga un `pago_correcto` | — |
