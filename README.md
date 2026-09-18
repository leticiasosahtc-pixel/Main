# Agente de Reintegros

Notebook: `Ejercicio_Agente_Ollama_Docker.ipynb` — LangGraph + Ollama (`phi3:mini`) + sandbox Docker.

## Flujo

El usuario solicita un reintegro → el agente busca en la base de facturación si corresponde →
ejecuta las devoluciones que puede ejecutar solo y reserva al humano las que no.

## Dos fuentes de datos

| Dataset | Qué es | Confiabilidad |
|---|---|---|
| `solicitudes.csv` | Lo que **pide** el usuario, en lenguaje natural | No confiable: es el reclamo del cliente |
| `facturas.csv` | Lo que dice el **sistema de pagos** | Fuente de verdad |

Cuando el pedido y el registro no coinciden, gana el registro.

## Niveles de autoridad

| Nivel | Caso | Quién actúa |
|---|---|---|
| 1 — ejecuta solo | `cobro_duplicado`, monto bajo el límite, solicitante = titular | El agente |
| 2 — responde solo | `pago_correcto` | El agente informa; no mueve dinero |
| 3 — reservado | Estado no automatizable, monto sobre el límite, factura inexistente, titular que no coincide, o desacuerdo entre el LLM y la política | Una persona (el grafo se pausa) |

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
