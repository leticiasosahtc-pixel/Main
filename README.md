# Agente de Reembolsos

Notebook: `Ejercicio_Agente_Ollama_Docker.ipynb`

Lee el listado de solicitudes de reintegro y aplica la tabla de decisión:

| `estado` en la base | Acción |
|---|---|
| `cobro_duplicado` | Devolución directa (previa validación de monto en el sandbox) |
| `pago_correcto` | Informa que no corresponde la devolución |
| cualquier otro (incl. `no_encontrada`) | Deriva a supervisión humana; el grafo se pausa |

## Regla de diseño

El ruteo lo decide el campo `estado` de la base, de forma determinística.
El LLM redacta la justificación que queda en el log de auditoría — nunca decide el camino.
