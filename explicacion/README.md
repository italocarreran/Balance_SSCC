# explicacion/

Páginas que explican paso a paso el costo de oportunidad BESS, con un ejemplo
calculado por el código real. Se abren con doble clic (no necesitan servidor).

| Página | Qué explica |
|---|---|
| `re545.html` | Calculo RE545 (ventana con oferta completa, ResEx CNE 545 §2.2.2.2) |
| `ecostos.html` | Calculo E Costos (ventana sin oferta completa, Estudio de Costos §10.2.5) |

## Quién toca qué

El contenido y el diseño están separados para que dos personas (o dos IA)
trabajen sin pisarse.

| Archivo | Dueño | Qué contiene |
|---|---|---|
| `contenido/*.js` | **Contenido** (Claude) | Textos, pasos, cuentas, avisos. Nada de estilo. |
| `estilos.css` | **Diseño** (ChatGPT) | Colores, tipografía, espaciado, claro/oscuro. |
| `pagina.js` | **Diseño** | Cómo se arma la página a partir del contenido. |
| `figuras.js` | **Diseño** | Gráficos y tablas. Solo dibujan datos del ejemplo. |
| `re545.html`, `ecostos.html` | **Diseño** | Esqueleto: carga los archivos en orden. |
| `datos/ejemplo.js` | Generado | No editar a mano. Sale de `generar_ejemplo.py`. |
| `generar_ejemplo.py` | Contenido | Corre las funciones reales de `Script/nucleo` sobre el caso de ejemplo. |

Regla: si un cambio necesita tocar el archivo del otro, se pide, no se hace.

## Contrato entre contenido y diseño

`contenido/X.js` define `window.CONTENIDO`:

```js
{
  hoja: "re545" | "ecostos",        // elige el color de acento
  eyebrow, titulo, bajada,          // encabezado
  caso: [{valor, detalle}],         // datos del ejemplo arriba
  nota_inicial,                     // opcional
  pasos: [{
    id, titulo, texto,              // texto admite <b>, <i>, <code>
    partes: ["..."],                // lista numerada (opcional)
    figura: "nombre",               // window.FIGURAS[nombre] (opcional)
    calculo: "...",                 // bloque monoespaciado; \n = salto
    resultado: {etiqueta, valor, detalle},
    bloques: [{tipo, titulo, texto, opciones, casos}],
    siguiente: true                 // paso que queda pendiente
  }],
  pie: ["..."]
}
```

- `bloques[].tipo`: `nota`, `aviso`, `revision` (algo en duda), `pausa` (decisión pendiente).
  `opciones: [{nombre, texto, preferida}]`, `casos: [{titulo, texto, cuenta, malo}]`.
- En cualquier texto, `{nombre}` se reemplaza por `EJEMPLO.valores.nombre`
  con formato chileno; `{nombre:0}` sin decimales. Los nombres disponibles
  están en `generar_ejemplo.py` (diccionario `v`).
- Figuras disponibles: E Costos `filtro`, `bloques`, `ordenar`, `puestos`,
  `precios_descuento`, `dual`, `wf_c1`, `wf_c2`; RE545 `ed`, `orden_horas`,
  `curva`, `inyeccion`, `cajas`.

Quien diseña puede cambiar libremente cómo se ve todo eso (orden visual,
colores, figuras), pero no los textos ni los nombres del contrato. Si una
figura nueva necesita datos que no están en `EJEMPLO`, se pide que se agreguen
en `generar_ejemplo.py`.

## Reglas del contenido

- Nunca citar letras de columnas de Excel. Se nombra el concepto.
- Todo número del ejemplo va como `{placeholder}`, no escrito a mano.
- Si algo está en revisión o en pausa, se marca con su bloque, no se esconde.

## Actualizar el ejemplo

```
python explicacion/generar_ejemplo.py
```

Si cambia el cálculo en `Script/nucleo`, se vuelve a correr y las páginas
muestran los números nuevos.
