# CLAUDE.md

Este archivo da contexto a Claude Code (o a Jairo en 6 meses) al retomar este repositorio.

---

## Qué es

**Botiquión de Escritorio** — sistema de punto de venta para un negocio tipo botiquín/tienda
(venta libre, préstamos/fiado a clientes, abonos con saldo FIFO, promociones, inventario,
reportes fiscales). Corre como **Google Apps Script** (backend `Code.gs`) con una interfaz web
servida por Apps Script (`Index.html`) y datos en una Google Sheet vinculada al script.

**Stack:** Google Apps Script (JavaScript sobre V8 runtime de Google) + HTML/CSS/JS en un solo
archivo servido por `doGet()` + Google Sheets como base de datos.

## Para quién / contexto de negocio

Herramienta de uso interno para operar la caja de un negocio real — registra ventas, controla
crédito de clientes (fiado) con lógica de abonos FIFO (el abono más viejo se paga primero), y
genera reportes fiscales. *(Pendiente completar: nombre del negocio/cliente si aplica, y quién
opera la caja día a día — ver plantilla de contexto en `ESTANDARES_DESARROLLO.md` sección 7.)*

## Reglas duras — esto NUNCA se toca sin pensarlo dos veces

1. **`registrarVenta` y `registrarAbono` usan `LockService.getScriptLock()`** — dos cobros o dos
   abonos simultáneos (dos pestañas abiertas, doble click, o dos cajeros a la vez) no deben poder
   corromper el saldo del cliente ni duplicar una venta. No quitar el lock ni acortar su alcance
   sin entender por qué se agregó (v9, ver historial de commits).
2. **El saldo de crédito de un cliente se calcula FIFO** (`_aplicarAbonoFIFO`) — un abono nuevo
   siempre paga primero la deuda más antigua pendiente. Cambiar este orden altera retroactivamente
   cómo se interpretan todos los abonos históricos de un cliente.
3. **No reintroducir `marcarPagado`** — se eliminó deliberadamente por quedar como código muerto
   tras el nuevo flujo de abonos (ver commit `efd9437`). Si algo parece necesitarlo, es señal de
   que el flujo de pagos se está reconstruyendo mal, no de que haya que revivir la función vieja.

## Arquitectura

```
Code.gs (backend, ~2080 líneas)
├── Configuración / catálogo:  obtenerConfiguracion, guardarConfiguracion, getCatalogo,
│                               getCatalogoParaTarjetas, agregarProducto, guardarEdicion
├── Clientes / crédito:        getClientes, getPerfilCliente, guardarCliente, _leerCreditoCliente
├── Ventas y pagos (crítico):  registrarVenta (LockService), registrarAbono (LockService),
│                               _aplicarAbonoFIFO, _validarRepartoVenta, _metodoEtiquetaDePagos,
│                               _registrarPagosLedger, _pagosSheet
├── Inventario:                getInventario, getStockActual, procesarRestock,
│                               getStockAnomalo, ajustarInventario (+ _logAjusteInventario)
├── Promociones:                agregarPromo, actualizarPromo, eliminarPromo,
│                               obtenerPromoComponentes / guardarPromoComponentes
├── Reportes:                  getVentas, getFiados, getReportes, getReporteFiscal,
│                               getReporteAvanzado, getListaCompras
├── Mantenimiento:             reconstruirFiados, archivarMes, limpiarCache
└── doGet() / include()        sirve Index.html como la web app

Index.html (frontend, ~4660 líneas) — SPA de un solo archivo (HTML+CSS+JS inline),
consume las funciones de Code.gs vía google.script.run.

botiquin_v19_fixed2.xlsx — snapshot/plantilla local de la estructura de datos.
La fuente de datos REAL en producción es la Google Sheet vinculada al script
(no vive en este repo — Apps Script es "bound" a una Sheet específica en Drive).
```

**Por qué esta separación:** todo el estado vive en la Sheet (fuente de verdad única);
`Code.gs` son funciones sin estado propio que leen/escriben esa Sheet bajo demanda desde el
frontend. No hay base de datos externa ni servidor propio — el runtime lo da Google.

## Cómo ejecutar / desplegar

Este proyecto vive en el editor de Apps Script (script.google.com), vinculado a una Google
Sheet. Opciones para trabajar con él fuera del editor web:

```bash
# Si se usa clasp (Google Apps Script CLI) para sincronizar con este repo local:
npm install -g @google/clasp
clasp login
clasp pull    # trae el estado actual del editor de Apps Script a este repo
clasp push    # sube los cambios de este repo al editor de Apps Script
```

*(Pendiente confirmar: si ya se usa `clasp` en este proyecto o los cambios se pegan a mano en
el editor de Apps Script — si es a mano, cada cambio debe reflejarse aquí también para que este
repo no quede desincronizado de lo que realmente corre en producción.)*

No hay framework de tests ni linter configurado. Verificación mínima antes de un cambio:
revisar en el editor de Apps Script que no haya errores de sintaxis (Apps Script los marca al
guardar) y probar el flujo afectado manualmente en la web app.

## Historial

- **`7e14787`** — Backend: bloquear `registrarVenta`/`registrarAbono` con `LockService` (evita
  condiciones de carrera en cobros/abonos concurrentes).
- **`efd9437`** — Backend: eliminar `marcarPagado` (código muerto tras el nuevo flujo de abonos).
- **`2ec26aa`** — Frontend: checkout de pago mixto/parcial, venta libre/préstamo, abonos y recibo.
- **`b066d9d`** — Backend: pago mixto/parcial, abonos con saldo FIFO, venta libre/préstamo.
- **`cb41eff`** — Recuperación del proyecto (backend v7 + frontend v8.2) — commit inicial de este
  repo; el proyecto ya existía y se recuperó/versionó a partir de ese punto.

## Pendientes conocidos

- Repo recién respaldado a GitHub (11 ago 2026) — antes vivía solo en local, sin remoto. Sin
  `CLAUDE.md` hasta ahora pese a tener 5 commits reales de trabajo de producto.
- Confirmar mecanismo real de sincronización con el editor de Apps Script (`clasp` vs copiar/pegar
  a mano) y documentarlo arriba una vez confirmado.
- Sin tests automatizados — dado que `registrarVenta`/`registrarAbono` son las funciones más
  críticas (dinero real, lock de concurrencia), son las primeras candidatas si algún día se agrega
  algo de testing (ver sección 9.3 de `ESTANDARES_DESARROLLO.md`).

---

*CLAUDE.md creado: 11 Ago 2026 (el proyecto y sus commits son anteriores; solo faltaba este doc).*
