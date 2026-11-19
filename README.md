# Capacidad 2 — Control de costos de proyectos

Aplicación FastAPI + SQLAlchemy + PostgreSQL (Render) / SQLite (local), UI corporativa azul y exportación XLSX/PDF.

## Inicio local
```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload
```
Abrir http://127.0.0.1:8000 . API health: `/health`.

## Despliegue Render
1. Crear un repositorio GitHub y subir **el contenido de esta carpeta** (app.py, templates/, requirements.txt, render.yaml).
2. En Render: New > Blueprint > conectar repositorio > Apply.
3. Render creará servicio web y PostgreSQL con `DATABASE_URL` enlazada.
4. Abrir la URL asignada por Render. Si ya tienes una base PostgreSQL, crea solo el Web Service y configura `DATABASE_URL` en Environment.
5. Verificar `/health` y crear un proyecto.

## Orden de operación
1. Crear proyecto y paquetes WBS.
2. Crear línea base en borrador; registrar partidas por WBS y fecha; aprobar.
3. Registrar compromisos, gastos, solicitudes de cambio y avance físico.
4. Aprobar cambios cuando proceda; consultar dashboard y exportar Excel/PDF.

## Modelo de datos
`projects` → `work_packages`; `projects` → `baselines` → `budget_lines` (asociadas a WBS); `commitments` → opcionalmente `actuals` por `commitment_id`; `changes` y `progress` por WBS.

## Reglas de cálculo
- BAC: suma de partidas de la línea base aprobada vigente.
- PV: partidas con período <= fecha de corte.
- EV: BAC × último avance físico acumulado / 100.
- AC: suma de gastos reales devengados al corte.
- CV = EV − AC; SV = EV − PV; CPI = EV/AC; SPI = EV/PV.
- EAC = AC + (BAC − EV)/CPI (estimación de desempeño, si CPI > 0); VAC = BAC − EAC.
- Presupuesto ajustado = BAC + cambios aprobados.
- Compromisos y AC se muestran **separados**. La exposición `compromisos − AC` es solo orientativa, no representa saldo contractual exacto si hay gastos no vinculados a órdenes.

## Consideraciones antes de producción
Este es un **MVP funcional**, no un ERP de producción. Para uso empresarial agregar SSO/autenticación, RBAC y segregación de funciones para aprobaciones, auditoría inmutable, migraciones Alembic, pruebas automatizadas, validación de estados, multi-moneda con tipo de cambio, reversos y cancelaciones, vinculaciones completas de facturas a contratos y costos, cierre de períodos, y reglas de re-baselining.

**Nota EVM:** la versión actual calcula EVM contra la línea base vigente; no incorpora cambios aprobados al BAC/EV hasta crear y aprobar una nueva versión de línea base. Esto conserva trazabilidad pero requiere gobernar la replanificación. La curva PV depende de la fecha de cada partida presupuestal.

**Diseño visual:** se proporciona una interfaz corporativa aproximada; para replicar exactamente el look & feel del Directorio de Clientes se necesita reutilizar sus CSS, componentes y assets originales.
