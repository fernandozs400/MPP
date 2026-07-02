# ESTRUCTURA TÉCNICA - PLAN MPP 2027 EN EXCEL

## 📋 HOJA 1: DASHBOARD EJECUTIVO

### Sección Superior - KPIs Principales
```
┌─────────────────────────────────────────────────────────┐
│         PLAN ANUAL MPP 2027 - DASHBOARD EJECUTIVO       │
│                                                          │
│  Total Procesos: 51  │  Completados: [FÓRMULA]  [XX%]  │
│  En Proceso: [XX]    │  Pendientes: [XX]        [XX%]  │
│  En Riesgo: [XX]     │  Avance General: [████] [XX%]    │
└─────────────────────────────────────────────────────────┘
```

### Indicadores por Área (Tabla Dinámica)
| Área | Total | Completados | En Proceso | Pendientes | % Avance | Semáforo |
|------|-------|------------|-----------|-----------|----------|----------|
| Aclaración de Cuentas | 11 | [↑] | [↑] | [↑] | [FÓRMULA] | [COLOR] |
| Aportaciones Voluntarias | 4 | [↑] | [↑] | [↑] | [FÓRMULA] | [COLOR] |
| Recaudación | 13 | [↑] | [↑] | [↑] | [FÓRMULA] | [COLOR] |
| Registro y Traspaso | 7 | [↑] | [↑] | [↑] | [FÓRMULA] | [COLOR] |
| Retiros | 16 | [↑] | [↑] | [↑] | [FÓRMULA] | [COLOR] |

### Estado Detallado
- Procesos en revisión con Cumplimiento: [XX]
- Procesos pendientes por Negocio: [XX]
- Procesos con retraso (Rojo): [XX]
- Holgura utilizada (Amarillo): [XX]

### Gráficos
1. **Gráfico de Barras:** Avance por Área (Completados/En Proceso/Pendientes)
2. **Gráfico de Líneas:** Tendencia mensual de avance (Oct 2026 - Oct 2027)
3. **Gráfico de Dona:** Distribución de estados (Completado | En Proceso | Pendiente)
4. **Barras de Progreso:** Una por área

---

## 📊 HOJA 2: CRONOGRAMA GANTT + CATÁLOGO DE PROCESOS

### Sección A: Catálogo Maestro (Lado Izquierdo)
```
ID | Proceso | Área | Responsable | Inicio | Fin | Estado | Riesgo
---|---------|------|-------------|--------|-----|--------|-------
1.1| Modificación de Datos | AC | [Nombre] | [Fecha] | [Fecha] | [Estado] | [Semáforo]
1.2| Conformación Expediente | AC | [Nombre] | [Fecha] | [Fecha] | [Estado] | [Semáforo]
...
```

### Sección B: Gantt Visual (Lado Derecho)
- Cronograma horizontal tipo Microsoft Project
- Barras de color por fase (Negocio → CO → Cumplimiento → CO Observaciones → Holgura)
- Cada fila representa un proceso
- Columnas representan semanas (Oct 2026 - Oct 2027)
- Formato condicional: Verde (Completado) → Amarillo (En Proceso) → Rojo (Retraso)
- Líneas de dependencia entre fases

### Fórmulas Automáticas
- **Cálculo de fechas:** Suma automática de días hábiles por fase
- **Estado:** Comparación de fecha actual vs fecha de fin planificada
- **Semáforos:** Verde si on-time, Amarillo si ±3 días, Rojo si >3 días

### Distribución Temporal (Automatizada)
- **Oct 2026:** Aportaciones Voluntarias (4 procesos, en cascada)
- **Nov-Dic 2026:** Aclaración de Cuentas (11 procesos, 2 semanas holgura)
- **Ene-Mar 2027:** Retiros (16 procesos, 3 meses)
- **Abr-May 2027:** Registro y Traspaso (7 procesos)
- **Jun-Ago 2027:** Recaudación (13 procesos)
- **Sep 2027:** Regularización y rezagados
- **Oct 2027:** Liberación anual

---

## 🔧 FÓRMULAS CLAVE

### En Dashboard
```excel
Completados = COUNTIF(Gantt!Estado:Estado,"Completado")
En_Proceso = COUNTIF(Gantt!Estado:Estado,"En Proceso")
Pendientes = COUNTIF(Gantt!Estado:Estado,"Pendiente")
% = Completados / 51

Estado_Cumplimiento = COUNTIFS(Gantt!Fase:Fase,"Cumplimiento",Gantt!Estado:Estado,"En Proceso")
Pendientes_Negocio = COUNTIFS(Gantt!Fase:Fase,"Negocio",Gantt!Estado:Estado,"Pendiente")
```

### En Gantt
```excel
Fecha_Fin = Inicio + (Días_Negocio + Días_CO + Días_Cumplimiento + Días_CO2 + Holgura) días hábiles
Estado = IF(TODAY() > Fecha_Fin, "Retraso", IF(TODAY() > Fecha_Fin - 3, "Amarillo", "On-Time"))
Semáforo = IF(Estado="Retraso", "🔴", IF(Estado="Amarillo", "🟡", "🟢"))
```

---

## 📌 CARACTERÍSTICAS AUTOMÁTICAS

✅ **Actualización dinámica:** Los KPIs del Dashboard se actualizan al cambiar estados en Gantt
✅ **Formato condicional:** Colores automáticos por semáforo
✅ **Cascada:** Procesos en paralelo optimizando tiempos
✅ **Dependencias visuales:** Líneas conectando fases
✅ **Hoja protegida:** Para evitar borrados accidentales de fórmulas
✅ **Tabla interactiva:** Filtros en catálogo de procesos

---

## 📱 POWER BI - ARCHIVO .PBIX

### Páginas Propuestas
1. **Portada / Resumen Ejecutivo**
2. **Avance General (KPIs + Tendencias)**
3. **Avance por Área (Desglosado)**
4. **Timeline (Gantt interactivo)**
5. **Indicadores de Riesgo**
6. **Tabla Detallada de Procesos**

### Visualizaciones
- Card visual para KPIs
- Clustered bar chart (Áreas)
- Line chart (Tendencia mensual)
- Matrix (Procesos detallados)
- Slicer (Filtros por Área, Estado, Mes)
- Gantt personalizado (por proceso y fecha)
- Heatmap (Riesgo vs Tiempo)

---

## 📥 ENTREGABLES

1. **MPP_2027_Dashboard_Gantt.xlsx** → Archivo Excel con 2 hojas
2. **MPP_2027_Dashboard.pbix** → Archivo Power BI
3. **README.md** → Instrucciones de uso y actualización
4. **Catálogo_Procesos.csv** → Base de datos con los 51 procesos (para importar en Power BI)
