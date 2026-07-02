# 🎯 ESPECIFICACIÓN TÉCNICA CORREGIDA - MPP 2027 EN PYTHON

## ⚠️ PROBLEMAS IDENTIFICADOS Y SOLUCIONES

### ❌ PROBLEMA 1: Procesos no se desglosan por área
**Causa:** Las fórmulas COUNTIFS estaban incorrectas o no iteraban por todos los procesos

**Solución:** Estructura de datos tipo "catálogo maestro" que divide claramente:
- Procesos de Aclaración (1.1 - 1.11)
- Procesos de Aportaciones (2.1 - 2.4)
- Procesos de Recaudación (3.1 - 3.13)
- Procesos de Registro (4.1 - 4.7)
- Procesos de Retiros (5.1 - 5.16)

### ❌ PROBLEMA 2: Gantt no se genera correctamente
**Causa:** 
- Fórmula de fechas muy compleja
- Lógica de cascada no implementada
- Fórmula de visualización Gantt incorrecta

**Solución:** 
- Calcular fechas en Python (no en Excel)
- Generar fechas cascada directamente en datos
- Usar fórmula Gantt simple basada en comparación de fechas

---

## 📊 ARCHIVO EXCEL: SOLUCIÓN CORREGIDA

### ESTRUCTURA GENERAL

**Total de filas de datos:** 56 (5 header + 51 procesos + encabezados intermedios)
**Total de columnas:** 60 (14 datos + 46 semanas Gantt)

---

## 🔍 HOJA 1: DASHBOARD EJECUTIVO (Exactamente como debe ser)

### ZONA 1: Encabezado (A1:H3)
```
┌──────────────────────────────────────────────────────────────┐
│           PLAN ANUAL MPP 2027 - DASHBOARD EJECUTIVO          │
│         Período: Octubre 2026 - Octubre 2027 (52 semanas)    │
│                      Total: 51 Procesos                      │
└──────────────────────────────────────────────────────────────┘
```
- Merge: A1:H1, A2:H2, A3:H3
- Font: Calibri 14pt Bold, Blanco
- Fill: #1F4E78 (azul oscuro)
- Alignment: Centro, Centro
- Height: A1=15, A2=15, A3=15

### ZONA 2: KPIs Principales (A5:H8)

**Estructura exacta:**

```
Fila 5:
┌─────────────────────┬──────────────────────┬─────────────────────┐
│ Total Procesos      │ Completados          │ En Proceso          │
│ 51                  │ =COUNTIF(...         │ =COUNTIF(...        │
└─────────────────────┴──────────────────────┴─────────────────────┘
Col A-B              Col C-D                Col E-F               Col G-H

Fila 6:
┌─────────────────────┬──────────────────────┬─────────────────────┐
│ Pendientes          │ En Riesgo (Rojo)     │ En Alerta (Amarillo)│
│ =COUNTIF(...        │ =COUNTIF(...         │ =COUNTIF(...        │
└─────────────────────┴──────────────────────┴─────────────────────┘

Fila 7:
┌─────────────────────┬──────────────────────┬─────────────────────┐
│ Avance General      │ % de Completitud     │ Procesos Críticos   │
│ [BARRA PROGRESO]    │ 0%                   │ 0                   │
└─────────────────────┴──────────────────────┴─────────────────────┘
```

**Fórmulas exactas:**

- B5 (Total): 51
- D5 (Completados): `=COUNTIF(Gantt!G:G,"Completado")`
- F5 (En Proceso): `=COUNTIF(Gantt!G:G,"En Proceso")`

- B6 (Pendientes): `=COUNTIF(Gantt!G:G,"Pendiente")`
- D6 (En Riesgo): `=COUNTIF(Gantt!H:H,"Rojo")`
- F6 (En Alerta): `=COUNTIF(Gantt!H:H,"Amarillo")`

- B7 (Avance): `=D5/51`
- B7.NumberFormat: "0%"

### ZONA 3: Tabla Indicadores por Área (A10:G16)

**Encabezado (Fila 10):**
```
│ ÁREA                           │ TOTAL │ COMPLETADOS │ EN PROCESO │ PENDIENTES │ % AVANCE │ ESTADO  │
│ A                              │ B     │ C           │ D          │ E          │ F        │ G       │
```

**Fila 11-15: Datos por cada área**

```
ACLARACIÓN DE CUENTAS          │ 11    │ =COUNTIFS(Gantt!$C:$C,"Aclaración de Cuentas",Gantt!$G:$G,"Completado") │ =COUNTIFS(Gantt!$C:$C,"Aclaración de Cuentas",Gantt!$G:$G,"En Proceso") │ =11-C11-D11 │ =C11/11 │ [COLOR] │

APORTACIONES VOLUNTARIAS       │ 4     │ =COUNTIFS(Gantt!$C:$C,"Aportaciones Voluntarias",Gantt!$G:$G,"Completado") │ ... │ =4-C12-D12 │ =C12/4 │ [COLOR] │

RECAUDACIÓN                    │ 13    │ =COUNTIFS(Gantt!$C:$C,"Recaudación",Gantt!$G:$G,"Completado") │ ... │ =13-C13-D13 │ =C13/13 │ [COLOR] │

REGISTRO Y TRASPASO            │ 7     │ =COUNTIFS(Gantt!$C:$C,"Registro y Traspaso",Gantt!$G:$G,"Completado") │ ... │ =7-C14-D14 │ =C14/7 │ [COLOR] │

RETIROS                        │ 16    │ =COUNTIFS(Gantt!$C:$C,"Retiros",Gantt!$G:$G,"Completado") │ ... │ =16-C15-D15 │ =C15/16 │ [COLOR] │
```

**Formato Columna G (ESTADO):**
- Formato condicional: Escala de color
  - Si F > 0.8 → Verde (#00B050)
  - Si F > 0.5 → Amarillo (#FFC000)
  - Si F <= 0.5 → Rojo (#FF0000)
- Valor: Barra de progreso horizontal
- O simplemente: "●" (bullet) coloreado

### ZONA 4: Estado Detallado (A18:C23)

```
ESTADO DETALLADO - DETALLES OPERACIONALES

Procesos en Fase Cumplimiento:    B19 = =COUNTIFS(Gantt!$I:$I,"Cumpl.",Gantt!$G:$G,"En Proceso")
Procesos Pendientes de Negocio:   B20 = =COUNTIFS(Gantt!$I:$I,"Negocio",Gantt!$G:$G,"Pendiente")
Procesos con Retraso (Rojo):      B21 = =COUNTIF(Gantt!$H:$H,"Rojo")
Procesos en Alerta (Amarillo):    B22 = =COUNTIF(Gantt!$H:$H,"Amarillo")
Procesos On-Time (Verde):         B23 = =COUNTIF(Gantt!$H:$H,"Verde")
```

### ZONA 5: Gráficos (E10:H50)

**Gráfico 1: Column Chart - Avance por Área (E10:H23)**
- Título: "AVANCE POR ÁREA"
- Datos: A11:A15 (Áreas)
- Series: C11:C15 (Completados), D11:D15 (En Proceso), E11:E15 (Pendientes)
- Colores: Verde, Amarillo, Gris
- Legend: Abajo
- Height: 180px

**Gráfico 2: Line Chart - Tendencia Mensual (E25:H35)**
- Título: "TENDENCIA DE AVANCE (%)"
- X-Axis: Meses (Oct 2026 - Oct 2027)
- Y-Axis: % de avance (0% - 100%)
- Line Color: Azul (#0070C0)
- Height: 120px

**Gráfico 3: Pie/Doughnut - Estados (E37:H47)**
- Título: "DISTRIBUCIÓN DE ESTADOS"
- Data: B5, F5, B6 (Completados, En Proceso, Pendientes)
- Colors: Verde, Amarillo, Gris
- Height: 150px

### Ancho de columnas Dashboard
- A: 35
- B: 12
- C: 12
- D: 12
- E: 12
- F: 12
- G: 12
- H: 12

---

## 📈 HOJA 2: CRONOGRAMA GANTT (LA CRÍTICA)

### ESTRUCTURA BASE

**Total de columnas: 60**
- A-N: Información de proceso (14 columnas)
- O-AR: Gantt visual (44 semanas)

**Total de filas: 65**
- Fila 1: Título
- Fila 3: Encabezado principal
- Fila 4: Encabezado semanas
- Fila 5: Separador
- Filas 6-56: Datos de 51 procesos

---

### PARTE 1: Catálogo de Procesos (Columnas A-N)

#### Fila 1: Título (Merge A1:N1)
```
CRONOGRAMA GANTT - PLAN MPP 2027 (Octubre 2026 - Octubre 2027)
```
- Font: 14pt Bold Blanco
- Fill: #1F4E78
- Height: 25px

#### Fila 4: Encabezados (A4:N4)
```
┌────┬──────────────────────────┬───────────────────┬──────────┬───────┬───────┬──────────┬────────┬────────┬────┬────┬────┬────┬─────┐
│ ID │ PROCESO                  │ ÁREA              │ RESP.    │ INICIO│ FIN   │ ESTADO   │ RIESGO │ FASE   │ N° │ CO │ CU │ CO2│ HOL │
│ A  │ B                         │ C                 │ D        │ E     │ F     │ G        │ H      │ I      │ J  │ K  │ L  │ M  │ N   │
└────┴──────────────────────────┴───────────────────┴──────────┴───────┴───────┴──────────┴────────┴────────┴────┴────┴────┴────┴─────┘
```

**Formato encabezado:**
- Font: 10pt Bold Blanco
- Fill: #1F4E78
- Alignment: Centro, Wrap text
- Height: 30px
- Borders: Completo

#### Filas 6-56: DATOS DE LOS 51 PROCESOS

**ESTRUCTURA CRÍTICA - ESTO DEBE SER CORRECTO:**

```
Fila 6-16: ACLARACIÓN DE CUENTAS (11 procesos)
├─ 1.1│Modificación de datos de Trabajadores          │Aclaración│...│2026-10-01│...│Pendiente│Verde│Negocio│5│5│5│5│5
├─ 1.2│Conformación de expediente                      │Aclaración│...│2026-10-08│...│Pendiente│Verde│Negocio│5│5│5│5│5
├─ 1.3│Actualización de expediente                     │Aclaración│...│2026-10-15│...
├─ 1.4│Separación de Cuentas                           │Aclaración│...│2026-10-22│...
├─ 1.5│Unificación de Cuentas                          │Aclaración│...│2026-10-29│...
├─ 1.6│Generación / Recuperación NIP                   │Aclaración│...│2026-11-05│...
├─ 1.7│Mesa de aclaraciones presencial                 │Aclaración│...│2026-11-12│...
├─ 1.8│Mesa de aclaraciones remoto                     │Aclaración│...│2026-11-19│...
├─ 1.9│Homologación anual                              │Aclaración│...│2026-11-26│...
├─1.10│Noti.Afo                                         │Aclaración│...│2026-12-03│...
└─1.11│Posibles Duplicados                              │Aclaración│...│2026-12-10│...

Fila 17-20: APORTACIONES VOLUNTARIAS (4 procesos)
├─ 2.1│Aportaciones Voluntarias                        │Aportaciones│...│2026-10-01│...
├─ 2.2│Redes Comerciales                               │Aportaciones│...│2026-10-08│...
├─ 2.3│Transferencia de Recursos de Plan Privado       │Aportaciones│...│2026-10-15│...
└─ 2.4│Recepción Aportaciones Voluntarias              │Aportaciones│...│2026-10-22│...

Fila 21-33: RECAUDACIÓN (13 procesos)
├─ 3.1│Modelo de Notificación - Op.99 y Op.98          │Recaudación│...│2027-06-01│...
├─ 3.2│Cambio de Régimen Pensionario                   │Recaudación│...│2027-06-08│...
├─ 3.3│Reintegro de semanas cotizadas                  │Recaudación│...│2027-06-15│...
├─ 3.4│Solicitud a través de OSS                       │Recaudación│...│2027-06-22│...
├─ 3.5│Recaudación de recursos de dispersión           │Recaudación│...│2027-06-29│...
├─ 3.6│Aplicación cobro de comisión por saldo          │Recaudación│...│2027-07-06│...
├─ 3.7│Generación de Estado de Cuenta                  │Recaudación│...│2027-07-13│...
├─ 3.8│Anexo 71 - Operación 16                         │Recaudación│...│2027-07-20│...
├─ 3.9│Redención de bono ISSSTE                        │Recaudación│...│2027-07-27│...
├─3.10│Requisitos para la Entrega de Información       │Recaudación│...│2027-08-03│...
├─3.11│Devolución de Pagos sin Justificación Legal     │Recaudación│...│2027-08-10│...
├─3.12│Ajuste de bono                                  │Recaudación│...│2027-08-17│...
└─3.13│Recepción validación y acreditación de vivienda │Recaudación│...│2027-08-24│...

Fila 34-40: REGISTRO Y TRASPASO (7 procesos)
├─ 4.1│Apertura y registro de cuentas individuales      │Registro│...│2027-04-01│...
├─ 4.2│Apertura y Registro Móvil de cuentas            │Registro│...│2027-04-08│...
├─ 4.3│Emancipación menores de edad                    │Registro│...│2027-04-15│...
├─ 4.4│Apertura y Traspaso Móvil de cuentas            │Registro│...│2027-04-22│...
├─ 4.5│Traspaso de Cuentas Individuales                │Registro│...│2027-05-01│...
├─ 4.6│Posibles Duplicados                             │Registro│...│2027-05-08│...
└─ 4.7│Asignación y Reasignación de Cuentas            │Registro│...│2027-05-15│...

Fila 41-56: RETIROS (16 procesos)
├─ 5.1│Retiros Parciales Presencial / Profuturo Móvil  │Retiros│...│2027-01-01│...
├─ 5.2│Retiros Parciales AFORE Móvil / Web             │Retiros│...│2027-01-08│...
├─ 5.3│Retiros Parciales ISSSTE                        │Retiros│...│2027-01-15│...
├─ 5.4│Disposición de Recursos Totales                 │Retiros│...│2027-01-22│...
├─ 5.5│Retiros Programados y Pensión Garantizada       │Retiros│...│2027-01-29│...
├─ 5.6│Saldos Previos / Reinversión                    │Retiros│...│2027-02-05│...
├─ 5.7│Transferencias Extemporáneas                    │Retiros│...│2027-02-12│...
├─ 5.8│Transferencias ISSSTE                           │Retiros│...│2027-02-19│...
├─ 5.9│Reinversión Retiros No Cobrados                 │Retiros│...│2027-02-26│...
├─5.10│Retiro de Aportaciones Voluntarias              │Retiros│...│2027-03-05│...
├─5.11│Retiro de Aportaciones Voluntarias Profuturo    │Retiros│...│2027-03-12│...
├─5.12│Retiros de Aportaciones Voluntarias por AFORE   │Retiros│...│2027-03-19│...
├─5.13│Estado de Cuenta Pensionado                     │Retiros│...│2027-03-26│...
├─5.14│Reexpediciones                                  │Retiros│...│2027-04-02│...
├─5.15│Certificación de cuentas                        │Retiros│...│2027-04-09│...
└─5.16│Planes Privados de Pensión                      │Retiros│...│2027-04-16│...
```

---

### PARTE 2: Detalles de columnas (A-N)

#### Columna A: ID (readonly)
- Valores exactos: 1.1, 1.2, ..., 5.16
- Font: 9pt, Centro
- Width: 6

#### Columna B: Proceso (editable con restricciones)
- Nombres completos de procesos
- Font: 9pt, Izquierda, Wrap text
- Width: 40

#### Columna C: Área (readonly)
- Valores exactos:
  - "Aclaración de Cuentas"
  - "Aportaciones Voluntarias"
  - "Recaudación"
  - "Registro y Traspaso"
  - "Retiros"
- Font: 9pt, Centro
- Width: 18

#### Columna D: Responsable (EDITABLE)
- Vacío inicialmente
- Font: 9pt, Centro
- Width: 12

#### Columna E: Inicio (EDITABLE)
- Formato: YYYY-MM-DD
- Fechas iniciales según cronograma (Ver abajo)
- Font: 9pt, Centro
- Width: 12

#### Columna F: Fin (FÓRMULA - readonly)
```excel
=IF(ISNUMBER(E6),WORKDAY(E6,25),"")
```
- Calcula 25 días hábiles desde E6
- Formato: YYYY-MM-DD
- Font: 9pt, Centro
- Width: 12

#### Columna G: Estado (EDITABLE con lista desplegable)
- Valores permitidos: "Pendiente", "En Proceso", "Completado"
- Inicial: "Pendiente"
- Data Validation: List
- Font: 9pt, Centro, Bold
- Conditional Formatting:
  - "Pendiente" → Fondo Gris (#D0D0D0)
  - "En Proceso" → Fondo Amarillo (#FFC000)
  - "Completado" → Fondo Verde (#00B050)
- Width: 12

#### Columna H: Riesgo (FÓRMULA - readonly)
```excel
=IF(ISNUMBER(F6),
    IF(TODAY()>F6,"Rojo",
    IF(TODAY()>F6-3,"Amarillo","Verde")),
    "")
```
- Calcula automáticamente
- Font: 9pt, Centro, Bold
- Conditional Formatting:
  - "Rojo" → Fondo Rojo (#FF0000)
  - "Amarillo" → Fondo Amarillo (#FFC000)
  - "Verde" → Fondo Verde (#00B050)
- Width: 10

#### Columna I: Fase_Actual (FÓRMULA - readonly)
```excel
=IF(AND(ISNUMBER(E6),ISNUMBER(F6)),
    IF(G6="Completado","Completado",
    IF(TODAY()<WORKDAY(E6,5),"Negocio",
    IF(TODAY()<WORKDAY(E6,10),"CO",
    IF(TODAY()<WORKDAY(E6,15),"Cumpl.",
    IF(TODAY()<WORKDAY(E6,20),"CO-Obs.","Holgura"))))),
    "")
```
- Muestra en qué fase está cada proceso
- Font: 9pt, Centro
- Width: 12

#### Columnas J-N: Duración (readonly)
- J: Negocio = 5
- K: CO = 5
- L: Cumpl. = 5
- M: CO-Obs. = 5
- N: Holgura = 5
- Font: 9pt, Centro
- Width: 6 cada una

---

### PARTE 3: Gantt Visual (Columnas O-AR)

#### Encabezado de Semanas (Fila 4)

**Columnas O-AR = 44 semanas**

- O4 = "S1-Ene" (Semana 1 - Enero 2027)
- P4 = "S2-Ene"
- ...
- AR4 = "S44-Oct" (Semana 44 - Octubre 2027)

**Cálculo de semanas (en Python):**
```python
fecha_inicio = datetime(2026, 10, 1)
semanas = []
for i in range(44):
    fecha = fecha_inicio + timedelta(weeks=i)
    semana_label = f"S{i+1}-{fecha.strftime('%b')}"
    semanas.append(semana_label)
```

#### Fila 5: Separador visual
- Relleno: Gris (#E0E0E0)
- Height: 2px

#### Filas 6-56: Barras Gantt

**Fórmula para cada celda (O6:AR56):**

```excel
=IF(AND(ISNUMBER($E6),ISNUMBER($F6)),
    IF(AND(O$4>=$E6,O$4<=$F6),
        IF($G6="Completado",3,
        IF($G6="En Proceso",2,1)),
        0),
    0)
```

**Pero MEJOR en Python:**
- Calcular visualmente qué semana cae en cada rango de fechas
- Generar valor: 0, 1, 2, 3
- Excel aplica formato condicional

#### Formato Condicional (Rango O6:AR56)

**Escala de 3 colores:**
- Mínimo (0): Blanco (#FFFFFF)
- Medio (2): Amarillo (#FFC000)
- Máximo (3): Verde (#00B050)

**Alternativa con iconos:**
- 0 → Vacío
- 1 → Cuadrado gris
- 2 → Cuadrado amarillo
- 3 → Cuadrado verde

#### Ancho de columnas
- A: 6
- B: 40
- C: 18
- D: 12
- E: 12
- F: 12
- G: 12
- H: 10
- I: 12
- J-N: 6 cada una
- O-AR: 3 cada una (muy estrecho para ver como barra)

---

## 📅 CRONOGRAMA DE FECHAS DE INICIO

### REGLA DE CASCADA EN PYTHON

Cada proceso DEBE tener calculada su fecha de inicio de forma que:
1. Se ejecutan en cascada (paralelos en fases diferentes)
2. Optimizan el tiempo total
3. No se solapan indebidamente

### Distribución por Área

#### APORTACIONES VOLUNTARIAS (4 procesos)
- Oct 01, 2026: Proceso 2.1
- Oct 08, 2026: Proceso 2.2
- Oct 15, 2026: Proceso 2.3
- Oct 22, 2026: Proceso 2.4

#### ACLARACIÓN DE CUENTAS (11 procesos)
- Nov 02, 2026: Proceso 1.1
- Nov 09, 2026: Proceso 1.2
- Nov 16, 2026: Proceso 1.3
- Nov 23, 2026: Proceso 1.4
- Nov 30, 2026: Proceso 1.5
- Dic 07, 2026: Proceso 1.6
- Dic 14, 2026: Proceso 1.7
- Dic 21, 2026: Proceso 1.8
- Dic 28, 2026: Proceso 1.9
- Ene 04, 2027: Proceso 1.10
- Ene 11, 2027: Proceso 1.11

#### RETIROS (16 procesos)
- Ene 15, 2027: Proceso 5.1
- Ene 22, 2027: Proceso 5.2
- Ene 29, 2027: Proceso 5.3
- Feb 05, 2027: Proceso 5.4
- Feb 12, 2027: Proceso 5.5
- Feb 19, 2027: Proceso 5.6
- Feb 26, 2027: Proceso 5.7
- Mar 05, 2027: Proceso 5.8
- Mar 12, 2027: Proceso 5.9
- Mar 19, 2027: Proceso 5.10
- Mar 26, 2027: Proceso 5.11
- Abr 02, 2027: Proceso 5.12
- Abr 09, 2027: Proceso 5.13
- Abr 16, 2027: Proceso 5.14
- Abr 23, 2027: Proceso 5.15
- Abr 30, 2027: Proceso 5.16

#### REGISTRO Y TRASPASO (7 procesos)
- Abr 15, 2027: Proceso 4.1
- Abr 22, 2027: Proceso 4.2
- Abr 29, 2027: Proceso 4.3
- May 06, 2027: Proceso 4.4
- May 13, 2027: Proceso 4.5
- May 20, 2027: Proceso 4.6
- May 27, 2027: Proceso 4.7

#### RECAUDACIÓN (13 procesos)
- Jun 01, 2027: Proceso 3.1
- Jun 08, 2027: Proceso 3.2
- Jun 15, 2027: Proceso 3.3
- Jun 22, 2027: Proceso 3.4
- Jun 29, 2027: Proceso 3.5
- Jul 06, 2027: Proceso 3.6
- Jul 13, 2027: Proceso 3.7
- Jul 20, 2027: Proceso 3.8
- Jul 27, 2027: Proceso 3.9
- Ago 03, 2027: Proceso 3.10
- Ago 10, 2027: Proceso 3.11
- Ago 17, 2027: Proceso 3.12
- Ago 24, 2027: Proceso 3.13

---

## 🔧 CÓDIGO PYTHON - ESTRUCTURA REQUERIDA

### constants.py
```python
AREAS_PROCESOS = {
    'Aclaración de Cuentas': [
        '1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10', '1.11'
    ],
    'Aportaciones Voluntarias': ['2.1', '2.2', '2.3', '2.4'],
    'Recaudación': ['3.1', '3.2', '3.3', '3.4', '3.5', '3.6', '3.7', '3.8', '3.9', '3.10', '3.11', '3.12', '3.13'],
    'Registro y Traspaso': ['4.1', '4.2', '4.3', '4.4', '4.5', '4.6', '4.7'],
    'Retiros': ['5.1', '5.2', '5.3', '5.4', '5.5', '5.6', '5.7', '5.8', '5.9', '5.10', '5.11', '5.12', '5.13', '5.14', '5.15', '5.16']
}

FECHA_INICIO_PROCESOS = {
    '2.1': '2026-10-01', '2.2': '2026-10-08', '2.3': '2026-10-15', '2.4': '2026-10-22',
    '1.1': '2026-11-02', '1.2': '2026-11-09', ... (ver lista arriba),
    # ... todos los procesos con sus fechas
}
```

### excel_generator.py - Función clave
```python
def generar_gantt_con_datos_correctos(df_procesos):
    """
    CRÍTICO: Generar Gantt con:
    1. Todos los 51 procesos en orden correcto por área
    2. Fechas de inicio calculadas correctamente
    3. Fechas de fin usando WORKDAY
    4. Barras Gantt visibles
    """
    # Asegurar que df_procesos tiene:
    # - ID (exacto: 1.1, 1.2, etc)
    # - Proceso (nombre completo)
    # - Área (nombre área completo)
    # - Inicio (fecha calculada)
    # - Fin (WORKDAY formula)
    # - Estado (inicial: "Pendiente")
    # - Riesgo (fórmula)
    # - Fase_Actual (fórmula)
```

---

## ✅ VALIDACIONES CRÍTICAS

Antes de generar el Excel, verificar:

- [ ] Se cargan EXACTAMENTE 51 procesos
- [ ] Se dividen CORRECTAMENTE en 5 áreas
- [ ] Cada proceso tiene su ID único (1.1 - 5.16)
- [ ] Las fechas de inicio están esparcidas correctamente
- [ ] Las fechas de fin se calculan (25 días hábiles)
- [ ] El Gantt tiene 44+ semanas
- [ ] Las fórmulas COUNTIFS funcionan por área
- [ ] El Dashboard se actualiza al cambiar estados
- [ ] Los semáforos cambian de color correctamente
- [ ] No hay errores #REF! o #VALUE!

---

**Esta especificación CORREGIDA soluciona los problemas. Úsala en Copilot.**

