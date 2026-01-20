# Sistema de Control de Inventario - ConcentradoHG

## 📋 Descripción General

Sistema completo de control de inventario para Excel con VBA que gestiona entradas, catálogo, inventario y concentrado de materiales.

## 🗂️ Estructura de Archivos Creados

### Módulos Principales
1. **Modulo1_Nuevo.vba** - Funciones auxiliares y navegación
2. **Modulo2_Nuevo.vba** - Gestión de inventario

### Hojas de Trabajo
3. **HojaEntradas_Nuevo.vba** - Captura y validación de entradas
4. **HojaCatalogo_Nuevo.vba** - Catálogo de productos
5. **HojaInventario_Nuevo.vba** - Control de existencias
6. **HojaConcentrado_Nuevo.vba** - Registro de consumos

### Configuración
7. **ThisWorkbook_Nuevo.vba** - Eventos del libro y configuración
8. **REFERENCIA_COLUMNAS.md** - Tabla de referencia completa

---

## 🚀 Instrucciones de Implementación

### Paso 1: Crear el Archivo Excel
1. Crea un nuevo archivo Excel (.xlsm - con macros habilitadas)
2. Nómbralo como desees (ejemplo: `ConcentradoHG_Nuevo.xlsm`)

### Paso 2: Crear las Hojas
Crea las siguientes hojas en Excel:
- **Entradas** (con encabezados en fila 6: Referencia B, Lote D, Remisión G)
- **Catalogo** (con encabezados: Descripción B, ID/Referencia C, Lote D, Cant. Lotes E)
- **Inventario** (con encabezados: Referencia B, Lote C, Cantidad D)
- **Concentrado** (con encabezados desde columna Q según tabla de referencia)

### Paso 3: Insertar el Código VBA

#### A) Abrir el Editor VBA
1. Presiona `Alt + F11` para abrir el Editor VBA
2. En el panel izquierdo verás el proyecto de tu libro

#### B) Insertar Módulos
1. Click derecho en tu proyecto → Insertar → Módulo
2. Copia el contenido de **Modulo1_Nuevo.vba** y pégalo
3. Repite para crear otro módulo con **Modulo2_Nuevo.vba**

#### C) Código de las Hojas
Para cada hoja (Entradas, Catalogo, Inventario, Concentrado):
1. En el panel izquierdo, haz doble clic en la hoja correspondiente (ejemplo: Hoja1 (Entradas))
2. Copia el código del archivo correspondiente:
   - **Entradas** → HojaEntradas_Nuevo.vba
   - **Catalogo** → HojaCatalogo_Nuevo.vba
   - **Inventario** → HojaInventario_Nuevo.vba
   - **Concentrado** → HojaConcentrado_Nuevo.vba

#### D) Código de ThisWorkbook
1. En el panel izquierdo, haz doble clic en "ThisWorkbook"
2. Copia el contenido de **ThisWorkbook_Nuevo.vba**

### Paso 4: Guardar y Probar
1. Guarda el archivo (Ctrl + S)
2. Cierra el Editor VBA (Alt + Q)
3. Cierra y vuelve a abrir el archivo Excel
4. Habilita las macros cuando se solicite

---

## 📊 Flujo de Trabajo del Sistema

### 1. HOJA ENTRADAS (Captura Inicial)
**Validación:** Los 3 campos deben estar completos antes de registrar:
- **Referencia (B):** Código del producto
- **Lote (D):** Número de lote
- **Remisión (G):** Número de remisión

**Comportamiento:**
- Al completar la Remisión con Referencia y Lote válidos:
  - Se incrementa el inventario (+1)
  - Se actualiza el catálogo automáticamente
- Si se borra o cambia algún dato:
  - Se revierten las existencias anteriores
  - Se actualiza el catálogo

### 2. HOJA CATÁLOGO (Auto-generada)
**Columnas:**
- **Descripción (B):** Descripción del producto
- **ID/Referencia (C):** Código de referencia
- **Lote (D):** Número de lote
- **Cantidad de lotes repetidos (E):** Contador automático

**Comportamiento:**
- Se actualiza automáticamente desde Entradas
- Cuenta cuántas veces se registra cada combinación Referencia+Lote
- Principalmente de consulta (no modificar manualmente la columna E)

### 3. HOJA INVENTARIO (Control de Existencias)
**Columnas:**
- **Referencia (B):** Código del producto
- **Lote (C):** Número de lote
- **Cantidad (D):** Existencia actual

**Validaciones:**
- ❌ No se puede ingresar Cantidad antes que Lote
- ❌ No se puede ingresar Lote antes que Referencia
- ❌ No se permiten duplicados (misma Referencia + Lote)
- ⚠️ Si Cantidad llega a 0, pregunta si desea eliminar el registro

**Comportamiento:**
- Cada entrada desde Entradas suma +1
- Cada consumo desde Concentrado resta la cantidad
- Si se cambia/borra un dato con Remisión, revierte existencias

### 4. HOJA CONCENTRADO (Registro de Consumos)
**Estructura:** 30 grupos de 4 columnas (Q hasta EF)
Cada grupo tiene:
- **Material:** Descripción del material (**AUTO-GENERADO con fórmulas**)
- **Referencia:** Código de referencia (**AUTO-GENERADO con listas desplegables**)
- **Cantidad:** Cantidad consumida (**MANUAL** - Usuario captura)
- **Lote:** Número de lote (**MANUAL** - Usuario captura)

> **IMPORTANTE:** Solo debes capturar manualmente **Cantidad** y **Lote**. Material y Referencia se llenan automáticamente.

**Validaciones:**
- ❌ No se puede ingresar Cantidad antes que Lote
- ❌ No se puede ingresar Cantidad si no hay Referencia (auto-generada)
- ⚠️ Al cambiar Lote con cantidad existente → Devuelve inventario y limpia cantidad

**Navegación con Enter (optimizada para captura manual):**
- Desde **Cantidad** → Salta +1 al **Lote** del mismo grupo
- Desde **Lote** → Salta +3 a la **Cantidad** del siguiente grupo
- Esto evita las columnas auto-generadas (Material y Referencia)

**Comportamiento del Inventario:**
- ✅ Al ingresar Cantidad → Descuenta del inventario
- ✅ Al reducir Cantidad → Devuelve diferencia al inventario
- ✅ Al borrar Cantidad → Devuelve todo al inventario
- ⚠️ Verifica disponibilidad antes de descontar (muestra error si no hay suficiente)

---

## 🔐 Funciones de Seguridad

### Contraseña del Sistema
**Contraseña:** `2024comerlat`

### Funciones Disponibles

#### 1. Mostrar/Ocultar Inventario
```vba
Call MostrarOcultarInventario_Nuevo
```
Requiere contraseña para alternar visibilidad de la hoja Inventario.

#### 2. Proteger Hojas
```vba
Call ProtegerHojas
```
Protege hojas sensibles (Inventario y Catálogo).

#### 3. Desproteger Hojas
```vba
Call DesprotegerHojas
```
Requiere contraseña para desproteger hojas.

#### 4. Limpiar Datos de Prueba
```vba
Call LimpiarDatosPrueba
```
Limpia todos los datos de todas las hojas (requiere contraseña y confirmación).

#### 5. Generar Reporte de Inventario
```vba
Call GenerarReporteInventario
```
Genera un reporte en pantalla del inventario actual.

---

## 📍 Tabla Rápida de Columnas - Concentrado

Ver archivo **REFERENCIA_COLUMNAS.md** para la tabla completa.

### Resumen por Tipo de Columna:

| Tipo | Patrón | Columnas (letras) | Números |
|------|--------|------------------|---------|
| **Material** | Cada 4 desde Q | Q, U, Y, AC, AG... EC | 17, 21, 25... 133 |
| **Referencia** | Cada 4 desde R | R, V, Z, AD, AH... ED | 18, 22, 26... 134 |
| **Cantidad** | Cada 4 desde S | S, W, AA, AE, AI... EE | 19, 23, 27... 135 |
| **Lote** | Cada 4 desde T | T, X, AB, AF, AJ... EF | 20, 24, 28... 136 |

**Total:** 30 grupos de materiales (120 columnas en total)

---

## ⚠️ Mensajes de Error Comunes

### "Debe ingresar primero el LOTE y después la CANTIDAD"
- **Causa:** Intentaste poner cantidad sin lote
- **Solución:** Primero ingresa el lote, luego la cantidad

### "No se puede registrar la remisión: falta Referencia o Lote"
- **Causa:** Intentaste poner remisión sin completar referencia o lote
- **Solución:** Completa referencia y lote primero

### "Inventario insuficiente"
- **Causa:** No hay suficiente inventario para el consumo solicitado
- **Solución:** Verifica el inventario disponible en Hoja Inventario

### "Ya existe un registro con esta combinación de Referencia y Lote"
- **Causa:** Duplicado en Inventario
- **Solución:** Usa el registro existente o modifica el lote

---

## 🎯 Casos de Uso Típicos

### Caso 1: Registrar Nueva Entrada
1. Ve a **Hoja Entradas**
2. Ingresa **Referencia** (columna B)
3. Ingresa **Lote** (columna D)
4. Ingresa **Remisión** (columna G)
5. ✅ Se actualiza automáticamente Catálogo e Inventario

### Caso 2: Registrar Consumo
1. Ve a **Hoja Concentrado**
2. Las columnas **Material** y **Referencia** ya están auto-generadas (con fórmulas/listas)
3. Posiciónate en una columna de **Cantidad** (S, W, AA, AE, etc.)
4. Ingresa el **Lote** primero (columna T, X, AB, AF, etc.)
5. Presiona Enter (te lleva a Cantidad)
6. Ingresa la **Cantidad** consumida
7. Presiona Enter (salta al Lote del siguiente grupo)
8. ✅ Se descuenta automáticamente del inventario

### Caso 3: Corregir una Entrada
1. Ve a **Hoja Entradas**
2. Modifica Referencia, Lote o Remisión
3. ✅ El sistema devuelve inventario anterior y registra el nuevo

### Caso 4: Consultar Inventario
1. Ejecuta macro: `GenerarReporteInventario` (Alt+F8)
2. O ve directamente a **Hoja Inventario** (si está visible)

---

## 🔧 Mantenimiento y Personalización

### Cambiar Contraseña
Edita en los archivos:
- **Modulo2_Nuevo.vba** → `MostrarOcultarInventario_Nuevo`
- **ThisWorkbook_Nuevo.vba** → `DesprotegerHojas`, `LimpiarDatosPrueba`

Busca: `"2024comerlat"` y reemplaza por tu contraseña

### Agregar Más Grupos en Concentrado
1. Actualiza los rangos en **HojaConcentrado_Nuevo.vba**
2. Modifica las condiciones: `If Target.Column >= 18 And Target.Column <= XXX`
3. Actualiza **REFERENCIA_COLUMNAS.md**

### Personalizar Validaciones
Edita las funciones `Worksheet_Change` en cada hoja según necesites.

---

## 📞 Soporte

Para dudas o problemas:
1. Revisa **REFERENCIA_COLUMNAS.md** para verificar columnas
2. Verifica que todas las hojas tengan los nombres correctos
3. Asegúrate de que las macros estén habilitadas
4. Verifica que la contraseña sea correcta: `2024comerlat`

---

## ✅ Checklist de Implementación

- [ ] Crear archivo Excel (.xlsm)
- [ ] Crear las 4 hojas (Entradas, Catalogo, Inventario, Concentrado)
- [ ] Insertar 2 módulos (Modulo1_Nuevo y Modulo2_Nuevo)
- [ ] Copiar código en cada hoja
- [ ] Copiar código en ThisWorkbook
- [ ] Configurar encabezados de columnas
- [ ] Guardar y cerrar Excel
- [ ] Abrir y habilitar macros
- [ ] Probar con datos de ejemplo
- [ ] Verificar que el inventario se actualice correctamente
- [ ] Probar navegación con Enter en Concentrado
- [ ] Verificar validaciones de cada hoja

---

## 📝 Notas Adicionales

- **Backup:** Guarda copias de seguridad regularmente
- **Pruebas:** Usa la función `LimpiarDatosPrueba` para resetear y practicar
- **Protección:** Usa `ProtegerHojas` antes de entregar a usuarios finales
- **Performance:** Para mejor rendimiento, desactiva cálculo automático con muchos datos

---

**Versión:** 1.0  
**Fecha:** Enero 2026  
**Compatibilidad:** Excel 2016 o superior con macros habilitadas
