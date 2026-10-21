/*=========================================================
 SIFARMU v5.1 - MODO LIBRO DIARIO (LEDGER) CERO-LOGIN
=========================================================*/

const HOJAS = {
  DESPACHO: "00_DESPACHO",
  MEDICAMENTOS: "02_MEDICAMENTOS",
  LOTES: "03_LOTES",
  MOVIMIENTOS: "04_MOVIMIENTOS",
  KARDEX: "05_KARDEX"
};

/**
 * Trigger automático: Detecta si el usuario cambia el estado de una fila.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  if (sheet.getName() !== HOJAS.DESPACHO) return;
  
  // Validar si editaron la Columna 6 (Estado Acción) usando getValue() que es 100% seguro
  if (e.range.getColumn() === 6) {
    const valorCelda = String(e.range.getValue()).trim();
    
    if (valorCelda === "APLICAR SALIDA") {
      procesarFilaDespacho(sheet, e.range.getRow(), e.range);
    }
  }
}

function procesarFilaDespacho(hojaDespacho, fila, celdaEstado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const nombreCompleto = String(hojaDespacho.getRange(fila, 2).getValue()).trim();
  const cantidad = Number(hojaDespacho.getRange(fila, 3).getValue());
  const persona = String(hojaDespacho.getRange(fila, 4).getValue()).trim();
  const ubicacion = String(hojaDespacho.getRange(fila, 5).getValue()).trim();
  
  if (!nombreCompleto || cantidad <= 0 || isNaN(cantidad)) {
    celdaEstado.setValue("❌ FALTAN DATOS");
    return;
  }
  
  const responsable = persona || "Anónimo";
  const destino = ubicacion || "No especificado";
  
  try {
    registrarSalidaOptimizada(ss, nombreCompleto, cantidad, destino, responsable, "Despacho Fila " + fila);
    celdaEstado.setValue("✅ PROCESADO");
  } catch(err) {
    // Captura el error exacto y lo imprime en la celda para debug
    celdaEstado.setValue("❌ " + err.message);
  }
}

function registrarSalidaOptimizada(ss, llaveBuscada, cantidadRequerida, destino, responsable, observaciones) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(10000); 
    const hojaLotes = ss.getSheetByName(HOJAS.LOTES);
    
    if (!hojaLotes) throw new Error("No se encontró 03_LOTES");
    
    const lotesData = hojaLotes.getDataRange().getValues();
    let lotesActivos = [];
    
    for(let i = 1; i < lotesData.length; i++) {
      const celdaLote = String(lotesData[i][1]).trim();
      const estado = String(lotesData[i][10]).trim();
      const disponible = Number(lotesData[i][4]);
      
      if(celdaLote === llaveBuscada && estado === "ACTIVO" && disponible > 0) {
        lotesActivos.push({
          filaHoja: i + 1,
          disponible: disponible,
          vencimiento: new Date(lotesData[i][7])
        });
      }
    }
    
    lotesActivos.sort((a, b) => a.vencimiento - b.vencimiento);
    
    let stockAnterior = lotesActivos.reduce((acc, lote) => acc + lote.disponible, 0);
    
    if (stockAnterior < cantidadRequerida) {
      throw new Error("STOCK INSUFICIENTE");
    }
    
    let restante = cantidadRequerida;
    let actualizaciones = []; 
    
    for(let i = 0; i < lotesActivos.length; i++) {
      if(restante <= 0) break;
      
      let lote = lotesActivos[i];
      let descontar = Math.min(lote.disponible, restante);
      
      lote.disponible -= descontar;
      restante -= descontar;
      
      actualizaciones.push({
        fila: lote.filaHoja,
        nuevoDisponible: lote.disponible
      });
    }
    
    actualizaciones.forEach(act => {
      hojaLotes.getRange(act.fila, 5).setValue(act.nuevoDisponible); 
      if(act.nuevoDisponible === 0) {
        hojaLotes.getRange(act.fila, 11).setValue("AGOTADO"); 
      }
    });
    
    const hojaKardex = ss.getSheetByName(HOJAS.KARDEX);
    const hojaMovimientos = ss.getSheetByName(HOJAS.MOVIMIENTOS);
    const fecha = new Date();
    
    const pad = (n) => (n < 10 ? '0' + n : n);
    const dateStr = pad(fecha.getDate()) + '/' + pad(fecha.getMonth() + 1) + '/' + fecha.getFullYear();
    const timeStr = pad(fecha.getHours()) + ':' + pad(fecha.getMinutes()) + ':' + pad(fecha.getSeconds());
    const randomId = Math.random().toString(36).substring(2, 10).toUpperCase();
    
    hojaKardex.appendRow([
      randomId, dateStr, timeStr, llaveBuscada, "SALIDA",
      cantidadRequerida, stockAnterior, stockAnterior - cantidadRequerida,
      "FEFO Fila", responsable, observaciones
    ]);
    
    hojaMovimientos.appendRow([
      randomId, dateStr, timeStr, "SALIDA", llaveBuscada,
      cantidadRequerida, "Múltiples (FEFO)", destino, responsable, observaciones
    ]);
    
    return true;

  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}
