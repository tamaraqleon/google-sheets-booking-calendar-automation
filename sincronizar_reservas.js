/**
 * Servicio de Sincronización Bidireccional: Google Sheets <-> Google Calendar API
 * Gestiona validación de colisiones, cálculo de balances y ciclo de vida de eventos.
 */
function syncCalendarReservations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const START_ROW = 3;
  const lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) return;

  // Rango de columnas operativas (16 columnas de datos)
  const range = sheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, 16);
  const data = range.getValues();
  const calendar = CalendarApp.getDefaultCalendar();
  const colisiones = [];

  for (let i = 0; i < data.length; i++) {
    const fila = data[i];
    const filaIndex = START_ROW + i;

    // Extracción y sanitización de campos base
    const titular         = fila[0] ? fila[0].toString().trim() : "";
    const tipoEvento      = fila[1] ? fila[1].toString().trim() : "Evento Operativo";
    const entidad         = fila[2] ? fila[2].toString().trim() : "";
    const contacto        = fila[3] ? fila[3].toString().trim() : "Por confirmar";
    const fecha           = new Date(fila[4]);
    const horario         = fila[5] ? fila[5].toString().trim() : "Por definir";
    const cantAdultos     = fila[6] !== "" ? fila[6] : "-";
    const cantNinos       = fila[7] !== "" ? fila[7] : "-";
    const cantInfantes    = fila[8] !== "" ? fila[8] : "-";
    const serviciosExtra  = fila[9] !== "" ? fila[9].toString().trim() : "0";
    const descuento       = fila[10] !== "" ? fila[10].toString().trim() : "0";
    const montoTotal      = fila[11] !== "" ? fila[11] : 0;
    const montoAbonado    = fila[12] !== "" ? fila[12] : 0;
    const saldoPendiente  = fila[13] !== "" ? fila[13].toString().trim() : "";
    const estado          = fila[14] ? fila[14].toString().trim() : "";
    let eventoId          = fila[15] ? fila[15].toString().trim() : "";

    // Consolidación del aforo
    const numAdultos = Number(cantAdultos) || 0;
    const numNinos = Number(cantNinos) || 0;
    const numInfantes = Number(cantInfantes) || 0;
    const aforoTotal = numAdultos + numNinos + numInfantes;
    const aforoDetalle = aforoTotal > 0 ? `${aforoTotal} pers.` : "Aforo sin confirmar";

    // Evaluación de estado contable
    const saldoUpper = saldoPendiente.toUpperCase();
    let estadoPago = "PENDIENTE";
    if (saldoUpper === "PAGADO" || descuento === "-") {
      estadoPago = "PAGADO";
    } else if (Number(montoTotal) === 0) {
      estadoPago = "POR DEFINIR";
    }

    // Manejo de revocación / bajas
    const estadoLower = estado.toLowerCase();
    const esBaja = ["cancelar", "cancelado", "eliminar", "borrar"].includes(estadoLower);
    
    if (esBaja && eventoId) {
      try {
        const evento = calendar.getEventById(eventoId);
        if (evento) evento.deleteEvent();
      } catch (err) {
        console.warn(`No se pudo eliminar el evento ${eventoId}: ${err.message}`);
      }
      sheet.getRange(filaIndex, 15).setValue("Cancelado");
      sheet.getRange(filaIndex, 16).setValue("");
      continue;
    }

    // Procesamiento y sincronización de registros activos
    const tieneIdentificador = titular !== "" || tipoEvento !== "" || entidad !== "";
    if (!isNaN(fecha.getTime()) && tieneIdentificador && estadoLower !== "cancelado") {

      // Prevención de colisiones de agenda para registros nuevos
      if (!eventoId) {
        const eventosExistentes = calendar.getEventsForDay(fecha);
        if (eventosExistentes.length > 0) {
          sheet.getRange(filaIndex, 15).setValue("Conflicto: Fecha ocupada");
          colisiones.push(`Fila ${filaIndex} (${tipoEvento || titular}) - ${fecha.toLocaleDateString()}`);
          continue;
        }
      }

      // Construcción del título normalizado
      const componentesTitulo = [tipoEvento];
      if (titular) componentesTitulo.push(titular);
      if (entidad) componentesTitulo.push(`(${entidad})`);
      componentesTitulo.push(`(${aforoDetalle})`);
      const tituloEvento = componentesTitulo.join(" - ");

      // Formato monetario
      const formatoMoneda = (valor) => "$" + Number(valor).toLocaleString();
      const txtDescuento = (descuento !== "0" && descuento !== "")
        ? `- Descuento / Bonificación: ${isNaN(Number(descuento)) ? descuento : formatoMoneda(descuento)}\n`
        : "";

      let displaySaldo = "$0";
      if (saldoUpper === "PAGADO") {
        displaySaldo = "$0 (PAGADO)";
      } else if (!isNaN(Number(saldoPendiente)) && Number(saldoPendiente) > 0) {
        displaySaldo = formatoMoneda(saldoPendiente);
      }

      // Payload descriptivo estructurado
      const descripcionEvento = 
        "DETALLES DEL REGISTRO OPERATIVO\n" +
        "-----------------------------------------\n" +
        `Tipo de Evento: ${tipoEvento}\n` +
        `Entidad / Dependencia: ${entidad || "No especificada"}\n` +
        `Responsable: ${titular || "Por confirmar"}\n` +
        `Teléfono de Contacto: ${contacto}\n` +
        `Horario: ${horario}\n\n` +
        "ASISTENTES:\n" +
        `- Adultos: ${cantAdultos}\n` +
        `- Niños: ${cantNinos}\n` +
        `- Menores: ${cantInfantes}\n` +
        `- Requerimientos extra: ${serviciosExtra}\n\n` +
        `ESTADO FINANCIERO (${estadoPago}):\n` +
        txtDescuento +
        `- Total: ${Number(montoTotal) > 0 ? formatoMoneda(montoTotal) : "Por definir"}\n` +
        `- Abonado: ${Number(montoAbonado) > 0 ? formatoMoneda(montoAbonado) : "$0"}\n` +
        `- Saldo Pendiente: ${displaySaldo}`;

      // Actualización de registro preexistente
      if (eventoId) {
        try {
          const ev = calendar.getEventById(eventoId);
          if (ev) {
            ev.setTitle(tituloEvento);
            ev.setDescription(descripcionEvento);
            ev.setAllDayDate(fecha);
          }
        } catch (err) {
          console.error(`Fallo en actualización de evento ${eventoId}: ${err.message}`);
          eventoId = ""; // Si no existe en calendar, se fuerza la re-creación
        }
      }

      // Creación de nuevo evento e indexación de ID persistente
      if (!eventoId) {
        const nuevoEvento = calendar.createAllDayEvent(tituloEvento, fecha, {
          description: descripcionEvento
        });
        sheet.getRange(filaIndex, 15).setValue("Agendado");
        sheet.getRange(filaIndex, 16).setValue(nuevoEvento.getId());
      }
    }
  }

  // Notificación de advertencia en caso de sobrecupos detectados
  if (colisiones.length > 0) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Alerta de Colisión: Fechas Ocupadas",
      "Las siguientes entradas no pudieron agendarse debido a eventos concurrentes en Google Calendar:\n\n" +
      colisiones.join("\n") +
      "\n\nVerifique la disponibilidad del calendario antes de reintentar la sincronización.",
      ui.ButtonSet.OK
    );
  }
}
