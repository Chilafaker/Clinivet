const SPREADSHEET_NAME = 'PV_Base_Datos';
const SHEETS = {
  CLIENTES: 'Base_de_Clientes',
  PACIENTES: 'Pacientes',
  SERVICIOS: 'Servicios_Realizados',
  PRECIOS: 'Lista_Precios',
};

const HEADERS = {
  CLIENTES: [
    'ID_Cliente',
    'Nombre_Dueno',
    'Telefono',
    'Fecha_Primera_Visita',
    'Fecha_Ultima_Visita',
    'Mascotas',
    'Servicios_Frecuentes',
    'Notas',
  ],
  PACIENTES: [
    'ID_Paciente',
    'ID_Cliente',
    'Nombre_Paciente',
    'Especie',
    'Raza',
    'Edad',
    'Peso',
    'Alergias',
    'Frecuencia_Bano',
    'Fecha_Proxima_Vacuna',
    'Fecha_Proximo_Bano',
    'Notas_Clinicas',
  ],
  SERVICIOS: [
    'ID_Servicio',
    'Fecha',
    'ID_Paciente',
    'Nombre_Paciente',
    'Servicio',
    'Precio',
    'Metodo_Pago',
    'Upsell',
    'Medicamentos',
    'Hallazgos',
    'Proxima_Visita',
    'Notas',
  ],
  PRECIOS: ['Servicio', 'Precio', 'Categoria'],
};

let cachedSpreadsheet = null;

/**
 * Obtiene el archivo PV_Base_Datos por nombre y lo mantiene en caché.
 */
function getSpreadsheet() {
  if (cachedSpreadsheet) return cachedSpreadsheet;
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (!files.hasNext()) {
    throw new Error(`No se encontró el archivo de Google Sheets "${SPREADSHEET_NAME}".`);
  }
  cachedSpreadsheet = SpreadsheetApp.open(files.next());
  return cachedSpreadsheet;
}

/**
 * Devuelve una hoja existente o lanza error si no está presente.
 * @param {string} name
 */
function getSheet(name) {
  const sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`La hoja "${name}" no existe en ${SPREADSHEET_NAME}.`);
  }
  return sheet;
}

/**
 * Crea un mapa de encabezados => índice para facilitar lecturas/escrituras.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});
}

/**
 * Convierte una fila en objeto usando el mapa de encabezados.
 * @param {any[]} row
 * @param {Record<string, number>} headerMap
 */
function rowToObject(row, headerMap) {
  return Object.keys(headerMap).reduce((acc, key) => {
    acc[key] = row[headerMap[key]] ?? '';
    return acc;
  }, {});
}

/**
 * Obtiene el siguiente ID disponible con prefijo y relleno de ceros.
 */
function getNextId(sheet, prefix) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return `${prefix}0001`;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const max = ids
    .filter((id) => typeof id === 'string' && id.startsWith(prefix))
    .reduce((acc, id) => {
      const number = parseInt(id.replace(prefix, ''), 10);
      return isNaN(number) ? acc : Math.max(acc, number);
    }, 0);
  const nextNumber = max + 1;
  return `${prefix}${nextNumber.toString().padStart(4, '0')}`;
}

/**
 * Busca la primera fila que coincida con un valor en una columna.
 */
function findRowByColumn(sheet, headerMap, columnName, value) {
  const colIndex = headerMap[columnName];
  if (colIndex === undefined) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    if (`${values[i][colIndex]}` === `${value}`) {
      return { row: i + 2, data: values[i] };
    }
  }
  return null;
}

/**
 * Busca un cliente por teléfono y devuelve sus pacientes asociados.
 */
function getClienteByTelefono(telefono) {
  const sheet = getSheet(SHEETS.CLIENTES);
  const headerMap = getHeaderMap(sheet);
  const match = findRowByColumn(sheet, headerMap, 'Telefono', telefono);
  if (!match) return null;
  const cliente = rowToObject(match.data, headerMap);
  const pacientes = getPacientesByCliente(cliente.ID_Cliente);
  return { cliente, pacientes };
}

/**
 * Obtiene pacientes vinculados a un cliente.
 */
function getPacientesByCliente(idCliente) {
  const sheet = getSheet(SHEETS.PACIENTES);
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return values
    .filter((row) => `${row[headerMap['ID_Cliente']]}` === `${idCliente}`)
    .map((row) => rowToObject(row, headerMap));
}

/**
 * Crea o actualiza un cliente usando el teléfono como llave única.
 */
function crearOActualizarCliente(datosCliente) {
  if (!datosCliente.Telefono) {
    throw new Error('El teléfono del cliente es obligatorio.');
  }
  const sheet = getSheet(SHEETS.CLIENTES);
  const headerMap = getHeaderMap(sheet);
  const existingByPhone = findRowByColumn(sheet, headerMap, 'Telefono', datosCliente.Telefono);
  const existingById = datosCliente.ID_Cliente
    ? findRowByColumn(sheet, headerMap, 'ID_Cliente', datosCliente.ID_Cliente)
    : null;
  const existing = existingByPhone || existingById;
  const nextId = getNextId(sheet, 'C');

  if (existing) {
    const existingObj = rowToObject(existing.data, headerMap);
    const mergedObj = HEADERS.CLIENTES.reduce((acc, key) => {
      if (key === 'ID_Cliente') {
        acc[key] = existingObj[key] || datosCliente[key] || nextId;
      } else {
        acc[key] = datosCliente[key] ?? existingObj[key] ?? '';
      }
      return acc;
    }, {});
    const mergedRow = HEADERS.CLIENTES.map((key) => mergedObj[key] ?? '');
    sheet.getRange(existing.row, 1, 1, HEADERS.CLIENTES.length).setValues([mergedRow]);
    return mergedObj;
  }

  const newRecord = HEADERS.CLIENTES.reduce((acc, key) => {
    acc[key] = key === 'ID_Cliente' ? nextId : datosCliente[key] ?? '';
    return acc;
  }, {});
  sheet.appendRow(HEADERS.CLIENTES.map((key) => newRecord[key] ?? ''));
  return newRecord;
}

/**
 * Crea o actualiza un paciente. Si no se envía ID, se genera uno nuevo.
 */
function crearOActualizarPaciente(datosPaciente) {
  if (!datosPaciente.ID_Cliente) {
    throw new Error('El paciente debe estar asociado a un ID_Cliente.');
  }
  const sheet = getSheet(SHEETS.PACIENTES);
  const headerMap = getHeaderMap(sheet);
  const existing = datosPaciente.ID_Paciente
    ? findRowByColumn(sheet, headerMap, 'ID_Paciente', datosPaciente.ID_Paciente)
    : null;
  const nextId = getNextId(sheet, 'P');

  if (existing) {
    const existingObj = rowToObject(existing.data, headerMap);
    const mergedObj = HEADERS.PACIENTES.reduce((acc, key) => {
      if (key === 'ID_Paciente') {
        acc[key] = existingObj[key] || datosPaciente[key] || nextId;
      } else {
        acc[key] = datosPaciente[key] ?? existingObj[key] ?? '';
      }
      return acc;
    }, {});
    const mergedRow = HEADERS.PACIENTES.map((key) => mergedObj[key] ?? '');
    sheet.getRange(existing.row, 1, 1, HEADERS.PACIENTES.length).setValues([mergedRow]);
    return mergedObj;
  }

  const newRecord = HEADERS.PACIENTES.reduce((acc, key) => {
    acc[key] = key === 'ID_Paciente' ? nextId : datosPaciente[key] ?? '';
    return acc;
  }, {});
  sheet.appendRow(HEADERS.PACIENTES.map((key) => newRecord[key] ?? ''));
  return newRecord;
}

/**
 * Registra un servicio realizado tomando el precio de Lista_Precios cuando no se envía.
 */
function registrarServicio(datosServicio) {
  const sheet = getSheet(SHEETS.SERVICIOS);
  const headerMap = getHeaderMap(sheet);
  const nextId = getNextId(sheet, 'S');
  const precio = datosServicio.Precio ?? getPrecioServicio(datosServicio.Servicio)?.Precio ?? '';

  const newRecord = HEADERS.SERVICIOS.reduce((acc, key) => {
    if (key === 'ID_Servicio') acc[key] = datosServicio[key] || nextId;
    else if (key === 'Precio') acc[key] = precio;
    else acc[key] = datosServicio[key] ?? '';
    return acc;
  }, {});

  sheet.appendRow(HEADERS.SERVICIOS.map((key) => newRecord[key] ?? ''));
  return newRecord;
}

/**
 * Busca pacientes por ID exacto o coincidencia parcial del nombre.
 */
function buscarPacientes(query) {
  const sheet = getSheet(SHEETS.PACIENTES);
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const normalizedQuery = `${query}`.toLowerCase().trim();
  return values
    .map((row) => rowToObject(row, headerMap))
    .filter((paciente) => {
      if (!normalizedQuery) return true;
      return (
        `${paciente.ID_Paciente}`.toLowerCase() === normalizedQuery ||
        `${paciente.Nombre_Paciente}`.toLowerCase().includes(normalizedQuery)
      );
    });
}

/**
 * Devuelve el precio y categoría de un servicio de Lista_Precios.
 */
function getPrecioServicio(nombreServicio) {
  if (!nombreServicio) return null;
  const sheet = getSheet(SHEETS.PRECIOS);
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = rowToObject(values[i], headerMap);
    if (`${row.Servicio}`.trim().toLowerCase() === `${nombreServicio}`.trim().toLowerCase()) {
      return row;
    }
  }
  return null;
}

/**
 * Lista de servicios disponibles para poblar menús desplegables.
 */
function getListaServicios() {
  const sheet = getSheet(SHEETS.PRECIOS);
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return values.map((row) => rowToObject(row, headerMap));
}

/**
 * Historial de servicios por paciente.
 */
function getHistorialServicios(idPaciente) {
  if (!idPaciente) return [];
  const sheet = getSheet(SHEETS.SERVICIOS);
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return values
    .map((row) => rowToObject(row, headerMap))
    .filter((servicio) => `${servicio.ID_Paciente}` === `${idPaciente}`);
}

/**
 * Utilidad para incluir archivos HTML dentro de plantillas (<?!= include('app'); ?>).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Renderiza la página solicitada. Ej: ?page=servicio.
 */
function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : 'index';
  const template = HtmlService.createTemplateFromFile(page);
  return template
    .evaluate()
    .setTitle('PV-OS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
