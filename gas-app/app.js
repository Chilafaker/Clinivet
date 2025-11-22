// Utilidades de comunicación con Apps Script
const api = (() => {
  const run = (fn, ...args) =>
    new Promise((resolve, reject) => {
      google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[fn](...args);
    });

  return {
    buscarClientePorTelefono: (telefono) => run('getClienteByTelefono', telefono),
    guardarCliente: (datos) => run('crearOActualizarCliente', datos),
    guardarPaciente: (datos) => run('crearOActualizarPaciente', datos),
    registrarServicio: (datos) => run('registrarServicio', datos),
    buscarPacientes: (query) => run('buscarPacientes', query),
    precioServicio: (nombre) => run('getPrecioServicio', nombre),
    listaServicios: () => run('getListaServicios'),
    historialServicios: (idPaciente) => run('getHistorialServicios', idPaciente),
  };
})();

// Utilidades comunes de UI
const ui = (() => {
  const fillSelect = (select, items, valueKey, labelKey) => {
    select.innerHTML = '';
    items.forEach((item) => {
      const option = document.createElement('option');
      option.value = item[valueKey];
      option.textContent = item[labelKey];
      select.appendChild(option);
    });
  };

  const fillTable = (tableBody, rows, columns) => {
    tableBody.innerHTML = '';
    rows.forEach((row) => {
      const tr = document.createElement('tr');
      columns.forEach((col) => {
        const td = document.createElement('td');
        td.textContent = row[col] ?? '';
        tr.appendChild(td);
      });
      tableBody.appendChild(tr);
    });
  };

  const handleError = (error) => {
    console.error(error);
    alert(error.message || 'Ocurrió un error');
  };

  return { fillSelect, fillTable, handleError };
})();

// Módulo de entrada (registro de clientes y mascotas)
const entradaModule = (() => {
  const telefonoInput = document.getElementById('telefono');
  const clienteForm = document.getElementById('cliente-form');
  const pacienteForm = document.getElementById('paciente-form');
  const pacientesList = document.getElementById('pacientes-list');

  const renderPacientes = (pacientes) => {
    if (!pacientesList) return;
    pacientesList.innerHTML = '';
    pacientes.forEach((paciente) => {
      const li = document.createElement('li');
      li.textContent = `${paciente.ID_Paciente} - ${paciente.Nombre_Paciente} (${paciente.Especie || ''})`;
      pacientesList.appendChild(li);
    });
  };

  const loadCliente = async () => {
    if (!telefonoInput || !telefonoInput.value) return;
    try {
      const response = await api.buscarClientePorTelefono(telefonoInput.value.trim());
      if (!response) {
        alert('No se encontró cliente, completa los datos para crearlo.');
        if (clienteForm) clienteForm.reset();
        renderPacientes([]);
        return;
      }
      const { cliente, pacientes } = response;
      if (clienteForm) {
        clienteForm.elements['ID_Cliente'].value = cliente.ID_Cliente || '';
        clienteForm.elements['Nombre_Dueno'].value = cliente.Nombre_Dueno || '';
        clienteForm.elements['Telefono'].value = cliente.Telefono || '';
        clienteForm.elements['Mascotas'].value = cliente.Mascotas || '';
        clienteForm.elements['Servicios_Frecuentes'].value = cliente.Servicios_Frecuentes || '';
        clienteForm.elements['Notas'].value = cliente.Notas || '';
        clienteForm.elements['Fecha_Primera_Visita'].value = cliente.Fecha_Primera_Visita || '';
        clienteForm.elements['Fecha_Ultima_Visita'].value = cliente.Fecha_Ultima_Visita || '';
      }
      renderPacientes(pacientes || []);
    } catch (error) {
      ui.handleError(error);
    }
  };

  const saveCliente = async (event) => {
    event.preventDefault();
    if (!clienteForm) return;
    const formData = Object.fromEntries(new FormData(clienteForm).entries());
    try {
      const saved = await api.guardarCliente(formData);
      alert(`Cliente guardado con ID ${saved.ID_Cliente}`);
    } catch (error) {
      ui.handleError(error);
    }
  };

  const savePaciente = async (event) => {
    event.preventDefault();
    if (!pacienteForm || !clienteForm) return;
    const clienteId = clienteForm.elements['ID_Cliente'].value;
    if (!clienteId) {
      alert('Primero guarda el cliente para asignar el ID.');
      return;
    }
    const data = Object.fromEntries(new FormData(pacienteForm).entries());
    data.ID_Cliente = clienteId;
    try {
      const saved = await api.guardarPaciente(data);
      alert(`Paciente guardado con ID ${saved.ID_Paciente}`);
      await loadCliente();
      pacienteForm.reset();
    } catch (error) {
      ui.handleError(error);
    }
  };

  const init = () => {
    if (!clienteForm) return;
    const buscarBtn = document.getElementById('buscar-cliente');
    buscarBtn?.addEventListener('click', loadCliente);
    clienteForm.addEventListener('submit', saveCliente);
    pacienteForm?.addEventListener('submit', savePaciente);
  };

  return { init };
})();

// Módulo de registro de servicios
const servicioModule = (() => {
  const servicioForm = document.getElementById('servicio-form');
  const servicioSelect = document.getElementById('Servicio');
  const precioInput = document.getElementById('Precio');
  const pacienteBusquedaInput = document.getElementById('buscar-paciente');
  const pacienteSelect = document.getElementById('ID_Paciente');

  const loadServicios = async () => {
    if (!servicioSelect) return;
    try {
      const servicios = await api.listaServicios();
      ui.fillSelect(
        servicioSelect,
        servicios,
        'Servicio',
        (item) => `${item.Servicio} (${item.Categoria || 'General'})`
      );
    } catch (error) {
      ui.handleError(error);
    }
  };

  const updatePrecio = async () => {
    if (!servicioSelect || !precioInput) return;
    const selected = servicioSelect.value;
    if (!selected) return;
    try {
      const precio = await api.precioServicio(selected);
      if (precio?.Precio !== undefined) {
        precioInput.value = precio.Precio;
      }
    } catch (error) {
      ui.handleError(error);
    }
  };

  const buscarPacientes = async () => {
    if (!pacienteBusquedaInput || !pacienteSelect) return;
    try {
      const lista = await api.buscarPacientes(pacienteBusquedaInput.value || '');
      ui.fillSelect(
        pacienteSelect,
        lista,
        'ID_Paciente',
        (item) => `${item.ID_Paciente} - ${item.Nombre_Paciente}`
      );
    } catch (error) {
      ui.handleError(error);
    }
  };

  const saveServicio = async (event) => {
    event.preventDefault();
    if (!servicioForm) return;
    const data = Object.fromEntries(new FormData(servicioForm).entries());
    try {
      const saved = await api.registrarServicio(data);
      alert(`Servicio registrado con ID ${saved.ID_Servicio}`);
      servicioForm.reset();
      await buscarPacientes();
    } catch (error) {
      ui.handleError(error);
    }
  };

  const init = () => {
    if (!servicioForm) return;
    loadServicios();
    buscarPacientes();
    servicioSelect?.addEventListener('change', updatePrecio);
    pacienteBusquedaInput?.addEventListener('input', buscarPacientes);
    servicioForm.addEventListener('submit', saveServicio);
  };

  return { init };
})();

// Módulo de historial
const historialModule = (() => {
  const pacienteBusquedaInput = document.getElementById('buscar-historial');
  const pacienteSelect = document.getElementById('paciente-historial');
  const tablaBody = document.getElementById('historial-body');
  const columnas = [
    'ID_Servicio',
    'Fecha',
    'Servicio',
    'Precio',
    'Metodo_Pago',
    'Proxima_Visita',
    'Notas',
  ];

  const buscarPacientes = async () => {
    if (!pacienteBusquedaInput || !pacienteSelect) return;
    try {
      const lista = await api.buscarPacientes(pacienteBusquedaInput.value || '');
      ui.fillSelect(
        pacienteSelect,
        lista,
        'ID_Paciente',
        (item) => `${item.ID_Paciente} - ${item.Nombre_Paciente}`
      );
      if (lista[0]) {
        pacienteSelect.value = lista[0].ID_Paciente;
        await loadHistorial();
      }
    } catch (error) {
      ui.handleError(error);
    }
  };

  const loadHistorial = async () => {
    if (!pacienteSelect || !tablaBody) return;
    try {
      const historial = await api.historialServicios(pacienteSelect.value);
      ui.fillTable(tablaBody, historial, columnas);
    } catch (error) {
      ui.handleError(error);
    }
  };

  const init = () => {
    if (!pacienteSelect) return;
    pacienteBusquedaInput?.addEventListener('input', buscarPacientes);
    pacienteSelect.addEventListener('change', loadHistorial);
    buscarPacientes();
  };

  return { init };
})();

// Módulo de inventario: muestra Lista_Precios como referencia rápida
const inventarioModule = (() => {
  const tablaBody = document.getElementById('inventario-body');
  const columnas = ['Servicio', 'Categoria', 'Precio'];

  const loadInventario = async () => {
    if (!tablaBody) return;
    try {
      const servicios = await api.listaServicios();
      ui.fillTable(tablaBody, servicios, columnas);
    } catch (error) {
      ui.handleError(error);
    }
  };

  const init = () => {
    if (!tablaBody) return;
    loadInventario();
  };

  return { init };
})();

// Módulo de dashboard (placeholder con datos clave)
const dashboardModule = (() => {
  const resumenServicios = document.getElementById('resumen-servicios');

  const loadResumen = async () => {
    if (!resumenServicios) return;
    try {
      const lista = await api.listaServicios();
      const categorias = lista.reduce((acc, item) => {
        acc[item.Categoria] = (acc[item.Categoria] || 0) + 1;
        return acc;
      }, {});
      resumenServicios.innerHTML = Object.entries(categorias)
        .map(([categoria, total]) => `<li>${categoria}: ${total} servicios</li>`)
        .join('');
    } catch (error) {
      ui.handleError(error);
    }
  };

  const init = () => {
    if (!resumenServicios) return;
    loadResumen();
  };

  return { init };
})();

// Inicialización según página
(function bootstrap() {
  document.addEventListener('DOMContentLoaded', () => {
    entradaModule.init();
    servicioModule.init();
    historialModule.init();
    inventarioModule.init();
    dashboardModule.init();
  });
})();
