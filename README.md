<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Sistema de Control de Trabajadores</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.0/jspdf.plugin.autotable.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@fontsource/inter@5.0.0/index.css">
<style>
  body { font-family: 'Inter', sans-serif; background: linear-gradient(135deg,#f0f4f8 0%,#e2e8f0 100%); }
  .card { background: white; border-radius: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.06); }
  .btn-primary { background: linear-gradient(135deg,#2563eb,#1e40af); color:white; transition: all .2s; }
  .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 6px 16px rgba(37,99,235,.35); }
  .btn-ghost { background:#f1f5f9; color:#334155; transition: all .2s; }
  .btn-ghost:hover { background:#e2e8f0; }
  .input { width:100%; padding:.6rem .8rem; border:1px solid #cbd5e1; border-radius:8px; font-size:.9rem; transition:all .2s; background:white; }
  .input:focus { outline:none; border-color:#2563eb; box-shadow: 0 0 0 3px rgba(37,99,235,.15); }
  .label { font-size:.8rem; font-weight:600; color:#475569; margin-bottom:.3rem; display:block; }
  .section-title { font-size:1.1rem; font-weight:700; color:#1e293b; display:flex; align-items:center; gap:.5rem; }
  .badge { display:inline-flex; align-items:center; gap:.25rem; padding:.15rem .55rem; border-radius:999px; font-size:.7rem; font-weight:600; }
  .modal-bg { background: rgba(15,23,42,.55); backdrop-filter: blur(4px); }
  .tab-active { background:#2563eb; color:white; }
  .fade-in { animation: fadeIn .3s ease; }
  @keyframes fadeIn { from{opacity:0; transform:translateY(6px);} to{opacity:1; transform:translateY(0);} }
  details > summary { list-style:none; cursor:pointer; }
  details > summary::-webkit-details-marker { display:none; }
  .chev { transition: transform .2s; }
  details[open] .chev { transform: rotate(90deg); }
  table tbody tr { transition: background .15s; }
  table tbody tr:hover { background:#f8fafc; }
  .scroll-table { max-height: 420px; overflow:auto; }
  .scroll-table thead th { position: sticky; top:0; background:#f1f5f9; z-index:1; }
  .toast { position:fixed; bottom:24px; right:24px; z-index:100; }
  select.input { appearance: none; background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e"); background-repeat: no-repeat; background-position: right .6rem center; background-size: 1em; padding-right: 2.2rem; }
</style>
</head>
<body class="min-h-screen">

<!-- HEADER -->
<header class="bg-white shadow-sm border-b border-slate-200 sticky top-0 z-40">
  <div class="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between flex-wrap gap-2">
    <div class="flex items-center gap-3">
      <div class="w-11 h-11 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-700 flex items-center justify-center text-white shadow-md">
        <svg xmlns="http://www.w3.org/2000/svg" class="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M17 20h5v-2a4 4 0 00-3-3.87M9 20H4v-2a4 4 0 013-3.87m6-5.13a4 4 0 11-8 0 4 4 0 018 0zm6 0a4 4 0 11-8 0 4 4 0 018 0z"/></svg>
      </div>
      <div>
        <h1 class="text-lg font-bold text-slate-800">Sistema de Control de Trabajadores</h1>
        <p class="text-xs text-slate-500">Gestión integral de empresa y personal</p>
      </div>
    </div>
    <div class="flex gap-2 flex-wrap">
      <button onclick="openSyncModal()" class="btn-ghost px-3 py-2 rounded-lg text-sm font-medium flex items-center gap-2" title="Sincronizar">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
        Sincronizar
      </button>
      <button onclick="exportPDF()" class="btn-primary px-3 py-2 rounded-lg text-sm font-medium flex items-center gap-2">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M12 10v6m0 0l-3-3m3 3l3-3M3 17V7a2 2 0 012-2h14a2 2 0 012 2v10a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
        PDF
      </button>
      <button onclick="exportExcel()" class="btn-primary px-3 py-2 rounded-lg text-sm font-medium flex items-center gap-2" style="background:linear-gradient(135deg,#059669,#047857);">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M9 17v-2a4 4 0 014-4h4M9 17H5a2 2 0 01-2-2V5a2 2 0 012-2h14a2 2 0 012 2v10a2 2 0 01-2 2h-4"/></svg>
        Excel
      </button>
    </div>
  </div>
</header>

<main class="max-w-7xl mx-auto px-4 py-6 space-y-6">

  <!-- DATOS DE LA EMPRESA -->
  <section class="card p-5 fade-in">
    <details open>
      <summary class="flex items-center justify-between">
        <div class="section-title">
          <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5"/></svg>
          Datos de la Empresa / Centro de Trabajo
        </div>
        <svg class="chev w-5 h-5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M9 5l7 7-7 7"/></svg>
      </summary>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4">
        <div class="md:col-span-2">
          <label class="label">Nombre de la Empresa</label>
          <input id="emp_nombre" class="input" placeholder="Ej. Inversiones El Sol C.A."/>
        </div>
        <div>
          <label class="label">RIF</label>
          <input id="emp_rif" class="input" placeholder="J-12345678-9"/>
        </div>
        <div>
          <label class="label">Centro de Trabajo</label>
          <input id="emp_centro" class="input" placeholder="Ej. Sede Principal, Planta, Local #5"/>
        </div>

        <div>
          <label class="label">Estado</label>
          <select id="emp_estado" class="input" onchange="loadMunicipios()">
            <option value="">Seleccione un estado...</option>
          </select>
        </div>
        <div>
          <label class="label">Municipio</label>
          <select id="emp_municipio" class="input" onchange="loadParroquias()" disabled>
            <option value="">Seleccione primero un estado...</option>
          </select>
        </div>
        <div>
          <label class="label">Parroquia</label>
          <select id="emp_parroquia" class="input" disabled>
            <option value="">Seleccione primero un municipio...</option>
          </select>
        </div>
        <div>
          <label class="label">Dirección de Calle</label>
          <input id="emp_calle" class="input" placeholder="Av. Principal, Edif. X, Piso 2"/>
        </div>
        <div>
          <label class="label">Sector / Urbanización</label>
          <input id="emp_sector" class="input" placeholder="Centro, El Recreo, etc."/>
        </div>

        <!-- Sucursal -->
        <div class="md:col-span-2 bg-slate-50 border border-slate-200 rounded-xl p-4">
          <div class="flex items-center justify-between">
            <div>
              <p class="font-semibold text-slate-700">¿Es una sucursal?</p>
              <p class="text-xs text-slate-500">Indique si esta sede es una sucursal y dónde está ubicada.</p>
            </div>
            <label class="relative inline-flex items-center cursor-pointer">
              <input type="checkbox" id="emp_sucursal" class="sr-only peer" onchange="toggleSucursal()">
              <div class="w-11 h-6 bg-slate-300 peer-checked:bg-blue-600 rounded-full relative transition-all after:content-[''] after:absolute after:top-0.5 after:left-0.5 after:bg-white after:rounded-full after:w-5 after:h-5 after:transition-all peer-checked:after:translate-x-5"></div>
            </label>
          </div>
          <div id="sucursal_info" class="hidden mt-4 grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label class="label">Ubicación de la Sucursal</label>
              <input id="emp_sucursal_ubicacion" class="input" placeholder="Ciudad / Zona"/>
            </div>
            <div class="flex items-center gap-3">
              <input type="checkbox" id="emp_centrocomercial" onchange="toggleCentroComercial()" class="w-4 h-4 text-blue-600 rounded">
              <label for="emp_centrocomercial" class="text-sm text-slate-700">¿Está ubicada en un Centro Comercial?</label>
            </div>
            <div id="cc_info" class="hidden md:col-span-2 grid grid-cols-1 md:grid-cols-2 gap-3 bg-white p-3 rounded-lg border border-slate-200">
              <div>
                <label class="label">Nombre del Centro Comercial</label>
                <input id="emp_cc_nombre" class="input" placeholder="Ej. Sambil, Tolón, Ciudad Traki"/>
              </div>
              <div>
                <label class="label">Ciudad del Centro Comercial</label>
                <input id="emp_cc_ciudad" class="input" placeholder="Ej. Caracas"/>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div class="flex justify-end gap-2 mt-5">
        <button onclick="clearEmpresa()" class="btn-ghost px-4 py-2 rounded-lg text-sm font-medium">Limpiar</button>
        <button onclick="saveEmpresa()" class="btn-primary px-4 py-2 rounded-lg text-sm font-medium">Guardar Datos de Empresa</button>
      </div>
    </details>
  </section>

  <!-- FORMULARIO TRABAJADOR -->
  <section class="card p-5 fade-in">
    <div class="section-title mb-4">
      <svg class="w-5 h-5 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M18 9v3m0 0v3m0-3h3m-3 0h-3m-2-5a4 4 0 11-8 0 4 4 0 018 0zM3 20a6 6 0 0112 0v1H3v-1z"/></svg>
      Registrar Nuevo Trabajador
    </div>
    <div class="grid grid-cols-1 md:grid-cols-5 gap-4">
      <div>
        <label class="label">Nombre</label>
        <input id="trab_nombre" class="input" placeholder="Juan"/>
      </div>
      <div>
        <label class="label">Apellido</label>
        <input id="trab_apellido" class="input" placeholder="Pérez"/>
      </div>
      <div>
        <label class="label">Cédula</label>
        <input id="trab_cedula" class="input" placeholder="V-12345678"/>
      </div>
      <div>
        <label class="label">Fecha de Ingreso</label>
        <input id="trab_fecha" type="date" class="input"/>
      </div>
      <div>
        <label class="label">Cargo</label>
        <input id="trab_cargo" class="input" placeholder="Analista, Gerente, Obrero..."/>
      </div>
    </div>
    <div class="flex justify-end gap-2 mt-4">
      <button onclick="resetTrabForm()" class="btn-ghost px-4 py-2 rounded-lg text-sm font-medium">Limpiar</button>
      <button onclick="addTrabajador()" class="btn-primary px-4 py-2 rounded-lg text-sm font-medium" style="background:linear-gradient(135deg,#059669,#047857);">+ Agregar Trabajador</button>
    </div>
  </section>

  <!-- LISTADO -->
  <section class="card p-5 fade-in">
    <div class="flex items-center justify-between mb-4 flex-wrap gap-2">
      <div class="section-title">
        <svg class="w-5 h-5 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/></svg>
        Listado de Trabajadores
        <span id="countBadge" class="badge bg-indigo-100 text-indigo-700">0</span>
      </div>
      <div class="flex items-center gap-2">
        <input id="buscar" oninput="renderTrabajadores()" class="input" style="max-width:220px;" placeholder="🔍 Buscar por nombre, cédula, cargo..."/>
      </div>
    </div>

    <div class="scroll-table rounded-lg border border-slate-200">
      <table class="w-full text-sm">
        <thead>
          <tr class="text-left text-slate-600">
            <th class="px-3 py-2">#</th>
            <th class="px-3 py-2">Nombre</th>
            <th class="px-3 py-2">Apellido</th>
            <th class="px-3 py-2">Cédula</th>
            <th class="px-3 py-2">Fecha Ingreso</th>
            <th class="px-3 py-2">Cargo</th>
            <th class="px-3 py-2 text-right">Acciones</th>
          </tr>
        </thead>
        <tbody id="tbodyTrab" class="divide-y divide-slate-100"></tbody>
      </table>
    </div>
    <div id="emptyState" class="hidden text-center py-10 text-slate-400">
      <svg class="w-12 h-12 mx-auto mb-2 opacity-50" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="1.5"><path stroke-linecap="round" stroke-linejoin="round" d="M17 20h5v-2a4 4 0 00-3-3.87M9 20H4v-2a4 4 0 013-3.87m6-5.13a4 4 0 11-8 0 4 4 0 018 0z"/></svg>
      <p>No hay trabajadores registrados aún.</p>
    </div>
  </section>

  <footer class="text-center text-xs text-slate-400 py-4">
    © 2026 Sistema de Control de Trabajadores · Los datos se guardan localmente en su navegador.
  </footer>
</main>

<!-- MODAL SINCRONIZAR -->
<div id="syncModal" class="hidden fixed inset-0 z-50 modal-bg flex items-center justify-center p-4">
  <div class="card max-w-lg w-full p-6 fade-in">
    <div class="flex items-center justify-between mb-4">
      <h3 class="text-lg font-bold text-slate-800">Sincronizar con Servidor</h3>
      <button onclick="closeSyncModal()" class="text-slate-400 hover:text-slate-600">
        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>
      </button>
    </div>

    <div class="flex gap-2 mb-4 border-b border-slate-200">
      <button onclick="switchSyncTab('config')" id="tabConfig" class="px-3 py-2 text-sm font-medium rounded-t-lg tab-active">Configuración</button>
      <button onclick="switchSyncTab('actions')" id="tabActions" class="px-3 py-2 text-sm font-medium rounded-t-lg text-slate-500">Acciones</button>
    </div>

    <div id="panelConfig" class="space-y-3">
      <div>
        <label class="label">Servidor</label>
        <select id="sync_server" class="input" onchange="updateSyncHints()">
          <option value="googlesheets">Google Sheets</option>
          <option value="googledrive">Google Drive (JSON)</option>
          <option value="custom">Servidor personalizado (API REST)</option>
        </select>
      </div>
      <div>
        <label class="label">API Key / Token</label>
        <input id="sync_apikey" class="input" placeholder="Pega aquí tu API Key o Token"/>
      </div>
      <div>
        <label class="label">ID de hoja / URL / Endpoint</label>
        <input id="sync_target" class="input" placeholder="Ej. ID de la hoja de Google Sheets o URL del endpoint"/>
      </div>
      <p id="sync_hint" class="text-xs text-slate-500 bg-blue-50 border border-blue-100 rounded p-2"></p>
      <button onclick="saveSyncConfig()" class="btn-primary w-full py-2 rounded-lg text-sm font-medium">Guardar Configuración</button>
    </div>

    <div id="panelActions" class="hidden space-y-3">
      <button onclick="syncPush()" class="btn-primary w-full py-3 rounded-lg text-sm font-medium flex items-center justify-center gap-2">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M7 16V4m0 0L3 8m4-4l4 4m6 0v12m0 0l4-4m-4 4l-4-4"/></svg>
        Subir datos al servidor (Push)
      </button>
      <button onclick="syncPull()" class="w-full py-3 rounded-lg text-sm font-medium flex items-center justify-center gap-2 border border-slate-200 hover:bg-slate-50">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M17 4v12m0 0l4-4m-4 4l-4-4M7 20V8m0 0L3 12m4-4l4 4"/></svg>
        Descargar datos del servidor (Pull)
      </button>
      <p class="text-xs text-slate-500 text-center">La sincronización utiliza la configuración guardada.</p>
    </div>
  </div>
</div>

<!-- TOAST -->
<div id="toast" class="toast hidden">
  <div class="bg-slate-800 text-white px-4 py-3 rounded-lg shadow-lg text-sm flex items-center gap-2">
    <span id="toastMsg">Mensaje</span>
  </div>
</div>

<script>
/* ============== DATOS GEOGRÁFICOS COMPLETOS DE VENEZUELA ============== */
const GEO = {
  "Distrito Capital": {
    "Libertador": ["Altagracia","Antímano","Caricuao","Catedral","El Junquito","El Paraíso","El Recreo","San José","San Juan","San Pedro","Santa Rosalía","Santa Teresa","23 de Enero"]
  },
  "Amazonas": {
    "Alto Orinoco": ["La Esmeralda","Huachamacare","Marawaka","Mavaca","Sierra Parima"],
    "Atabapo": ["Fernando Girón Tovar","Guayapo","Isla Ratón","Luis Alberto Gómez","Platanillal","San Fernando de Atabapo","Yapacana"],
    "Atures": ["Alto Ventuari","Bajo Ventuari","Cacurí","Cataniapo","Montaña Verde","Puerto Ayacucho","Samariapo","Ucata"],
    "Autana": ["Alto Guayapo","Bajo Guayapo","Casiquiare","Kunavén","San Carlos del Río Negro","Santa Rosa de Tomo","Yapacana"],
    "Manapiare": ["Alto Ventuari","Bajo Ventuari","Santa María de Erebato","San Juan de Manapiare"],
    "Maroa": ["Castillo","Maroa","Victorino"],
    "Río Negro": ["Casiquiare","Guayapo","San Carlos del Río Negro","San Fernando de Atabapo","Santa Rosa de Tomo","Solano"]
  },
  "Anzoátegui": {
    "Anaco": ["Anaco","San Joaquín"],
    "Aragua de Barcelona": ["Aragua de Barcelona","Cachipo","Lechería","San Miguel"],
    "Diego Bautista Urbaneja": ["Lechería","Morón"],
    "Francisco de Miranda": ["Atapirire","Boca del Pao","El Pao","Pariaguán","Pargüaza"],
    "Francisco del Carmen Carvajal": ["Valle de Guanape","Santa Bárbara"],
    "Fernando de Peñalver": ["Píritu","San Miguel","Sucre"],
    "Guanta": ["Guanta","Chorrerón"],
    "Independencia": ["Mamo","Soledad"],
    "José Gregorio Monagas": ["Mapire","Piar","Santa Clara","Uverito","Zuata"],
    "Juan Antonio Sotillo": ["Barcelona","Bergatín","Caigua","El Carmen","El Pilar","Naricual"],
    "Juan Manuel Cajigal": ["Onoto","San Pablo"],
    "Libertad": ["San Mateo","Santa Ana"],
    "Manuel Ezequiel Bruzual": ["Clarines","Guanape","Sabana de Uchire"],
    "Pedro María Freites": ["Cantaura","Libertador","Pariaguán","Santa Bárbara","Urica"],
    "Píritu": ["Píritu","San Miguel","Sucre"],
    "San Juan de Capistrana": ["Boca de Uchire","Boca de Chávez"],
    "Santa Ana": ["Santa Ana","Pueblo Nuevo"],
    "Simón Bolívar": ["Bergatín","El Carmen","El Pilar","Naricual","Puerto La Cruz","Pozuelos","San Tomás"],
    "Simón Rodríguez": ["El Tigrito","El Tigre"]
  },
  "Apure": {
    "Achaguas": ["Achaguas","Apurito","El Yagual","Guachara","Mucuritas","Queseras del medio"],
    "Biruaca": ["Biruaca"],
    "Muñoz": ["Bruzual","Mantecal","Quintero","Rincón Hondo","San Vicente de la Cabeza"],
    "Páez": ["Aragüita","Arismendi","Guasdualito","Urdaneta"],
    "Pedro Camejo": ["Codazzi","San Juan de Payara"],
    "Rómulo Gallegos": ["Elorza","La Trinidad"],
    "San Fernando": ["El Recreo","Peñón de Matiyures","San Fernando de Apure","San Francisco de Apure"]
  },
  "Aragua": {
    "Bolívar": ["San Mateo"],
    "Camatagua": ["Camatagua","Carmen de Cura"],
    "Francisco Linares Alcántara": ["Santa Rita","Francisco de Miranda","Monseñor Feliciano González"],
    "Girardot": ["Choroni","Las Delicias","Madre María de San José","Joaquín Crespo","José Casanova Godoy","Pedro José Ovalles","El Consejo"],
    "José Ángel Lamas": ["Santa Cruz de Aragua"],
    "José Félix Ribas": ["La Victoria","Castor Nieves Ríos","Las Guacamayas","Pao de Zárate","Zuata"],
    "José Rafael Revenga": ["El Consejo"],
    "Libertador": ["Palo Negro","San Martín de Porres"],
    "Mario Briceño Iragorry": ["El Limón","Caña de Azúcar"],
    "Ocumare de la Costa de Oro": ["Ocumare de la Costa","Cata","Cuyagua","La Ciénaga","La Lima","Las Bocas","Ocumare de la Costa"],
    "San Casimiro": ["San Casimiro","Güiripa","Ollas de Caramacate","Valle Morín"],
    "San Sebastián": ["San Sebastián"],
    "Santiago Mariño": ["Turmero","Arevalo","Choroni","Samán de Güere","Alfredo Pacheco Miranda","Valentín Díaz"],
    "Santos Michelena": ["Las Tejerías","Tiara"],
    "Sucre": ["Cagua","Bella Vista"],
    "Tovar": ["Colonia Tovar"],
    "Urdaneta": ["Barbacoas","Las Peñitas","San Francisco de Asís","Valle de la Pascua"],
    "Zamora": ["Villa de Cura","Magdaleno","San Francisco de Asís","Valle de la Pascua","Guayabal"]
  },
  "Barinas": {
    "Alberto Arvelo Torrealba": ["Sabaneta","Rodríguez Domínguez"],
    "Andrés Eloy Blanco": ["El Cantón","La Luz","Santa Bárbara","San Camilo"],
    "Antonio José de Sucre": ["Ciudad de Nutrias","El Regalo","Guasimitos","Puerto Nutrias"],
    "Arismendi": ["Arismendi","Guadarrama","La Unión","San Antonio"],
    "Barinas": ["Alto Barinas","Barinas","Barinitas","Calderas","El Carmen","Humberto Janer","Los Corobos","Rómulo Betancourt","San Silvestre","Santa Inés","Santa Lucía","Torunos","Unare"],
    "Bolívar": ["Barinitas","Altamira de Cáceres","El Calderón","Santa Bárbara"],
    "Cruz Paredes": ["Barrancas","Barinitas","Obispos"],
    "Ezequiel Zamora": ["Santa Bárbara","Pedro Briceño Méndez","Ramón Ignacio Méndez","José Ignacio del Pumar"],
    "Obispos": ["Obispos","El Real","Guasimitos","La Esmeralda"],
    "Pedraza": ["Ciudad Bolivia","Ignacio Briceño","José Félix Ribas","Páez"],
    "Rojas": ["Libertad","Pablo VI","Dolores"],
    "Sosa": ["Ciudad de Nutrias","El Regalo","Guasimitos","Puerto Nutrias"],
    "Sucre": ["Ciudad de Nutrias","El Regalo","Guasimitos","Puerto Nutrias"],
    "Tórbes": ["Santa Lucía","San Rafael de Canaguá","Urimare"],
    "Zamora": ["Santa Bárbara","José Ignacio del Pumar","Pedro Briceño Méndez","Ramón Ignacio Méndez"]
  },
  "Bolívar": {
    "Angostura": ["Ciudad Piar","La Urbana","Pasa Tiempo","Salóm","San Francisco de Caura","Upata"],
    "Caroní": ["Cachamay","Chirica","Dalla Costa","Once de Marzo","Pozo Verde","San Félix","Unare","Vista al Sol"],
    "Cedeño": ["Altagracia","Ascensión Farreras","Guaniamo","La Urbana","Pijiguaos","Puerto Cedeño","Salóm"],
    "El Callao": ["El Callao"],
    "Gran Sabana": ["Ikabarú","Santa Elena de Uairén","Uairén"],
    "Heres": ["Agua Salada","Catedral","La Sabanita","Marhuanta","Vista al Sol","Zea"],
    "Piar": ["Andrés Eloy Blanco","Pedro Cova","Upata"],
    "Roscio": ["Guasipati","Salóm","San Isidro"],
    "Sifontes": ["Dalla Costa","San Isidro","Santa Elena de Uairén"],
    "Sucre": ["Aripao","Guarataro","Las Majadas","Moitaco"],
    "Padre Pedro Chien": ["El Palmar"]
  },
  "Carabobo": {
    "Bejuma": ["Bejuma","Canoabo","Simón Bolívar"],
    "Carlos Arvelo": ["Güigüe","Bélén","Carabobo","Central","Tacarigua"],
    "Diego Ibarra": ["Aguas Calientes","Mariara"],
    "Guacara": ["Ciudad Alianza","Guacara","Yagua"],
    "Juan José Mora": ["Morón","Urama"],
    "Libertador": ["Campo de Carabobo","Tocuyito"],
    "Los Guayos": ["Los Guayos"],
    "Miranda": ["Miranda","Urama"],
    "Montalbán": ["Montalbán"],
    "Naguanagua": ["Naguanagua"],
    "Puerto Cabello": ["Bartolomé Salom","Democracia","Fraternidad","Goaigoaza","Juan José Flores","Unión","Borburata","Patanemo"],
    "San Diego": ["San Diego"],
    "San Joaquín": ["San Joaquín"],
    "Valencia": ["Candelaria","Catedral","El Socorro","Miguel Peña","Rafael Urdaneta","San Blas","San José","Santa Rosa"]
  },
  "Cojedes": {
    "Anzoátegui": ["Cojedes","El Amparo"],
    "Ezequiel Zamora": ["San Carlos","El Amparo","Sucre"],
    "Girardot": ["El Baúl","Sucre"],
    "Lima Blanco": ["La Aguadita","Macapo"],
    "Pao de San Juan Bautista": ["El Pao"],
    "Ricaurte": ["El Baúl","Libertad"],
    "Rómulo Gallegos": ["Las Vegas","Libertad de Cojedes"],
    "Tinaquillo": ["Tinaquillo"],
    "Tucupido": ["Tucupido"]
  },
  "Delta Amacuro": {
    "Antonio Díaz": ["Curiapo","El Rabo","Juan Millán","Manoa","Moruca","San Francisco","San José"],
    "Casacoima": ["Arasca","Imataca","Juan Bautista Arismendi","Manuel Renaud","Padre Barral","San Clemente"],
    "Pedernales": ["Pedernales","Luis Beltrán Prieto Figueroa"],
    "Tucupita": ["Almirante Luis Brión","Aniceto Arvelo","José Vidal Marcano","Juan Bautista Arismendi","San José","San Rafael","Virgen del Valle"]
  },
  "Falcón": {
    "Acosta": ["San Juan de los Cayos","Capadare","La Pastora"],
    "Bolívar": ["San Luis","Pecaya","Guabina"],
    "Buchivacoa": ["Capatarida","Bariro","Borojó","Guajiro","Seque","Zazárida","Valle de Eroa"],
    "Cacique Manaure": ["Yaracal"],
    "Carirubana": ["Carirubana","Norte","Piedras Negras","Punto Fijo"],
    "Colina": ["La Vela de Coro","El Hato","El Vínculo"],
    "Daburo": ["Daburo","Canelón","Guayamuri","Agua Larga"],
    "Democracia": ["Pedregal","Agua Clara","Acurigua","El Barrio","El Cardón","El Hato","El Vínculo"],
    "Falcón": ["Pueblo Nuevo","El Charal","Las Cumaraguas","Santa Ana","El Hato","El Vínculo"],
    "Federación": ["Churuguara","Agua Larga","El Paují","Independencia","Maparari"],
    "Jacura": ["Jacura","Agua Linda","Boca de Tocuyo","El Vínculo"],
    "Los Taques": ["Santa Cruz de los Taques","Los Taques","El Vínculo"],
    "Mauroa": ["Mene de Mauroa","Chichiriviche","Santa Cruz de Bucaral","Boca de Tocuyo"],
    "Miranda": ["Santa Ana de Coro","El Hato","El Vínculo","Guzmán Blanco","San Gabriel","San Juan y Carrizal","El Vínculo"],
    "Monseñor Iturriza": ["Chichiriviche","Boca de Tocuyo","El Vínculo"],
    "Palcasúa": ["Palma Sola","El Vínculo"],
    "Petit": ["Cabure","Curimagua","El Vínculo"],
    "Píritu": ["Píritu","San José","El Vínculo"],
    "San Francisco": ["Mirimire","El Vínculo"],
    "Silva": ["Tucacas","Boca de Aroa","El Vínculo"],
    "Sucre": ["La Cruz","El Vínculo"],
    "Tocópero": ["Tocópero","El Vínculo"],
    "Unión": ["Santa Cruz de Bucaral","El Vínculo"],
    "Urumaco": ["Urumaco","El Vínculo"],
    "Zamora": ["Puerto Cumarebo","Cumaraguas","La Ciénaga","Pueblo Nuevo","El Vínculo"]
  },
  "Guárico": {
    "Camaguán": ["Camaguán","Puerto Miranda","Uverito"],
    "Chaguaramas": ["Chaguaramas"],
    "El Socorro": ["El Socorro"],
    "Francisco de Miranda": ["Calabozo","Tucupido","El Rastro","Guamal"],
    "José Félix Ribas": ["Valle de la Pascua","Espino"],
    "José Tadeo Monagas": ["Altagracia de Orituco","San Rafael de Orituco","Lezama","Cantagallo","Paso Real","Libertad de Orituco"],
    "Juan Germán Roscio": ["San Juan de los Morros","Cantagallo","Parapara","Cardonal"],
    "Julián Mellado": ["San Juan de los Morros","Cantagallo","Parapara"],
    "Las Mercedes": ["Las Mercedes","Cabruta","Santa Rita de Manapire"],
    "Leonardo Infante": ["Valle de la Pascua","Espino"],
    "Manuel Ezequiel Bruzual": ["Paso Real de Macaira","Soto","Uverito"],
    "Pedro Zaraza": ["Zaraza","San José de Unare","Guamal","El Sombrero"],
    "Ortíz": ["Ortiz","San José de Tiznados","San Francisco de Tiznados","San Lorenzo de Tiznados"],
    "San Gerónimo de Guayabal": ["Guayabal","Cabruta"],
    "San José de Guaribe": ["San José de Guaribe"],
    "Santa María de Ipire": ["Santa María de Ipire","Altamira"]
  },
  "Lara": {
    "Crespo": ["Aguedo Felipe Alvarado","Buena Vista","Catedral","El Cují","Juan de Villegas"],
    "Iribarren": ["Catedral","Concepción","El Cují","Juárez","Juan de Villegas","Santa Rosa","Tamaca","Unión"],
    "Jiménez": ["Juan Bautista Rodríguez","Cuara","Diego de Lozada","Paraíso de San José","San Miguel","Tintorero","José Gregorio Bastidas","Río Claro"],
    "Morán": ["Anzoátegui","Guarico","Hilario Luna y Luna","Las Mercedes","Manuel Morillo","Sarare","Tranquera de Lara","Unión"],
    "Palavecino": ["Cabudare","José Gregorio Bastidas","Aguedo Felipe Alvarado"],
    "Simón Planas": ["Sarare","Buría","Gustavo Vega"],
    "Torres": ["Tranquera de Lara","Aguedo Felipe Alvarado","Anzoátegui","Guarico","Hilario Luna y Luna","Las Mercedes","Manuel Morillo","Sarare","Unión"],
    "Urdaneta": ["Siquisique","Moroturo","San Miguel","Xaguas"]
  },
  "Mérida": {
    "Alberto Adriani": ["El Vigía","Caño El Tigre","Guayabales","La Palizada","Mesa Bolívar","Mesa de Las Palmas","Pueblo Nuevo de Piñango"],
    "Andrés Bello": ["La Azulita","San Antonio de las Lomas"],
    "Antonio Pinto Salinas": ["Mucuchíes","Cacute","La Mesa","San Rafael"],
    "Aricagua": ["Aricagua","San Antonio"],
    "Arzobispo Chacón": ["Canaguá","Araguaney","El Milagro","Guaimaral"],
    "Campo Elías": ["Ejido","Aricagua","La Trinidad","San José de las Mesas","Timotes"],
    "Caracciolo Parra Olmedo": ["Tucaní","Florencio Ramírez","Pueblo Llano"],
    "Cardenal Quintero": ["Santo Domingo","Las Piedras","Pueblo Nuevo","Zea"],
    "Guaraque": ["Guaraque","Mesa de Quintero","Río Bonito"],
    "Julio César Salas": ["Arapuey","Palmira"],
    "Justo Briceño": ["Torondoy","San Cristóbal de Torondoy"],
    "Libertador": ["Mérida","Arias","Caracciolo Parra Pérez","Domingo Peña","El Llano","Gonzalo Picón Febres","Jacinto Plaza","Juan Rodríguez Suárez","Lasso de la Vega","Luis Quintano","Mariano Picón Salas","Milla","Osuna Rodríguez","Sagrario","El Faro"],
    "Miranda": ["Timotes","Andrés Eloy Blanco","Los Nevados","Santa Cruz de Mora"],
    "Obispo Ramos de Lora": ["Santa Elena de Arenales","Eloy Paredes","San Rafael de Alcázar"],
    "Padre Noguera": ["Santa María de Caparo"],
    "Pueblo Llano": ["Pueblo Llano"],
    "Rangel": ["Mucuchíes","Cacute","La Mesa","San Rafael"],
    "Rivas Dávila": ["Bailadores","Gerónimo Maldonado"],
    "Santos Marquina": ["Tabay"],
    "Sucre": ["Lagunillas","Chiguará","Estánquez","La Trampa","Pueblo Nuevo del Sur","San Juan"],
    "Tovar": ["Tovar","El Amparo","El Atuto","La Alpujarra","Pueblo Nuevo de Tovar"],
    "Tulio Febres Cordero": ["Nueva Bolivia","Independencia","La Azulita","Santa Apolonia"],
    "Zea": ["Zea","Caño El Tigre","El Milagro","Guaimaral"]
  },
  "Miranda": {
    "Acevedo": ["Aragüita","Arévalo González","Capaya","Caucagua","El Café","Marizapa","Panaquire","Ribas","San José de Barlovento"],
    "Andrés Bello": ["San José de Barlovento","Caucagüita","El Cafetal"],
    "Araure": ["Higuerote","Curiepe","Tacarigua"],
    "Baruta": ["Baruta","El Cafetal","Las Minas de Baruta"],
    "Brión": ["Higuerote","Curiepe","Tacarigua"],
    "Buroz": ["Mamporal","Cumbo","Las Brisas","San José de Barlovento"],
    "Carrizal": ["Carrizal"],
    "Chacao": ["Chacao"],
    "Cristóbal Rojas": ["Charallave","Las Brisas"],
    "El Hatillo": ["El Hatillo"],
    "Guaicaipuro": ["Los Teques","Altagracia de la Montaña","Cecilio Acosta","El Jarillo","San Pedro","Tácata","Paracotos"],
    "Independencia": ["Santa Teresa del Tuy","El Cartanal"],
    "Lander": ["Ocumare del Tuy","La Democracia","Santa Bárbara"],
    "Los Salias": ["San Antonio de Los Altos"],
    "Páez": ["Río Chico","El Guapo","Tacarigua","Paparo","San José de Barlovento"],
    "Paz Castillo": ["Santa Lucía"],
    "Pedro Gual": ["Cúa","Nueva Cúa"],
    "Plaza": ["Guarenas"],
    "Simón Bolívar": ["San José de Barlovento","Caucagüita","El Cafetal"],
    "Sucre": ["Petare","Caucagüita","Leoncio Martínez","Filas de Mariche","San José de Galipán"],
    "Urdaneta": ["Cúa","Nueva Cúa"],
    "Zamora": ["Guatire","Bolívar"]
  },
  "Monagas": {
    "Acosta": ["San Antonio de Capayacuar","Guayusa"],
    "Aguasay": ["Aguasay"],
    "Bolívar": ["Caripito","El Furrial","La Toscana"],
    "Caripe": ["Caripe","El Guácharo","La Guanota","Sabana de Piedra"],
    "Cedeño": ["Caicara de Maturín","Areo","San Félix de Cantalicio","Vieja Caracas"],
    "Ezequiel Zamora": ["Punta de Mata","Chaguaramas","Las Alhuacas","Tabasca"],
    "Libertador": ["Temblador","Cañafístola","Chaguaramas"],
    "Maturín": ["Alto de la Cruz","Aparicio","Aragua de Maturín","Barrancas","Cachipo","El Corozo","El Furrial","Jusepín","La Pica","Las Cocuizas","Maturín","San Simón","El Tesoro"],
    "Piar": ["Aragua de Maturín","Chaguaramal","El Pinto","Guataro","La Toscana","Los Barrancos de Fajardo"],
    "Punceres": ["Quiriquire","Cachipo"],
    "Santa Bárbara": ["Santa Bárbara"],
    "Sotillo": ["Barrancas del Orinoco","Barrancas de Maturín","Los Barrancos de Fajardo"],
    "Uracoa": ["Uracoa","Moriche"]
  },
  "Nueva Esparta": {
    "Antolín del Campo": ["La Plaza de Paraguachi","La Asunción","Juangriego","Pampatar"],
    "Arismendi": ["La Asunción","Arismendi","Juan Griego","Pedro González"],
    "Díaz": ["San Juan Bautista","Zabala"],
    "García": ["El Valle del Espíritu Santo","La Guardia","Güinima","Paraíso"],
    "Gómez": ["Maneiro","Aguirre","Bolívar","Guevara","Sucre"],
    "Maneiro": ["Maneiro","Aguiar","Bolívar","Guevara","Sucre"],
    "Marcano": ["Porlamar"],
    "Mariño": ["Porlamar","Bella Vista"],
    "Península de Macanao": ["Boca de Río","San Francisco de Macanao"],
    "Tubores": ["Punta de Piedras","Los Barales","San Pedro de Coche","Güinima"],
    "Villalba": ["San Pedro de Coche","Vicente Fuentes","Bichar"]
  },
  "Portuguesa": {
    "Agua Blanca": ["Agua Blanca"],
    "Araure": ["Araure","Río Acarigua"],
    "Esteller": ["Píritu","Uveral"],
    "Guanare": ["Guanare","Cordové","San José de Saguas","San Juan de Guanaguanare","Virgen de la Coromoto"],
    "José Vicente de Unda": ["Paraíso de Chabasquén","Peña Blanca","Chabasquén"],
    "La Estrella": ["La Estrella","Chaguaramal"],
    "Monseñor José Vicenti de Unda": ["Paraíso de Chabasquén","Peña Blanca"],
    "Monseñor José Vicente de Unda": ["Paraíso de Chabasquén"],
    "Ospino": ["Ospino","Aparición","La Estación"],
    "Páez": ["Acarigua","Payara","Pimpinela","Ramón Peraza"],
    "Papón": ["El Playón","Florentino"],
    "Píritu": ["Píritu","Uveral"],
    "San Genaro de Boconoíto": ["Boconoíto","El Playón","Florentino"],
    "Santa Rosalía": ["Santa Rosalía"],
    "Sucre": ["Biscucuy","Concepción","Guayabita","Penitente","San Rafael de Onoto","Santa Fe"],
    "Turén": ["Villa Bruzual","Chivacoa","El Playón","Santa Rita"]
  },
  "Sucre": {
    "Andrés Eloy Blanco": ["Casanay","San Antonio de Guacharo"],
    "Andrés Mata": ["Cumanacoa","Arenas","Cocollar","San Lorenzo","San Pedro"],
    "Arismendi": ["Río Caribe","El Morro de Puerto Santo","Guaraúnos","Maracapana","Puerto Santo"],
    "Benítez": ["El Pilar","El Rincón","General Francisco Antonio Váquez","Guaraúnos","Maracapana","Río Caribe"],
    "Bermúdez": ["Carúpano","Catuaro","Rendón","San Cruz","Santa Catalina"],
    "Bolívar": ["Marigüitar","Cariaco","Catuaro","Rendón","San Cruz","Santa María","Villa Frontado"],
    "Cajigal": ["Yaguaraparo","El Paujil","Libertad"],
    "Cruz Salmerón Acosta": ["Araya","Chacopata","Manicuare"],
    "Libertad": ["Tunapuy","Campo Elías","El Altico","Guacamaya","Paraguachí"],
    "Mariño": ["Irapa","Campo Florido","Dumar","Los Moruyos","San Antonio de Irapa"],
    "Mejía": ["San Antonio del Golfo"],
    "Montes": ["Cumaná","Ayacucho","San Juan","Valentín Valiente"],
    "Ribero": ["Cariaco","Cumanacoa","Río Caribe","San Juan","Santa María"],
    "Sucre": ["Cumaná","Altagracia","Ayacucho","San Juan","Valentín Valiente","Villa Frontado"]
  },
  "Táchira": {
    "Andrés Bello": ["Cordero","El Cobre","Las Mesas","San Antonio"],
    "Antonio Rómulo Costa": ["Las Mesas","San Antonio"],
    "Ayacucho": ["San Juan de Colón","Rivas Berti","San Pedro del Río"],
    "Bolívar": ["San Antonio","Palotal","Juan Vicente Gómez","Isaías Medina Angarita"],
    "Cárdenas": ["Táriba","Aguas Calientes","San Rafael"],
    "Córdoba": ["Santa Ana del Táchira","Palmira","Ricaurte"],
    "Fernández Feo": ["San Rafael del Piñal","Quinimarí","Bramón","Doradas","El Piñal","La Tendida"],
    "Francisco de Miranda": ["San José de Bolívar","Las Delicias","El Cobre"],
    "García de Hevia": ["La Fría","Palmira","Ricaurte","Río de Oro"],
    "Guásimos": ["Palmira","La Grita","Ricaurte"],
    "Independencia": ["Independencia","El Guayabo","Cocorote","Río de Oro"],
    "Jáuregui": ["La Grita","Palmira","Ricaurte"],
    "José María Vargas": ["El Cobre","Las Mesas"],
    "Junín": ["Rubio","La Petrolea","Rivas Berti","San Pedro del Río"],
    "Libertad": ["Capacho Nuevo","Capacho Viejo","Cipriano Castro","Manuel Felipe Rugeles"],
    "Libertador": ["Abejales","Borota","La Concordia","La Pedrera","Lobatera","San Rafael"],
    "Libertador (Táchira)": ["Abejales","Borota","La Concordia","La Pedrera","Lobatera","San Rafael"],
    "Michelena": ["Michelena"],
    "Panamericano": ["Coloncito","La Palmita"],
    "Pedro María Ureña": ["Ureña","Nueva Arcadia"],
    "Rafael Urdaneta": ["Delicias","San Rafael"],
    "Samuel Darío Maldonado": ["La Tendida","Río de Oro"],
    "San Cristóbal": ["San Cristóbal","Altamira","Emeterio Ochoa","San Juan Eudes"],
    "San Judas Tadeo": ["Umuquena"],
    "Seboruco": ["Seboruco"],
    "Simón Rodríguez": ["San Simón"],
    "Sucre": ["Quinimarí","Bramón"],
    "Torbes": ["San Josecito","San Pedro del Río","Bobures"],
    "Uribante": ["Pregonero","Cardenal Quintero","Juan Vicente Gómez"]
  },
  "Trujillo": {
    "Andrés Bello": ["Santa Isabel","José Gregorio Hernández"],
    "Boconó": ["Boconó","Ayacucho","Burbusay","El Carmen","General Ribas","Guaramacal","Monseñor Jáuregui","Rafael Rangel","San Miguel","San José"],
    "Bolívar": ["Sabana Grande","Cheregüé","Reinaldo Cesar"],
    "Candelaria": ["Chejendé","Arnoldo Gabaldón","Bolivia","Carrillo","Cegarra","Guamontey","José Felipe Márquez Cañizalez","Manuel Salvador Ulloa","José Gregorio Hernández"],
    "Carache": ["Carache","La Concepción","Cuicas","Panamericana","Santa Cruz"],
    "Escuque": ["Escuque","La Ceiba","Santa Ana","José Felipe Márquez Cañizalez"],
    "José Felipe Márquez Cañizalez": ["El Paradero","José Felipe Márquez Cañizalez"],
    "Juan Vicente Campos Elías": ["Campo Elías","Arnoldo Gabaldón","Santa Apolonia","Timotes"],
    "La Ceiba": ["La Ceiba","Santa Apolonia","Profría"],
    "Miranda": ["El Dividive","Agua Santa","Valerita","Cuba"],
    "Monte Carmelo": ["Monte Carmelo","Altamira","Guarorí"],
    "Motatán": ["Motatán","El Baño","Jalisco"],
    "Pampán": ["Pampán","Flor de Patria","La Paz","Santa Ana"],
    "Pampanito": ["Pampanito","La Concepción","Pampanito II"],
    "Rafael Rangel": ["Betijoque","La Pueblita","Los Cedros","José Gregorio Hernández"],
    "San Rafael de Carvajal": ["Carvajal","Campo Alegre","José Leonardo Suárez"],
    "Sucre": ["Sabana Grande","Cheregüé","Reinaldo Cesar"],
    "Trujillo": ["Trujillo","Andrés Linares","Chiquinquirá","Cruz Carrillo","Cristóbal Mendoza","El Recreo","Guayabito","La Mesa de Esnujaque","Matriz","Tostós"],
    "Urdaneta": ["La Ceiba","Santa Apolonia","Profría","Sabana Grande"],
    "Valera": ["Valera","El Recreo","José Felipe Márquez Cañizalez","La Mesa de Esnujaque","Juan Ignacio Montilla"]
  },
  "La Guaira": {
    "Caraballeda": ["Caraballeda"],
    "Caruao": ["Caruao","Camuri","Chichiriviche de la Costa","El Carmen","El Joyero","Huaira","Las Tinas","Nueva Granada","Urimare"],
    "Catia La Mar": ["Catia La Mar"],
    "El Junko": ["El Junko"],
    "La Guaira": ["La Guaira","Caraballeda","Catia La Mar"],
    "Macuto": ["Macuto"],
    "Maiquetía": ["Maiquetía"],
    "Naiguatá": ["Naiguatá"],
    "Urimare": ["Urimare","Camuri","Chichiriviche de la Costa"],
    "Carlos Soublette": ["Carlos Soublette"],
    "José María España": ["José María España"]
  },
  "Yaracuy": {
    "Arístides Bastidas": ["San Pablo"],
    "Bolívar": ["Aroa","Chivacoa"],
    "Bruzual": ["Chivacoa","Campo Elías"],
    "Cocorote": ["Cocorote"],
    "Independencia": ["Independencia","El Guayabo","Cocorote"],
    "José Antonio Páez": ["Sabana de Parra"],
    "La Trinidad": ["La Trinidad"],
    "Manuel Monge": ["Yumare"],
    "Nirgua": ["Nirgua","Salom","Temerla","Urama"],
    "Peña": ["Yaritagua","Urachiche"],
    "San Felipe": ["San Felipe","Albarico","Aroa","Bariquimeto"],
    "Sucre": ["Guama"],
    "Urachiche": ["Urachiche"],
    "Veroes": ["Farriar","El Guayabo"]
  },
  "Zulia": {
    "Almirante Padilla": ["El Toro","Isla de San Carlos","San Timoteo"],
    "Baralt": ["San Timoteo","El Guabo","Pueblo Nuevo"],
    "Cabimas": ["Ambrosio","Carmen Herrera","Germán Ríos Linares","Jorge Hernández","La Rosa","Rómulo Betancourt"],
    "Catatumbo": ["Encontrados","Río Negro","Udón Pérez"],
    "Colón": ["San Carlos del Zulia","Santa Bárbara","Urribarrí"],
    "Francisco Javier Pulgar": ["Pueblo Nuevo","El Chivo","Río Negro","Simoncito"],
    "Jesús Enrique Lossada": ["La Concepción","José Ramón Yépez","Mariano Parra León","San José de Perijá"],
    "Jesús María Semprún": ["Casigua-El Cubo","El Carmelo","Gibraltar","Santa Rita"],
    "La Cañada de Urdaneta": ["Concepción","Donaldo García","El Rosario","Olasini","Santa Rita"],
    "Lagunillas": ["Ciudad Ojeda","Alonso de Ojeda","Libertad","Venezuela","Eleazar López Contreras"],
    "Machiques de Perijá": ["Machiques","Bartolomé de las Casas","Bobiore","El Rosario","Libertad","Río Negro"],
    "Mara": ["San Rafael de El Moján","La Sierrita","Las Parcelas","Luis de Vicente","Ricaurte","Tamare"],
    "Maracaibo": ["Antonio Díaz","Chiquinquirá","Francisco Eugenio Bustamante","Idelfonzo Vásquez","Juana de Ávila","Olegario Villalobos","Raúl Leoni","San Isidro"],
    "Miranda": ["Los Puertos de Altagracia","Altagracia","Ana María Campos","San Antonio","San José de Perijá","San Timoteo"],
    "Páez": ["Sinamaica","Alta Guajira","Elías Sánchez Rubio","Guajira"],
    "Rosario de Perijá": ["La Villa del Rosario","Donaldo García","El Rosario","Olasini","San José de Perijá"],
    "San Francisco": ["San Francisco","El Bajo"],
    "Santa Rita": ["Santa Rita","El Mene","José Cenobio Urribarrí","Pedro Lucas Urribarrí"],
    "Simón Bolívar": ["Tía Juana","La Rosa","Pueblo Nuevo"],
    "Sucre": ["Bobures","Gibraltar","Heras","Monseñor Arturo Álvarez","Rómulo Gallegos","San Timoteo"],
    "Valmore Rodríguez": ["La Victoria","Rafael Urdaneta","Rómulo Betancourt"]
  },
  "Dependencias Federales": {
    "Dependencias Federales": ["Los Roques","La Orchila","Los Testigos","Las Aves","Chaguaramas"]
  }
};

/* ============== ESTADO ============== */
let trabajadores = JSON.parse(localStorage.getItem('trabajadores') || '[]');
let empresa = JSON.parse(localStorage.getItem('empresa') || 'null');
let syncCfg = JSON.parse(localStorage.getItem('syncCfg') || 'null');
let editingId = null;

/* ============== INIT ============== */
function init() {
  const sel = document.getElementById('emp_estado');
  Object.keys(GEO).sort((a,b)=>a.localeCompare(b,'es')).forEach(e => {
    const o = document.createElement('option');
    o.value = e; o.textContent = e;
    sel.appendChild(o);
  });
  if (empresa) loadEmpresa();
  renderTrabajadores();
}

/* ============== EMPRESA ============== */
function loadMunicipios() {
  const estado = document.getElementById('emp_estado').value;
  const mSel = document.getElementById('emp_municipio');
  const pSel = document.getElementById('emp_parroquia');
  mSel.innerHTML = '<option value="">Seleccione un municipio...</option>';
  pSel.innerHTML = '<option value="">Seleccione primero un municipio...</option>';
  pSel.disabled = true;
  if (!estado) { mSel.disabled = true; return; }
  mSel.disabled = false;
  Object.keys(GEO[estado]).sort((a,b)=>a.localeCompare(b,'es')).forEach(m => {
    const o = document.createElement('option');
    o.value = m; o.textContent = m;
    mSel.appendChild(o);
  });
}

function loadParroquias() {
  const estado = document.getElementById('emp_estado').value;
  const municipio = document.getElementById('emp_municipio').value;
  const pSel = document.getElementById('emp_parroquia');
  pSel.innerHTML = '<option value="">Seleccione una parroquia...</option>';
  if (!municipio) { pSel.disabled = true; return; }
  pSel.disabled = false;
  GEO[estado][municipio].forEach(p => {
    const o = document.createElement('option');
    o.value = p; o.textContent = p;
    pSel.appendChild(o);
  });
}

function toggleSucursal() {
  const on = document.getElementById('emp_sucursal').checked;
  document.getElementById('sucursal_info').classList.toggle('hidden', !on);
  if (!on) {
    document.getElementById('emp_centrocomercial').checked = false;
    toggleCentroComercial();
  }
}
function toggleCentroComercial() {
  const on = document.getElementById('emp_centrocomercial').checked;
  document.getElementById('cc_info').classList.toggle('hidden', !on);
}

function saveEmpresa() {
  empresa = {
    nombre: val('emp_nombre'),
    rif: val('emp_rif'),
    centro: val('emp_centro'),
    estado: val('emp_estado'),
    municipio: val('emp_municipio'),
    parroquia: val('emp_parroquia'),
    calle: val('emp_calle'),
    sector: val('emp_sector'),
    sucursal: document.getElementById('emp_sucursal').checked,
    sucursal_ubicacion: val('emp_sucursal_ubicacion'),
    enCentroComercial: document.getElementById('emp_centrocomercial').checked,
    cc_nombre: val('emp_cc_nombre'),
    cc_ciudad: val('emp_cc_ciudad'),
    actualizado: new Date().toISOString()
  };
  localStorage.setItem('empresa', JSON.stringify(empresa));
  toast('✅ Datos de empresa guardados correctamente');
}

function loadEmpresa() {
  if (!empresa) return;
  setVal('emp_nombre', empresa.nombre);
  setVal('emp_rif', empresa.rif);
  setVal('emp_centro', empresa.centro);
  setVal('emp_estado', empresa.estado);
  loadMunicipios();
  setVal('emp_municipio', empresa.municipio);
  loadParroquias();
  setVal('emp_parroquia', empresa.parroquia);
  setVal('emp_calle', empresa.calle);
  setVal('emp_sector', empresa.sector);
  document.getElementById('emp_sucursal').checked = !!empresa.sucursal;
  toggleSucursal();
  setVal('emp_sucursal_ubicacion', empresa.sucursal_ubicacion);
  document.getElementById('emp_centrocomercial').checked = !!empresa.enCentroComercial;
  toggleCentroComercial();
  setVal('emp_cc_nombre', empresa.cc_nombre);
  setVal('emp_cc_ciudad', empresa.cc_ciudad);
}

function clearEmpresa() {
  if (!confirm('¿Limpiar los datos de empresa?')) return;
  ['emp_nombre','emp_rif','emp_centro','emp_estado','emp_municipio','emp_parroquia','emp_calle','emp_sector','emp_sucursal_ubicacion','emp_cc_nombre','emp_cc_ciudad'].forEach(id => setVal(id,''));
  document.getElementById('emp_sucursal').checked = false;
  document.getElementById('emp_centrocomercial').checked = false;
  toggleSucursal();
  toggleCentroComercial();
  localStorage.removeItem('empresa');
  empresa = null;
  toast('Datos de empresa eliminados');
}

/* ============== TRABAJADORES ============== */
function addTrabajador() {
  const nombre = val('trab_nombre').trim();
  const apellido = val('trab_apellido').trim();
  const cedula = val('trab_cedula').trim();
  const fecha = val('trab_fecha');
  const cargo = val('trab_cargo').trim();
  if (!nombre || !apellido || !cedula || !fecha || !cargo) {
    toast('⚠️ Complete todos los campos del trabajador');
    return;
  }
  if (editingId) {
    const i = trabajadores.findIndex(t => t.id === editingId);
    trabajadores[i] = { ...trabajadores[i], nombre, apellido, cedula, fecha, cargo };
    editingId = null;
    toast('✅ Trabajador actualizado');
  } else {
    if (trabajadores.some(t => t.cedula === cedula)) {
      toast('⚠️ Ya existe un trabajador con esa cédula');
      return;
    }
    trabajadores.push({ id: Date.now(), nombre, apellido, cedula, fecha, cargo, creado: new Date().toISOString() });
    toast('✅ Trabajador agregado');
  }
  localStorage.setItem('trabajadores', JSON.stringify(trabajadores));
  resetTrabForm();
  renderTrabajadores();
}

function resetTrabForm() {
  ['trab_nombre','trab_apellido','trab_cedula','trab_fecha','trab_cargo'].forEach(id => setVal(id,''));
  editingId = null;
}

function editTrabajador(id) {
  const t = trabajadores.find(x => x.id === id);
  if (!t) return;
  setVal('trab_nombre', t.nombre);
  setVal('trab_apellido', t.apellido);
  setVal('trab_cedula', t.cedula);
  setVal('trab_fecha', t.fecha);
  setVal('trab_cargo', t.cargo);
  editingId = id;
  window.scrollTo({ top: document.querySelector('section:nth-of-type(2)').offsetTop - 80, behavior: 'smooth' });
  toast('Editando trabajador. Modifique y guarde.');
}

function deleteTrabajador(id) {
  if (!confirm('¿Eliminar este trabajador?')) return;
  trabajadores = trabajadores.filter(t => t.id !== id);
  localStorage.setItem('trabajadores', JSON.stringify(trabajadores));
  renderTrabajadores();
  toast('🗑️ Trabajador eliminado');
}

function renderTrabajadores() {
  const tbody = document.getElementById('tbodyTrab');
  const q = val('buscar').toLowerCase();
  const filtered = trabajadores.filter(t =>
    !q || (t.nombre+' '+t.apellido+' '+t.cedula+' '+t.cargo).toLowerCase().includes(q)
  );
  document.getElementById('countBadge').textContent = trabajadores.length;
  document.getElementById('emptyState').classList.toggle('hidden', filtered.length > 0);
  tbody.innerHTML = filtered.map((t,i) => `
    <tr>
      <td class="px-3 py-2 text-slate-400">${i+1}</td>
      <td class="px-3 py-2 font-medium text-slate-800">${escapeHtml(t.nombre)}</td>
      <td class="px-3 py-2 text-slate-700">${escapeHtml(t.apellido)}</td>
      <td class="px-3 py-2 text-slate-700 font-mono text-xs">${escapeHtml(t.cedula)}</td>
      <td class="px-3 py-2 text-slate-700">${formatDate(t.fecha)}</td>
      <td class="px-3 py-2"><span class="badge bg-blue-50 text-blue-700">${escapeHtml(t.cargo)}</span></td>
      <td class="px-3 py-2 text-right">
        <button onclick="editTrabajador(${t.id})" class="text-blue-600 hover:text-blue-800 p-1" title="Editar">
          <svg class="w-4 h-4 inline" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>
        </button>
        <button onclick="deleteTrabajador(${t.id})" class="text-red-500 hover:text-red-700 p-1 ml-1" title="Eliminar">
          <svg class="w-4 h-4 inline" fill="none" stroke="currentColor" viewBox="0 0 24 24" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6M1 7h22M9 7V4a1 1 0 011-1h4a1 1 0 011 1v3"/></svg>
        </button>
      </td>
    </tr>
  `).join('');
}

/* ============== EXPORT PDF ============== */
function exportPDF() {
  if (typeof window.jspdf === 'undefined' || !window.jspdf.jsPDF) {
    toast('❌ Error: librería PDF no cargada');
    return;
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const emp = empresa || {};

  // Encabezado
  doc.setFillColor(37,99,235);
  doc.rect(0,0,210,28,'F');
  doc.setTextColor(255,255,255);
  doc.setFontSize(16);
  doc.setFont('helvetica','bold');
  doc.text(emp.nombre || 'Empresa', 14, 13);
  doc.setFontSize(10);
  doc.setFont('helvetica','normal');
  doc.text(`RIF: ${emp.rif || 'N/A'}  |  Centro: ${emp.centro || 'N/A'}`, 14, 21);

  // Datos empresa
  doc.setTextColor(30,41,59);
  doc.setFontSize(12);
  doc.setFont('helvetica','bold');
  doc.text('Datos del Centro de Trabajo', 14, 38);
  doc.setFontSize(9);
  doc.setFont('helvetica','normal');
  const dir = `${emp.calle || ''}${emp.sector ? ', Sector '+emp.sector : ''}`;
  const ubi = `${emp.parroquia || ''}, ${emp.municipio || ''}, ${emp.estado || ''}`;
  doc.text(`Dirección: ${dir || 'N/A'}`, 14, 45);
  doc.text(`Ubicación: ${ubi || 'N/A'}`, 14, 51);
  let y = 57;
  if (emp.sucursal) {
    doc.text(`Sucursal: Sí - ${emp.sucursal_ubicacion || ''}`, 14, y); y += 6;
    if (emp.enCentroComercial) {
      doc.text(`Centro Comercial: ${emp.cc_nombre || ''} (${emp.cc_ciudad || ''})`, 14, y); y += 6;
    }
  }

  // Tabla trabajadores
  doc.setFontSize(12);
  doc.setFont('helvetica','bold');
  doc.text('Listado de Trabajadores', 14, y+8);
  const rows = trabajadores.map(t => [t.nombre, t.apellido, t.cedula, formatDate(t.fecha), t.cargo]);
  if (rows.length > 0) {
    doc.autoTable({
      startY: y+12,
      head: [['Nombre','Apellido','Cédula','Fecha Ingreso','Cargo']],
      body: rows,
      theme: 'striped',
      headStyles: { fillColor: [37,99,235], textColor: 255, fontStyle: 'bold' },
      styles: { fontSize: 9, cellPadding: 3 },
      alternateRowStyles: { fillColor: [241,245,249] }
    });
  } else {
    doc.setFontSize(10);
    doc.setFont('helvetica','italic');
    doc.text('No hay trabajadores registrados.', 14, y+20);
  }

  // Pie
  const pageCount = doc.internal.getNumberOfPages();
  for (let i=1; i<=pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(120);
    doc.text(`Generado: ${new Date().toLocaleString('es-VE')} — Página ${i} de ${pageCount}`, 14, 290);
  }

  const fileName = `Trabajadores_${(emp.nombre||'empresa').replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.pdf`;
  doc.save(fileName);
  toast('📄 PDF descargado: ' + fileName);
}

/* ============== EXPORT EXCEL ============== */
function exportExcel() {
  if (typeof XLSX === 'undefined') {
    toast('❌ Error: librería Excel no cargada');
    return;
  }
  const emp = empresa || {};
  const wb = XLSX.utils.book_new();

  // Hoja empresa
  const empData = [
    ['DATOS DE LA EMPRESA'],
    [],
    ['Campo', 'Valor'],
    ['Nombre', emp.nombre || ''],
    ['RIF', emp.rif || ''],
    ['Centro de Trabajo', emp.centro || ''],
    ['Estado', emp.estado || ''],
    ['Municipio', emp.municipio || ''],
    ['Parroquia', emp.parroquia || ''],
    ['Dirección', emp.calle || ''],
    ['Sector', emp.sector || ''],
    ['¿Es Sucursal?', emp.sucursal ? 'Sí' : 'No'],
    ['Ubicación Sucursal', emp.sucursal_ubicacion || ''],
    ['¿En Centro Comercial?', emp.enCentroComercial ? 'Sí' : 'No'],
    ['Nombre Centro Comercial', emp.cc_nombre || ''],
    ['Ciudad Centro Comercial', emp.cc_ciudad || '']
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(empData);
  ws1['!cols'] = [{wch:25},{wch:40}];
  ws1['!merges'] = [{s:{r:0,c:0},e:{r:0,c:1}}];
  XLSX.utils.book_append_sheet(wb, ws1, 'Empresa');

  // Hoja trabajadores
  const trabData = [['#','Nombre','Apellido','Cédula','Fecha Ingreso','Cargo']];
  trabajadores.forEach((t,i) => trabData.push([i+1, t.nombre, t.apellido, t.cedula, t.fecha, t.cargo]));
  const ws2 = XLSX.utils.aoa_to_sheet(trabData);
  ws2['!cols'] = [{wch:5},{wch:18},{wch:18},{wch:15},{wch:14},{wch:20}];
  XLSX.utils.book_append_sheet(wb, ws2, 'Trabajadores');

  const fileName = `Trabajadores_${(emp.nombre||'empresa').replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
  toast('📊 Excel descargado: ' + fileName);
}

/* ============== SINCRONIZACIÓN ============== */
function openSyncModal() {
  document.getElementById('syncModal').classList.remove('hidden');
  if (syncCfg) {
    setVal('sync_server', syncCfg.server);
    setVal('sync_apikey', syncCfg.apikey);
    setVal('sync_target', syncCfg.target);
  }
  updateSyncHints();
}
function closeSyncModal() { document.getElementById('syncModal').classList.add('hidden'); }
function switchSyncTab(t) {
  document.getElementById('panelConfig').classList.toggle('hidden', t!=='config');
  document.getElementById('panelActions').classList.toggle('hidden', t!=='actions');
  document.getElementById('tabConfig').classList.toggle('tab-active', t==='config');
  document.getElementById('tabActions').classList.toggle('tab-active', t==='actions');
  document.getElementById('tabConfig').classList.toggle('text-slate-500', t!=='config');
  document.getElementById('tabActions').classList.toggle('text-slate-500', t!=='actions');
}
function updateSyncHints() {
  const s = val('sync_server');
  const hints = {
    googlesheets: 'Configure una API Key de Google Cloud con acceso a Google Sheets API. En "Target" indique el ID de la hoja (ej. 1AbC...xYz).',
    googledrive: 'Use una API Key de Google Drive. En "Target" indique la carpeta o archivo JSON donde se guardarán los datos.',
    custom: 'Indique la URL completa del endpoint REST (POST para subir, GET para descargar). El token se enviará como Bearer.'
  };
  document.getElementById('sync_hint').textContent = hints[s];
}
function saveSyncConfig() {
  syncCfg = {
    server: val('sync_server'),
    apikey: val('sync_apikey'),
    target: val('sync_target')
  };
  localStorage.setItem('syncCfg', JSON.stringify(syncCfg));
  toast('✅ Configuración de sincronización guardada');
}
async function syncPush() {
  if (!syncCfg || !syncCfg.target) { toast('⚠️ Configure primero la sincronización'); return; }
  const payload = { empresa, trabajadores, timestamp: new Date().toISOString() };
  try {
    if (syncCfg.server === 'custom') {
      const res = await fetch(syncCfg.target, {
        method: 'POST',
        headers: { 'Content-Type':'application/json', 'Authorization':'Bearer '+syncCfg.apikey },
        body: JSON.stringify(payload)
      });
      if (!res.ok) throw new Error('HTTP '+res.status);
      toast('✅ Datos subidos al servidor');
    } else {
      const blob = new Blob([JSON.stringify(payload,null,2)], {type:'application/json'});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'backup_sincronizacion.json';
      a.click();
      URL.revokeObjectURL(url);
      toast('📦 Archivo JSON generado. Para integración automática, configure un Apps Script Web App.');
    }
  } catch(e) {
    toast('❌ Error al sincronizar: '+e.message);
  }
}
async function syncPull() {
  if (!syncCfg || !syncCfg.target) { toast('⚠️ Configure primero la sincronización'); return; }
  try {
    if (syncCfg.server === 'custom') {
      const res = await fetch(syncCfg.target, {
        headers: { 'Authorization':'Bearer '+syncCfg.apikey }
      });
      if (!res.ok) throw new Error('HTTP '+res.status);
      const data = await res.json();
      if (data.empresa) { empresa = data.empresa; localStorage.setItem('empresa', JSON.stringify(empresa)); loadEmpresa(); }
      if (data.trabajadores) { trabajadores = data.trabajadores; localStorage.setItem('trabajadores', JSON.stringify(trabajadores)); renderTrabajadores(); }
      toast('✅ Datos descargados del servidor');
    } else {
      const inp = document.createElement('input');
      inp.type = 'file'; inp.accept = 'application/json';
      inp.onchange = e => {
        const f = e.target.files[0]; if (!f) return;
        const r = new FileReader();
        r.onload = ev => {
          try {
            const data = JSON.parse(ev.target.result);
            if (data.empresa) { empresa = data.empresa; localStorage.setItem('empresa', JSON.stringify(empresa)); loadEmpresa(); }
            if (data.trabajadores) { trabajadores = data.trabajadores; localStorage.setItem('trabajadores', JSON.stringify(trabajadores)); renderTrabajadores(); }
            toast('✅ Datos importados desde archivo');
          } catch { toast('❌ Archivo inválido'); }
        };
        r.readAsText(f);
      };
      inp.click();
    }
  } catch(e) {
    toast('❌ Error al sincronizar: '+e.message);
  }
}

/* ============== UTILS ============== */
function val(id){ return document.getElementById(id).value; }
function setVal(id,v){ document.getElementById(id).value = v || ''; }
function escapeHtml(s){ return (s||'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function formatDate(d){ if(!d) return ''; const [y,m,day] = d.split('-'); return `${day}/${m}/${y}`; }
let toastTimer;
function toast(msg) {
  const t = document.getElementById('toast');
  document.getElementById('toastMsg').textContent = msg;
  t.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add('hidden'), 3200);
}

init();
</script>
</body>
</html>
