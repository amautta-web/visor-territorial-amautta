/* ─────────────────────────────────────────────────────────────────────────────
   VARIABLES GLOBALES
   ───────────────────────────────────────────────────────────────────────────── */

let RAW_ROWS = [];
let communeData = {};
let layersMap = {};
let activeCommune = null;
let currentSort = 'total';
let geoJsonLoaded = false;
let globalGeojson = null;
let geoLayer = null;

let filtroAñoActual = 'ALL';
let filtroTipoProyecto = 'GENERAL';

// Total de meses distintos presentes en el filtro actual (todas las comunas
// combinadas). Se usa como referencia para detectar comunas con reporte
// incompleto (ver renderList).
let _globalMesesCount = 0;

// Tweens de GSAP activos por pill, para poder cancelarlos si el usuario
// cambia de filtro antes de que termine la animación anterior.
const _pillTweens = {};

// ── Cache de agregación ───────────────────────────────────────────────────────
// Evita recalcular aggregateData() desde cero en cada cambio de filtro.
// Clave: JSON de los filtros activos → valor: communeData ya calculado.
const _aggregateCache = new Map();
const CACHE_MAX = 40; // máximo de entradas a conservar

// ── Índices precalculados ─────────────────────────────────────────────────────
// Se construyen UNA sola vez al cargar los datos. Cada filtro individual tiene
// su propio Set/Map para que filtrar sea O(n·k) donde k = nº de filtros activos.
let _indexByYear   = {};  // { 2024: Set<idx>, 2025: Set<idx>, 2026: Set<idx> }
let _indexByTipo   = {};  // { RO: Set<idx>, PP: Set<idx> }
let _indexByMes    = {};  // { 'ENERO': Set<idx>, ... }
let _indexBySvc    = {};  // { 'servicio...': Set<idx>, ... }
let _indexByCod    = {};  // { '200211': Set<idx>, ... }
let _indexAll      = null; // Set con todos los índices

// ── Debounce para applyFilters ────────────────────────────────────────────────
let _filterTimer = null;

const MESES_ORDER = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

const SERVICE_COLORS = [
    '#00d4aa','#4f9cf4','#f4a84f','#f46d6d','#b06df4',
    '#f4e24f','#4ff4a8','#f44fb0','#7af44f','#4fd4f4'
];

/* ─────────────────────────────────────────────────────────────────────────────
   CONFIGURACIÓN DE PROYECTOS
   ───────────────────────────────────────────────────────────────────────────── */

const PROJECT_METADATA = {
    "200211": { tipo: "RO" },
    "240013": { tipo: "RO" },
    "240034": { tipo: "RO" },
    "230013": { tipo: "PP" },
    "250018": { tipo: "PP" },
    "250071": { tipo: "PP" },
};

function getTipoProyecto(row) {
    // Usa el valor precalculado si existe (se guarda en buildIndexes)
    if (row.__TIPO) return row.__TIPO;
    const colOficial = row['TIPO DE RECURSO'] || row['TIPO DE RECURSO '];
    if (colOficial) {
        const v = String(colOficial).trim().toUpperCase();
        if (v === 'RO' || v === 'PP') return v;
    }
    const cod = row.__COD;
    if (cod && PROJECT_METADATA[cod]) return PROJECT_METADATA[cod].tipo;
    // Fallback heredado: solo si el código de proyecto no está catalogado
    if (row.__YEAR === 2026) return 'PP';
    return null;
}

let spotsLayerGroup = L.layerGroup();

function setLoad(pct, txt) {
    document.getElementById('loadBar').style.width = pct + '%';
    document.getElementById('loadingStatus').textContent = txt;
}

/* ─────────────────────────────────────────────────────────────────────────────
   ESCALA DE COLORES
   ───────────────────────────────────────────────────────────────────────────── */

function getColor(n) {
    return n > 4000 ? '#007d55'
        : n > 2000 ? '#7fa13b'
        : n > 1000 ? '#d4b22f'
        : n > 300  ? '#c95728'
        : n > 0    ? '#a82c2c'
        : '#141e2e';
}

/* ─────────────────────────────────────────────────────────────────────────────
   MAPA
   ───────────────────────────────────────────────────────────────────────────── */

const map = L.map('map', { center: [6.2518, -75.5636], zoom: 12, zoomControl: true });
L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
    attribution: '©OpenStreetMap ©CartoDB', maxZoom: 19
}).addTo(map);
spotsLayerGroup.addTo(map);

/* ─────────────────────────────────────────────────────────────────────────────
   LECTURA DE ARCHIVOS EXCEL  (sin cambios funcionales; solo añade log de tiempo)
   ───────────────────────────────────────────────────────────────────────────── */

async function leerArchivoExcel(nombreArchivo, año, fallbackMsg) {
    let data;
    try {
        const response = await fetch(nombreArchivo);
        if (!response.ok) throw new Error('HTTP ' + response.status);
        data = await response.arrayBuffer();
    } catch (e) {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        input.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:99999;padding:12px 20px;background:#00d4aa;color:#000;border:none;border-radius:8px;font-size:14px;cursor:pointer;font-family:inherit;';
        const lbl = document.createElement('div');
        lbl.style.cssText = 'position:fixed;top:calc(50% - 40px);left:50%;transform:translateX(-50%);z-index:99999;color:var(--text);font-size:13px;background:var(--surface2);padding:8px 16px;border-radius:8px;';
        lbl.textContent = fallbackMsg;
        document.body.appendChild(lbl);
        document.body.appendChild(input);
        data = await new Promise(resolve => {
            input.onchange = e => {
                const reader = new FileReader();
                reader.onload = ev => resolve(ev.target.result);
                reader.readAsArrayBuffer(e.target.files[0]);
                document.body.removeChild(input);
                document.body.removeChild(lbl);
            };
        });
    }

    const wb = XLSX.read(data, { type: 'array' });
    let sheetName = wb.SheetNames[0];
    if (año === 2024 && wb.SheetNames.includes('PLAN DE ACCION 2024')) {
        sheetName = 'PLAN DE ACCION 2024';
    }
    const ws = wb.Sheets[sheetName];
    const filas = XLSX.utils.sheet_to_json(ws);
    filas.forEach(row => { row.__YEAR = año; });
    return filas;
}

/* ─────────────────────────────────────────────────────────────────────────────
   PRECÁLCULO DE CAMPOS FRECUENTES  (se hace una sola vez)
   
   Normaliza y guarda en __COD, __TIPO, __COMUNAID, __MES, __SVC, __SEXO,
   __EDAD para que aggregateData() nunca repita String / trim / parseInt.
   ───────────────────────────────────────────────────────────────────────────── */

function precalcularCampos(rows) {
    rows.forEach(row => {
        // Código de proyecto
        row.__COD = obtenerCodigoProyectoRaw(row);

        // Tipo RO / PP
        const colOficial = row['TIPO DE RECURSO'] || row['TIPO DE RECURSO '];
        if (colOficial) {
            const v = String(colOficial).trim().toUpperCase();
            if (v === 'RO' || v === 'PP') { row.__TIPO = v; }
        }
        if (!row.__TIPO) {
            if (row.__COD && PROJECT_METADATA[row.__COD]) row.__TIPO = PROJECT_METADATA[row.__COD].tipo;
            // Fallback heredado: solo si el código de proyecto no está catalogado
            else if (row.__YEAR === 2026) row.__TIPO = 'PP';
            else row.__TIPO = null;
        }

        // Comuna ID
        const comunaStr = (row['* COMUNA DE RESIDENCIA'] || '').toString().trim();
        const match = comunaStr.match(/\d+/);
        row.__COMUNAID  = match ? parseInt(match[0]) : null;
        row.__COMUNASTR = comunaStr;

        // Mes
        row.__MES = (row[' MES DE REPORTE'] || '').trim().toUpperCase();

        // Servicio
        row.__SVC = (row['BIEN, PRODUCTO, SERVICIO RECIBIDO'] || '').trim();

        // Sexo
        row.__SEXO = (row['* SEXO'] || '').trim().toUpperCase();

        // Edad
        row.__EDAD = parseInt(row['AÑOS CUMPLIDOS AL INGRESO DEL PROGRAMA']) || 0;

        // Otros
        row.__ETNIA   = (row['* ETNIA'] || '').trim();
        row.__DISCAP  = (row['* CONDICIÓN DE DISCAPACIDAD'] || '').trim();
        row.__ESTRATO = (row['ESTRATO SOCIOECONÓMICO'] || '').trim();
        row.__VICTIMA = (row['CONDICIÓN DE VÍCTIMA/HECHO VICTIMIZANTE'] || '').trim();
        row.__AREA    = (row['ÁREA\n(RURAL/URBANA)'] || row['ÁREA (RURAL/URBANA)'] || '').trim().toUpperCase();
    });
}

/* ─────────────────────────────────────────────────────────────────────────────
   CONSTRUCCIÓN DE ÍNDICES INVERTIDOS
   ───────────────────────────────────────────────────────────────────────────── */

function buildIndexes(rows) {
    _indexByYear  = {};
    _indexByTipo  = {};
    _indexByMes   = {};
    _indexBySvc   = {};
    _indexByCod   = {};
    _indexAll     = new Set();

    rows.forEach((row, i) => {
        _indexAll.add(i);

        // Año
        const y = row.__YEAR;
        if (!_indexByYear[y]) _indexByYear[y] = new Set();
        _indexByYear[y].add(i);

        // Tipo
        const t = row.__TIPO;
        if (t) {
            if (!_indexByTipo[t]) _indexByTipo[t] = new Set();
            _indexByTipo[t].add(i);
        }

        // Mes
        const m = row.__MES;
        if (m) {
            if (!_indexByMes[m]) _indexByMes[m] = new Set();
            _indexByMes[m].add(i);
        }

        // Servicio
        const s = row.__SVC;
        if (s) {
            if (!_indexBySvc[s]) _indexBySvc[s] = new Set();
            _indexBySvc[s].add(i);
        }

        // Código de proyecto
        const c = row.__COD;
        if (c) {
            if (!_indexByCod[c]) _indexByCod[c] = new Set();
            _indexByCod[c].add(i);
        }
    });
}

/* ─────────────────────────────────────────────────────────────────────────────
   INTERSECCIÓN RÁPIDA DE ÍNDICES
   Usa el Set más pequeño como punto de partida (reduce iteraciones).
   ───────────────────────────────────────────────────────────────────────────── */

function intersectSets(sets) {
    // Ordena por tamaño ascendente
    sets.sort((a, b) => a.size - b.size);
    const [first, ...rest] = sets;
    const result = new Set();
    first.forEach(i => {
        if (rest.every(s => s.has(i))) result.add(i);
    });
    return result;
}

/* ─────────────────────────────────────────────────────────────────────────────
   CLAVE DE CACHE para los filtros actuales
   ───────────────────────────────────────────────────────────────────────────── */

function buildCacheKey() {
    return JSON.stringify({
        año:   filtroAñoActual,
        tipo:  filtroTipoProyecto,
        mes:   document.getElementById('mesFilter').value.toUpperCase(),
        svc:   document.getElementById('servicioFilter').value.trim(),
        proj:  document.getElementById('subProyectoFilter').value,
    });
}

/* ─────────────────────────────────────────────────────────────────────────────
   LECTURA EXCEL + ORQUESTACIÓN INICIAL
   ───────────────────────────────────────────────────────────────────────────── */

async function cargarExcel() {
    setLoad(10, 'Leyendo archivo 2026…');
    const filas2026 = await leerArchivoExcel('BASE_2026.xlsx', 2026, 'Selecciona BASE_2026.xlsx');

    setLoad(30, 'Leyendo archivo 2025…');
    const filas2025 = await leerArchivoExcel('PA_PM_2025.xlsx', 2025, 'Selecciona PA_PM_2025.xlsx');

    setLoad(50, 'Leyendo archivo 2024…');
    const filas2024 = await leerArchivoExcel('PA_CONSOLIDADO_2024.xlsx', 2024, 'Selecciona PA_CONSOLIDADO_2024.xlsx');

    RAW_ROWS = [...filas2026, ...filas2025, ...filas2024];

    setLoad(60, `Precalculando campos (${RAW_ROWS.length.toLocaleString('es-CO')} registros)…`);
    // Yield al navegador antes de la operación pesada
    await yieldToMain();
    precalcularCampos(RAW_ROWS);

    setLoad(68, 'Construyendo índices…');
    await yieldToMain();
    buildIndexes(RAW_ROWS);

    setLoad(72, 'Preparando filtros…');
    await yieldToMain();
    buildFilters();
    const _urlState = restoreUrlState();

    setLoad(78, 'Calculando métricas por comuna…');
    await yieldToMain();
    aggregateData();

    setLoad(85, 'Cargando mapa de Medellín…');
    await cargarGeoJSON();

    // Si el link traía una comuna específica, la abrimos una vez que el
    // mapa terminó de cargar (con un pequeño respiro para no pisar la
    // animación de entrada).
    if (_urlState.comuna) {
        const idComuna = parseInt(_urlState.comuna);
        setTimeout(() => {
            if (communeData[idComuna]) showDetail(idComuna);
            else showNoDataDetail(idComuna);
        }, 900);
    }
}

/* ─────────────────────────────────────────────────────────────────────────────
   ESTADO EN LA URL
   Refleja los filtros activos (y la comuna abierta, si hay una) como query
   params, para poder copiar el link tal cual y que otra persona vea
   exactamente la misma vista sin que le tengas que explicar los pasos.
   Usa replaceState (no pushState) para no llenar el historial del navegador
   con una entrada por cada cambio de filtro.
   ───────────────────────────────────────────────────────────────────────────── */

function syncUrlState() {
    const params = new URLSearchParams();

    if (filtroAñoActual !== 'ALL') params.set('anio', filtroAñoActual);
    if (filtroTipoProyecto !== 'GENERAL') params.set('tipo', filtroTipoProyecto);

    const mes = document.getElementById('mesFilter').value;
    if (mes) params.set('mes', mes);

    const svc = document.getElementById('servicioFilter').value;
    if (svc) params.set('servicio', svc);

    const proj = document.getElementById('subProyectoFilter').value;
    if (proj) params.set('proyecto', proj);

    if (activeCommune) params.set('comuna', activeCommune);

    const qs = params.toString();
    const newUrl = window.location.pathname + (qs ? '?' + qs : '');
    history.replaceState(null, '', newUrl);
}

function restoreUrlState() {
    const params = new URLSearchParams(window.location.search);

    const tipo = params.get('tipo');
    if (tipo === 'RO' || tipo === 'PP') {
        filtroTipoProyecto = tipo;
        document.querySelectorAll('.vtab').forEach(t => t.classList.remove('active'));
        const tab = document.getElementById(tipo === 'RO' ? 'vtab-ro' : 'vtab-pp');
        if (tab) tab.classList.add('active');
        rebuildSubProyectoSelector();
    }

    const proj = params.get('proyecto');
    if (proj) document.getElementById('subProyectoFilter').value = proj;

    const anio = params.get('anio');
    if (anio) {
        filtroAñoActual = anio;
        document.getElementById('añoFilter').value = anio;
    }

    const mes = params.get('mes');
    if (mes) document.getElementById('mesFilter').value = mes;

    const svc = params.get('servicio');
    if (svc) document.getElementById('servicioFilter').value = svc;

    return { comuna: params.get('comuna') };
}

function copyShareLink() {
    const btn = document.getElementById('shareBtn');
    const done = () => {
        const original = btn.dataset.original || btn.textContent;
        btn.dataset.original = original;
        btn.textContent = '✓';
        btn.classList.add('copied');
        setTimeout(() => { btn.textContent = original; btn.classList.remove('copied'); }, 1400);
    };
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(window.location.href).then(done).catch(() => {
            prompt('Copia este link:', window.location.href);
        });
    } else {
        prompt('Copia este link:', window.location.href);
    }
}

/* ─────────────────────────────────────────────────────────────────────────────
   YIELD AL HILO PRINCIPAL  (evita bloquear UI durante operaciones pesadas)
   ───────────────────────────────────────────────────────────────────────────── */

function yieldToMain() {
    return new Promise(resolve => {
        if (typeof scheduler !== 'undefined' && scheduler.yield) {
            scheduler.yield().then(resolve);
        } else {
            setTimeout(resolve, 0);
        }
    });
}

/* ─────────────────────────────────────────────────────────────────────────────
   CÓDIGO DE PROYECTO (versión raw, sin precálculo)
   ───────────────────────────────────────────────────────────────────────────── */

function obtenerCodigoProyectoRaw(row) {
    const colCanonica = row['* CÓDIGO PROYECTO'] || row['CÓDIGO PROYECTO'] ||
                        row['* CODIGO PROYECTO'] || row['CODIGO PROYECTO'];
    if (colCanonica) {
        const val = String(colCanonica).trim();
        if (val && val !== 'undefined' && val !== 'null') return val;
    }
    for (let key in row) {
        if (key === '__YEAR') continue;
        const val = String(row[key]).trim();
        if (/^\d{6}$/.test(val)) return val;
    }
    return null;
}

// Versión que usa el campo precalculado cuando disponible
function obtenerCodigoProyecto(row) {
    return row.__COD !== undefined ? row.__COD : obtenerCodigoProyectoRaw(row);
}

/* ─────────────────────────────────────────────────────────────────────────────
   CAMBIO DE PESTAÑA DE PROYECTO
   ───────────────────────────────────────────────────────────────────────────── */

function setProyectoTipo(tipo, el) {
    filtroTipoProyecto = tipo;
    document.querySelectorAll('.vtab').forEach(t => t.classList.remove('active'));
    el.classList.add('active');
    rebuildSubProyectoSelector();
    applyFilters();
}

function filtrarAño(valor) {
    filtroAñoActual = valor;
    applyFilters();
}

/* ─────────────────────────────────────────────────────────────────────────────
   CONSTRUCCIÓN DE FILTROS (UI)
   ───────────────────────────────────────────────────────────────────────────── */

function buildFilters() {
    // Meses — usa el índice ya construido
    const mesSelect = document.getElementById('mesFilter');
    MESES_ORDER.filter(m => _indexByMes[m]).forEach(m => {
        const o = document.createElement('option');
        o.value = m;
        o.textContent = m[0] + m.slice(1).toLowerCase();
        mesSelect.appendChild(o);
    });

    // Servicios — usa el índice ya construido
    const svcSelect = document.getElementById('servicioFilter');
    Object.entries(_indexBySvc)
        .map(([s, set]) => [s, set.size])
        .sort((a, b) => b[1] - a[1])
        .forEach(([s, cnt]) => {
            const o = document.createElement('option');
            o.value = s;
            const short = s.length > 40 ? s.slice(0, 40) + '…' : s;
            o.textContent = `${short} (${cnt.toLocaleString('es-CO')})`;
            svcSelect.appendChild(o);
        });

    rebuildSubProyectoSelector();
}

function rebuildSubProyectoSelector() {
    const sel = document.getElementById('subProyectoFilter');
    sel.innerHTML = '';

    const labelTodos = filtroTipoProyecto === 'RO' ? 'Todos los proyectos RO'
        : filtroTipoProyecto === 'PP' ? 'Todos los proyectos PP'
        : 'Todos los proyectos';
    const optAll = document.createElement('option');
    optAll.value = '';
    optAll.textContent = labelTodos;
    sel.appendChild(optAll);

    // Filtra por tipo usando índice
    const candidatos = filtroTipoProyecto !== 'GENERAL'
        ? (_indexByTipo[filtroTipoProyecto] || new Set())
        : _indexAll;

    const conteo = {};
    candidatos.forEach(i => {
        const cod = RAW_ROWS[i].__COD;
        if (cod) conteo[cod] = (conteo[cod] || 0) + 1;
    });

    Object.entries(conteo).sort((a, b) => b[1] - a[1]).forEach(([cod, cnt]) => {
        const o = document.createElement('option');
        o.value = cod;
        o.textContent = `${cod} (${cnt.toLocaleString('es-CO')})`;
        sel.appendChild(o);
    });
}

/* ─────────────────────────────────────────────────────────────────────────────
   AGREGACIÓN DE DATOS  (versión con índices + cache)
   ───────────────────────────────────────────────────────────────────────────── */

function aggregateData() {
    const cacheKey = buildCacheKey();
    if (_aggregateCache.has(cacheKey)) {
        communeData = _aggregateCache.get(cacheKey).communeData;
        updateHeaderPills(_aggregateCache.get(cacheKey).filteredIndices);
        renderList();
        if (geoJsonLoaded) { updateMapColors(); renderSpots(); }
        return;
    }

    const mesFilter   = document.getElementById('mesFilter').value.toUpperCase();
    const svcFilter   = document.getElementById('servicioFilter').value.trim();
    const subProjFilter = document.getElementById('subProyectoFilter').value;

    // Construye la lista de Sets a intersectar
    const sets = [];

    if (filtroTipoProyecto !== 'GENERAL') {
        sets.push(_indexByTipo[filtroTipoProyecto] || new Set());
    }
    if (filtroAñoActual !== 'ALL') {
        sets.push(_indexByYear[parseInt(filtroAñoActual)] || new Set());
    }
    if (mesFilter) {
        sets.push(_indexByMes[mesFilter] || new Set());
    }
    if (svcFilter) {
        sets.push(_indexBySvc[svcFilter] || new Set());
    }
    if (subProjFilter) {
        sets.push(_indexByCod[subProjFilter] || new Set());
    }

    // Si no hay filtros, usa todos los índices
    const filteredIndices = sets.length === 0
        ? _indexAll
        : sets.length === 1
            ? sets[0]
            : intersectSets(sets);

    communeData = {};

    filteredIndices.forEach(i => {
        const r = RAW_ROWS[i];
        const id = r.__COMUNAID;
        if (!id) return;

        if (!communeData[id]) {
            communeData[id] = {
                nameStr: r.__COMUNASTR, id,
                total: 0, mujeres: 0, hombres: 0, indefinido: 0,
                edadSuma: 0, edadCnt: 0, edad_promedio: 0,
                servicios: {}, etnias: {}, discapacidades: {}, estratos: {},
                victimas: 0, rural: 0, urbana: 0, meses: new Set(), mesesCount: 0
            };
        }

        const c = communeData[id];
        c.total++;
        if (r.__MES) c.meses.add(r.__MES);
        if (r.__SEXO === 'MUJER') c.mujeres++;
        else if (r.__SEXO === 'HOMBRE') c.hombres++;
        else c.indefinido++;

        const edad = r.__EDAD;
        if (edad > 0 && edad < 130) { c.edadSuma += edad; c.edadCnt++; }

        if (r.__SVC) c.servicios[r.__SVC] = (c.servicios[r.__SVC] || 0) + 1;

        const etnia = r.__ETNIA;
        if (etnia && etnia !== 'NINGUNO' && etnia !== 'SIN DATO')
            c.etnias[etnia] = (c.etnias[etnia] || 0) + 1;

        const discap = r.__DISCAP;
        if (discap && discap !== 'NO TIENE DISCAPACIDAD' && discap !== 'SIN DATO')
            c.discapacidades[discap] = (c.discapacidades[discap] || 0) + 1;

        if (r.__ESTRATO) c.estratos[r.__ESTRATO] = (c.estratos[r.__ESTRATO] || 0) + 1;
        if (r.__VICTIMA && r.__VICTIMA !== 'NINGUNA' && r.__VICTIMA !== 'SIN DATO') c.victimas++;
        if (r.__AREA === 'RURAL') c.rural++;
        else if (r.__AREA === 'URBANA') c.urbana++;
    });

    Object.values(communeData).forEach(c => {
        c.edad_promedio = c.edadCnt > 0 ? parseFloat((c.edadSuma / c.edadCnt).toFixed(1)) : 0;
        c.mesesCount = c.meses.size;
    });

    // Guardar en cache (con límite de tamaño)
    if (_aggregateCache.size >= CACHE_MAX) {
        const firstKey = _aggregateCache.keys().next().value;
        _aggregateCache.delete(firstKey);
    }
    _aggregateCache.set(cacheKey, { communeData: { ...communeData }, filteredIndices });

    updateHeaderPills(filteredIndices);
    renderList();
    if (geoJsonLoaded) { updateMapColors(); renderSpots(); }
}

/* ─────────────────────────────────────────────────────────────────────────────
   applyFilters con debounce  (evita recalcular mientras el usuario sigue haciendo
   cambios rápidos en los selectores)
   ───────────────────────────────────────────────────────────────────────────── */

function applyFilters() {
    clearTimeout(_filterTimer);
    _filterTimer = setTimeout(() => {
        aggregateData();
        closeDetail();
    }, 80);
}

/* ─────────────────────────────────────────────────────────────────────────────
   PILLS DEL HEADER  (recibe un Set o Array de índices)
   ───────────────────────────────────────────────────────────────────────────── */

function updateHeaderPills(indices) {
    const total = indices.size !== undefined ? indices.size : indices.length;

    // Calcula comunas desde communeData (ya calculado)
    const comunas = Object.keys(communeData).length;

    // Servicios únicos y edad promedio — itera una sola vez
    const svcSet  = new Set();
    let edadSuma  = 0, edadCnt = 0;
    const mesesSet = new Set();

    indices.forEach(i => {
        const r = RAW_ROWS[i];
        if (r.__SVC) svcSet.add(r.__SVC);
        if (r.__EDAD > 0 && r.__EDAD < 130) { edadSuma += r.__EDAD; edadCnt++; }
        if (r.__MES) mesesSet.add(r.__MES);
    });

    _globalMesesCount = mesesSet.size;

    animateCounter('pill-total', total);
    animateCounter('pill-comunas', comunas);
    animateCounter('pill-servicios', svcSet.size);
    animateCounter('pill-edad', edadCnt ? edadSuma / edadCnt : null, { decimals: 1, suffix: ' a' });
    animateCounter('pill-meses', mesesSet.size);
}

/* ─────────────────────────────────────────────────────────────────────────────
   CONTADOR ANIMADO (GSAP)
   Anima el valor numérico de una pill desde su valor actual hasta el nuevo,
   en vez de que el texto cambie de golpe. Cancela cualquier tween anterior
   sobre la misma pill para que cambios de filtro rápidos no se pisen entre sí.
   ───────────────────────────────────────────────────────────────────────────── */

function animateCounter(elId, endValue, opts = {}) {
    const el = document.getElementById(elId);
    if (!el) return;
    if (_pillTweens[elId]) _pillTweens[elId].kill();

    if (endValue === null || endValue === undefined || isNaN(endValue)) {
        el.textContent = opts.fallback !== undefined ? opts.fallback : '—';
        return;
    }

    const current = parseFloat((el.textContent || '').replace(/[^\d.-]/g, '')) || 0;
    const obj = { value: current };
    _pillTweens[elId] = gsap.to(obj, {
        value: endValue,
        duration: 0.55,
        ease: 'power2.out',
        onUpdate: () => {
            const num = opts.decimals
                ? obj.value.toFixed(opts.decimals)
                : Math.round(obj.value).toLocaleString('es-CO');
            el.textContent = (opts.prefix || '') + num + (opts.suffix || '');
        }
    });
}

/* ─────────────────────────────────────────────────────────────────────────────
   CARGA DEL GEOJSON
   ───────────────────────────────────────────────────────────────────────────── */

async function cargarGeoJSON() {
    let geojson;
    try {
        const r = await fetch('medellin.geojson');
        geojson = await r.json();
    } catch (e) {
        setLoad(100, 'medellin.geojson no encontrado. Colócalo junto al HTML.');
        setTimeout(() => document.getElementById('loadingOverlay').style.display = 'none', 2000);
        return;
    }

    globalGeojson = geojson;
    geoJsonLoaded = true;

    geoLayer = L.geoJSON(geojson, {
        style: feature => styleFeature(feature),
        onEachFeature: (feature, layer) => {
            const numGeo = parseInt(feature.properties.CODIGO);
            layersMap[numGeo] = layer;

            layer.on({
                click: () => {
                    const d = communeData[numGeo];
                    if (d) showDetail(numGeo);
                    else showNoDataDetail(numGeo);
                },
                mouseover: e => {
                    e.target.setStyle({ weight: 2.5, color: '#00d4aa', fillOpacity: 0.82 });
                    updateTooltip(e.target, numGeo);
                },
                mouseout: e => {
                    if (activeCommune !== numGeo) geoLayer.resetStyle(e.target);
                }
            });
        }
    }).addTo(map);

    setLoad(95, 'Preparando interfaz…');
    renderSpots();
    finishLoad();
}

function updateTooltip(layer, numGeo) {
    const d = communeData[numGeo];
    if (!d || d.total === 0) {
        layer.bindTooltip(`<strong>Comuna ${numGeo}</strong><br>Esta comuna no priorizó recursos`, { sticky: true }).openTooltip();
        return;
    }
    const topSvc = Object.entries(d.servicios).sort((a, b) => b[1] - a[1]).slice(0, 2)
        .map(([k, v]) => `<span style="color:#7a90a8">${k.slice(0, 30)}</span>: ${v.toLocaleString('es-CO')}`).join('<br>');
    layer.bindTooltip(`<strong>${d.nameStr}</strong><br>👥 ${d.total.toLocaleString('es-CO')} registros<br>${topSvc}`, { sticky: true }).openTooltip();
}

function styleFeature(feature) {
    const numGeo = parseInt(feature.properties.CODIGO);
    const d = communeData[numGeo];
    const total = d ? d.total : 0;
    return {
        fillColor: getColor(total),
        weight: 1,
        opacity: 1,
        color: '#0d1520',
        fillOpacity: total > 0 ? 0.80 : 0.12
    };
}

/* ─────────────────────────────────────────────────────────────────────────────
   updateMapColors  —  usa requestAnimationFrame para no bloquear el hilo
   ───────────────────────────────────────────────────────────────────────────── */

function updateMapColors() {
    if (!geoLayer) return;
    requestAnimationFrame(() => {
        geoLayer.setStyle(feature => styleFeature(feature));
    });
}

function renderSpots() {
    spotsLayerGroup.clearLayers();
    // (reservado para marcadores futuros)
}

function finishLoad() {
    setLoad(100, '¡Listo!');
    
    // Esperamos un instante antes de iniciar la magia
    setTimeout(() => {
        // Creamos una línea de tiempo de GSAP
        const tl = gsap.timeline();
        
        // 1. El overlay de carga se desliza hacia arriba
        tl.to("#loadingOverlay", {
            y: "-100%",
            opacity: 0,
            duration: 0.8,
            ease: "power3.inOut",
            onComplete: () => {
                document.getElementById('loadingOverlay').style.display = 'none';
            }
        })
        // 2. El header cae desde arriba con un rebote
        .from("header", {
            y: -30,
            opacity: 0,
            duration: 0.6,
            ease: "back.out(1.5)"
        }, "-=0.3") // Inicia 0.3 segundos antes de que termine la animación anterior
        // 3. Las pestañas (General, RO, PP) se revelan
        .from(".view-tabs", {
            y: -15,
            opacity: 0,
            duration: 0.4,
            ease: "power2.out"
        }, "-=0.4")
        // 4. El sidebar entra desde la izquierda
        .from(".sidebar", {
            x: -40,
            opacity: 0,
            duration: 0.6,
            ease: "power3.out"
        }, "-=0.4")
        // 5. El mapa hace un ligero zoom-in y fade
        .from("#mapView", {
            scale: 0.97,
            opacity: 0,
            duration: 0.8,
            ease: "power3.out"
        }, "-=0.6");
        
    }, 400);
}

/* ─────────────────────────────────────────────────────────────────────────────
   LISTA DE COMUNAS  —  renderización virtualizada básica
   Solo renderiza las comunas visibles + un buffer, usando un contenedor con
   altura fija para evitar reflowing completo del DOM.
   ───────────────────────────────────────────────────────────────────────────── */

// Cache de los ítems ya renderizados
const _renderedItems = new Map(); // id → elemento DOM

function renderList() {
    const list = document.getElementById('communeList');
    const query = document.getElementById('search').value.toLowerCase();

    let entries = Object.entries(communeData);
    if (query) entries = entries.filter(([, d]) => d.nameStr.toLowerCase().includes(query));
    entries.sort((a, b) => b[1][currentSort] - a[1][currentSort]);

    // Diferencial: solo actualiza lo necesario
    const existingIds = new Set([...list.querySelectorAll('.commune-item')].map(el => el.dataset.id));
    const newIds = new Set(entries.map(([id]) => String(id)));

    // Elimina los que ya no están
    existingIds.forEach(id => {
        if (!newIds.has(id)) {
            const el = list.querySelector(`[data-id="${id}"]`);
            if (el) el.remove();
        }
    });

    // Fragment para insertar/reordenar de una vez
    const frag = document.createDocumentFragment();

    entries.forEach(([id, d]) => {
        let item = list.querySelector(`[data-id="${id}"]`);
        const isActive = parseInt(id) === activeCommune;
        const topSvc = Object.entries(d.servicios).sort((a, b) => b[1] - a[1])[0];
        const topSvcText = topSvc
            ? topSvc[0].slice(0, 28) + (topSvc[0].length > 28 ? '…' : '') + ` (${topSvc[1].toLocaleString('es-CO')})`
            : '—';
        const countVal = d[currentSort].toLocaleString('es-CO');

        if (!item) {
            item = document.createElement('div');
            item.className = 'commune-item';
            item.dataset.id = id;
            item.innerHTML = `
                <div class="commune-dot" style="background:${getColor(d.total)}"></div>
                <div class="commune-info">
                    <div class="commune-name">${d.nameStr}</div>
                    <div class="commune-sub"></div>
                </div>
                <div class="commune-right">
                    <div class="commune-count"></div>
                    <div class="commune-meses"></div>
                </div>
            `;
            item.addEventListener('click', () => {
                showDetail(parseInt(id));
                const layer = layersMap[parseInt(id)];
                if (layer) map.flyToBounds(layer.getBounds(), { maxZoom: 14, duration: 0.8 });
            });
        }

        // Actualiza solo los campos que cambian
        item.querySelector('.commune-dot').style.background = getColor(d.total);
        item.querySelector('.commune-name').textContent = d.nameStr;
        item.querySelector('.commune-sub').textContent = topSvcText;
        item.querySelector('.commune-count').textContent = countVal;
        item.classList.toggle('active', isActive);

        // Indicador de meses reportados — solo tiene sentido mostrarlo
        // cuando el filtro actual abarca más de un mes; si la comuna tiene
        // menos meses con datos que el máximo visto en el resto del listado,
        // se resalta como posible reporte incompleto.
        const mesesEl = item.querySelector('.commune-meses');
        if (_globalMesesCount > 1) {
            const incompleto = d.mesesCount < _globalMesesCount;
            mesesEl.textContent = `${d.mesesCount}/${_globalMesesCount} meses`;
            mesesEl.classList.toggle('incomplete', incompleto);
            mesesEl.title = incompleto
                ? `Esta comuna solo tiene datos en ${d.mesesCount} de ${_globalMesesCount} meses del periodo filtrado — revisar si falta reporte.`
                : '';
        } else {
            mesesEl.textContent = '';
            mesesEl.classList.remove('incomplete');
        }

        frag.appendChild(item);
    });

    list.appendChild(frag);

    // --- NUEVO CÓDIGO GSAP ---
    // Seleccionamos los ítems recién renderizados y los animamos
    gsap.fromTo(list.querySelectorAll('.commune-item'), 
        { opacity: 0, x: -15 }, 
        { 
            opacity: 1, 
            x: 0, 
            duration: 0.35, 
            stagger: 0.03, // Cada ítem entra con 0.03s de diferencia
            ease: "power2.out",
            clearProps: "all" // Limpia los estilos inline al terminar para no romper CSS futuros
        }
    );
}


function sortBy(key, el) {
    currentSort = key;
    document.querySelectorAll('#sortTabs .ftab').forEach(t => t.classList.remove('active'));
    el.classList.add('active');
    renderList();
}

/* ─────────────────────────────────────────────────────────────────────────────
   MINI GRÁFICO DE EVOLUCIÓN MENSUAL (panel de detalle)

   Muestra, para la comuna seleccionada, la serie completa de atenciones por
   mes/año — TODOS los periodos disponibles — respetando los filtros de
   tipo/proyecto/servicio pero IGNORANDO el filtro de año/mes activo. La idea
   es precisamente poder ver si un total acumulado corresponde a varios meses
   normales o a un pico puntual (ver conversación sobre CVG y cupos).
   ───────────────────────────────────────────────────────────────────────────── */

const MES_SHORT = {
    ENERO: 'Ene', FEBRERO: 'Feb', MARZO: 'Mar', ABRIL: 'Abr', MAYO: 'May', JUNIO: 'Jun',
    JULIO: 'Jul', AGOSTO: 'Ago', SEPTIEMBRE: 'Sep', OCTUBRE: 'Oct', NOVIEMBRE: 'Nov', DICIEMBRE: 'Dic'
};

function getMonthlySeriesForCommune(idComuna) {
    const svcFilter = document.getElementById('servicioFilter').value.trim();
    const subProjFilter = document.getElementById('subProyectoFilter').value;

    const sets = [];
    if (filtroTipoProyecto !== 'GENERAL') sets.push(_indexByTipo[filtroTipoProyecto] || new Set());
    if (svcFilter) sets.push(_indexBySvc[svcFilter] || new Set());
    if (subProjFilter) sets.push(_indexByCod[subProjFilter] || new Set());

    const base = sets.length === 0 ? _indexAll
        : sets.length === 1 ? sets[0]
        : intersectSets(sets);

    const counts = {};
    base.forEach(i => {
        const r = RAW_ROWS[i];
        if (r.__COMUNAID !== idComuna || !r.__MES) return;
        const key = r.__YEAR + '|' + r.__MES;
        counts[key] = (counts[key] || 0) + 1;
    });

    return Object.keys(counts)
        .sort((a, b) => {
            const [ya, ma] = a.split('|');
            const [yb, mb] = b.split('|');
            if (ya !== yb) return ya - yb;
            return MESES_ORDER.indexOf(ma) - MESES_ORDER.indexOf(mb);
        })
        .map(k => {
            const [y, m] = k.split('|');
            return { year: parseInt(y), label: (MES_SHORT[m] || m.slice(0, 3)) + " '" + y.slice(2), value: counts[k] };
        });
}

function buildMiniLineChart(idComuna) {
    const data = getMonthlySeriesForCommune(idComuna);
    if (data.length === 0) {
        return '<div class="mc-empty">Sin datos mensuales para graficar con los filtros actuales</div>';
    }

    const W = 284, H = 84, padL = 6, padR = 6, padT = 8, padB = 16;
    const innerW = W - padL - padR;
    const innerH = H - padT - padB;
    const maxV = Math.max(...data.map(d => d.value), 1);
    const n = data.length;
    const stepX = n > 1 ? innerW / (n - 1) : 0;

    const points = data.map((d, i) => ({
        x: padL + (n > 1 ? i * stepX : innerW / 2),
        y: padT + innerH - (d.value / maxV) * innerH,
        ...d
    }));

    const poly = points.map(p => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
    const showAllLabels = n <= 7;

    const dotsAndLabels = points.map((p, i) => {
        const showLabel = showAllLabels || i === 0 || i === n - 1 || i % 2 === 0;
        return `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="2.6" class="mc-dot"><title>${p.label}: ${p.value.toLocaleString('es-CO')} atenciones</title></circle>`
             + (showLabel ? `<text x="${p.x.toFixed(1)}" y="${H - 4}" class="mc-label" text-anchor="middle">${p.label}</text>` : '');
    }).join('');

    return `<svg class="mini-chart" viewBox="0 0 ${W} ${H}">
        <polyline points="${poly}" class="mc-line"></polyline>
        ${dotsAndLabels}
    </svg>
    <div class="mc-hint">🔍 Toca para ver en detalle</div>`;
}

/* ─────────────────────────────────────────────────────────────────────────────
   MODAL DE GRÁFICO DETALLADO
   Versión ampliada del mismo gráfico, pensada para que se entienda sin
   necesidad de explicación: valores visibles sobre cada punto (no solo al
   pasar el mouse), ejes con escala, y una frase en lenguaje natural con el
   hallazgo (mínimo/máximo/promedio) antes del gráfico.
   ───────────────────────────────────────────────────────────────────────────── */

function describeActiveFilters() {
    const tipo = filtroTipoProyecto === 'GENERAL' ? null : filtroTipoProyecto;
    const svc = document.getElementById('servicioFilter').value.trim();
    const proj = document.getElementById('subProyectoFilter').value;
    const partes = [];
    if (tipo) partes.push(tipo === 'RO' ? 'Recursos Ordinarios (RO)' : 'Presupuesto Participativo (PP)');
    if (proj) partes.push(`proyecto ${proj}`);
    if (svc) partes.push(`servicio "${svc}"`);
    return partes.length ? partes.join(' · ') : 'todos los proyectos y servicios';
}

function buildChartSummary(data) {
    if (data.length === 0) {
        return 'No hay atenciones registradas para esta comuna con los filtros activos, en ningún mes disponible.';
    }
    if (data.length === 1) {
        return `En <strong>${data[0].label}</strong> se registraron <strong>${data[0].value.toLocaleString('es-CO')}</strong> atenciones — es el único periodo con datos para esta combinación de filtros.`;
    }
    const values = data.map(d => d.value);
    const maxV = Math.max(...values);
    const minV = Math.min(...values);
    const avgV = Math.round(values.reduce((a, b) => a + b, 0) / values.length);
    const maxLabel = data[values.indexOf(maxV)].label;
    const minLabel = data[values.indexOf(minV)].label;
    return `Entre <strong>${data[0].label}</strong> y <strong>${data[data.length - 1].label}</strong>, el valor mensual osciló entre <strong>${minV.toLocaleString('es-CO')}</strong> (${minLabel}) y <strong>${maxV.toLocaleString('es-CO')}</strong> (${maxLabel}), con un promedio de <strong>${avgV.toLocaleString('es-CO')}</strong> atenciones por mes.`;
}

function buildDetailedLineChart(data) {
    if (data.length === 0) {
        return '<div class="mc-empty" style="padding:36px 0;">Sin datos mensuales para graficar con los filtros actuales</div>';
    }

    const values = data.map(d => d.value);
    const maxV = Math.max(...values, 1);
    const minV = Math.min(...values);
    const maxIdx = values.indexOf(maxV);
    const minIdx = values.indexOf(minV);
    const n = data.length;

    // Con series largas (varios años combinados) no caben todas las
    // etiquetas de valor ni de mes sin encimarse — se recortan y se rota
    // el texto del eje X. Con series cortas (el caso típico: un año) todo
    // se muestra igual que antes.
    const showAllValueLabels = n <= 12;
    const rotateXLabels = n > 8;
    const maxXLabels = 9;
    const labelInterval = Math.max(1, Math.ceil(n / maxXLabels));

    const W = 640, H = 300, padL = 48, padR = 16, padT = 26;
    const padB = rotateXLabels ? 56 : 40;
    const innerW = W - padL - padR;
    const innerH = H - padT - padB;
    const stepX = n > 1 ? innerW / (n - 1) : 0;

    const points = data.map((d, i) => ({
        x: padL + (n > 1 ? i * stepX : innerW / 2),
        y: padT + innerH - (d.value / maxV) * innerH,
        ...d
    }));

    // Separadores y etiqueta de año — solo si la serie cruza más de un año,
    // para dejar claro de un vistazo que se está mirando varios periodos.
    const years = [...new Set(data.map(d => d.year))];
    let yearBands = '';
    if (years.length > 1) {
        years.forEach(yr => {
            const idxs = [];
            points.forEach((p, i) => { if (p.year === yr) idxs.push(i); });
            const first = points[idxs[0]];
            const last = points[idxs[idxs.length - 1]];
            const midX = (first.x + last.x) / 2;
            yearBands += `<text x="${midX.toFixed(1)}" y="${(padT - 10).toFixed(1)}" class="mcd-year-label" text-anchor="middle">${yr}</text>`;
            if (idxs[0] > 0) {
                const sepX = ((points[idxs[0] - 1].x + first.x) / 2).toFixed(1);
                yearBands += `<line x1="${sepX}" y1="${padT}" x2="${sepX}" y2="${(padT + innerH).toFixed(1)}" class="mcd-year-sep"></line>`;
            }
        });
    }

    // Líneas de referencia horizontales con su valor
    const gridSteps = 4;
    let gridLines = '';
    for (let g = 0; g <= gridSteps; g++) {
        const frac = g / gridSteps;
        const y = padT + innerH - frac * innerH;
        const val = Math.round(frac * maxV);
        gridLines += `<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W - padR}" y2="${y.toFixed(1)}" class="mcd-grid"></line>`;
        gridLines += `<text x="${(padL - 8).toFixed(1)}" y="${(y + 3).toFixed(1)}" class="mcd-axis-y" text-anchor="end">${val.toLocaleString('es-CO')}</text>`;
    }

    const poly = points.map(p => `${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ');
    const baseY = (padT + innerH).toFixed(1);
    const areaPath = `M${points[0].x.toFixed(1)},${baseY} `
        + points.map(p => `L${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' ')
        + ` L${points[n - 1].x.toFixed(1)},${baseY} Z`;

    const dotsLabelsAxis = points.map((p, i) => {
        const isExtreme = i === maxIdx || (i === minIdx && minIdx !== maxIdx);
        const dotClass = isExtreme ? 'mcd-dot mcd-dot-highlight' : 'mcd-dot';
        const r = isExtreme ? 4.5 : (n > 16 ? 2.4 : 3.5);

        // El valor exacto siempre está disponible al pasar el mouse/tocar.
        let out = `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${r}" class="${dotClass}"><title>${p.label}: ${p.value.toLocaleString('es-CO')} atenciones</title></circle>`;

        // El valor escrito solo se muestra siempre en series cortas; en
        // series largas se reserva a los puntos máximo/mínimo para no
        // encimar texto (como pasaba antes).
        if (showAllValueLabels) {
            out += `<text x="${p.x.toFixed(1)}" y="${(p.y - 11).toFixed(1)}" class="mcd-value" text-anchor="middle">${p.value.toLocaleString('es-CO')}</text>`;
        } else if (isExtreme) {
            out += `<text x="${p.x.toFixed(1)}" y="${(p.y - 9).toFixed(1)}" class="mcd-value mcd-value-small" text-anchor="middle">${p.value.toLocaleString('es-CO')}</text>`;
        }

        // Etiqueta de mes: solo cada N puntos (según cuántos quepan) + el
        // último siempre, para que nunca se encimen entre sí.
        if (i % labelInterval === 0 || i === n - 1) {
            if (rotateXLabels) {
                const ly = H - padB + 12;
                out += `<text x="${p.x.toFixed(1)}" y="${ly}" class="mcd-axis-x" text-anchor="end" transform="rotate(-40 ${p.x.toFixed(1)} ${ly})">${p.label}</text>`;
            } else {
                out += `<text x="${p.x.toFixed(1)}" y="${H - padB + 20}" class="mcd-axis-x" text-anchor="middle">${p.label}</text>`;
            }
        }
        return out;
    }).join('');

    return `<svg class="detailed-chart" viewBox="0 0 ${W} ${H}">
        <defs>
            <linearGradient id="mcdFill" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" class="mcd-stop-start"></stop>
                <stop offset="100%" class="mcd-stop-end"></stop>
            </linearGradient>
        </defs>
        ${gridLines}
        ${yearBands}
        <path d="${areaPath}" fill="url(#mcdFill)"></path>
        <polyline points="${poly}" class="mcd-line"></polyline>
        ${dotsLabelsAxis}
    </svg>`;
}

function openChartModal(idComuna) {
    const data = getMonthlySeriesForCommune(idComuna);
    const nombre = (communeData[idComuna] && communeData[idComuna].nameStr) || `Comuna ${idComuna}`;

    document.getElementById('chartModalTitle').textContent = `Evolución mensual — ${nombre}`;
    document.getElementById('chartModalFilters').textContent = `Filtros activos: ${describeActiveFilters()}.`;
    document.getElementById('chartModalSummary').innerHTML = buildChartSummary(data);
    document.getElementById('chartModalBody').innerHTML = buildDetailedLineChart(data);
    document.getElementById('chartModalNote').textContent = 'Cada punto cuenta atenciones registradas ese mes (filas de la base), no personas únicas — una persona atendida varios meses aparece una vez por cada mes. Se muestran todos los años y meses disponibles para esta comuna, sin importar el filtro de mes/año seleccionado en el panel principal.';

    document.getElementById('chartModalOverlay').classList.add('visible');
    document.body.style.overflow = 'hidden';
}

function closeChartModal() {
    document.getElementById('chartModalOverlay').classList.remove('visible');
    document.body.style.overflow = '';
}

document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeChartModal();
});

/* ─────────────────────────────────────────────────────────────────────────────
   PANEL DE DETALLE
   ───────────────────────────────────────────────────────────────────────────── */

function showNoDataDetail(idComuna) {
    activeCommune = idComuna;
    document.getElementById('dp-title').textContent = `Comuna ${idComuna}`;
    document.getElementById('dp-empty').textContent = 'Esta comuna no priorizó recursos con los filtros actuales';
    document.getElementById('dp-empty').style.display = 'block';
    document.getElementById('dp-total').textContent = '0';
    document.getElementById('dp-edad').textContent = '—';
    document.getElementById('dp-muj').textContent = '0';
    document.getElementById('dp-hom').textContent = '0';
    document.getElementById('dp-muj-pct').textContent = '—';
    document.getElementById('dp-hom-pct').textContent = '—';
    document.getElementById('dp-muj-bar').style.width = '0%';
    document.getElementById('dp-hom-bar').style.width = '0%';
    document.getElementById('dp-services').innerHTML = '';
    document.getElementById('dp-discap').innerHTML = '';
    document.getElementById('dp-monthly-chart').innerHTML = buildMiniLineChart(idComuna);
    document.getElementById('dp-monthly-chart').onclick = () => openChartModal(idComuna);
    document.getElementById('detailPanel').classList.add('visible');
    _highlightCommune(idComuna);
    _syncListActive(idComuna);
    syncUrlState();
}

function showDetail(idComuna) {
    const d = communeData[idComuna];
    if (!d) return;
    activeCommune = idComuna;
    document.getElementById('dp-empty').style.display = 'none';
    document.getElementById('dp-title').textContent = d.nameStr;
    document.getElementById('dp-total').textContent = d.total.toLocaleString('es-CO');
    document.getElementById('dp-edad').textContent = d.edad_promedio + ' a';
    document.getElementById('dp-muj').textContent = d.mujeres.toLocaleString('es-CO');
    document.getElementById('dp-hom').textContent = d.hombres.toLocaleString('es-CO');

    const pctMuj = d.total > 0 ? Math.round(d.mujeres / d.total * 100) : 0;
    const pctHom = d.total > 0 ? Math.round(d.hombres / d.total * 100) : 0;
    document.getElementById('dp-muj-pct').textContent = `${d.mujeres.toLocaleString('es-CO')} (${pctMuj}%)`;
    document.getElementById('dp-hom-pct').textContent = `${d.hombres.toLocaleString('es-CO')} (${pctHom}%)`;
    setTimeout(() => {
        document.getElementById('dp-muj-bar').style.width = pctMuj + '%';
        document.getElementById('dp-hom-bar').style.width = pctHom + '%';
    }, 50);

    const svcContainer = document.getElementById('dp-services');
    // Reusar innerHTML solo si cambió la comuna (evita reflow innecesario)
    const sortedSvc = Object.entries(d.servicios).sort((a, b) => b[1] - a[1]);
    svcContainer.innerHTML = sortedSvc.map(([svc, cnt], i) => {
        const pct = Math.round(cnt / d.total * 100);
        const color = SERVICE_COLORS[i % SERVICE_COLORS.length];
        return `<div class="dp-service-item">
            <div class="dp-service-dot" style="background:${color}"></div>
            <div class="dp-service-name">${svc.slice(0, 35)}${svc.length > 35 ? '…' : ''}</div>
            <div class="dp-service-count">${cnt.toLocaleString('es-CO')} (${pct}%)</div>
        </div>`;
    }).join('');

    const discapEntries = Object.entries(d.discapacidades).sort((a, b) => b[1] - a[1]).slice(0, 3);
    const etniaEntries  = Object.entries(d.etnias).sort((a, b) => b[1] - a[1]).slice(0, 2);
    let discapHTML = '';
    if (discapEntries.length) discapHTML += discapEntries.map(([k, v]) => `<span style="color:var(--accent4)">${k}</span>: ${v.toLocaleString('es-CO')}`).join(' · ');
    if (etniaEntries.length) discapHTML += (discapHTML ? '<br>' : '') + etniaEntries.map(([k, v]) => `<span style="color:var(--accent5)">${k}</span>: ${v.toLocaleString('es-CO')}`).join(' · ');
    if (!discapHTML) discapHTML = '<span style="color:var(--text-muted)">Sin datos de discapacidad/etnia</span>';
    document.getElementById('dp-discap').innerHTML = discapHTML;

    document.getElementById('dp-monthly-chart').innerHTML = buildMiniLineChart(idComuna);
    document.getElementById('dp-monthly-chart').onclick = () => openChartModal(idComuna);

    document.getElementById('detailPanel').classList.add('visible');
    _highlightCommune(idComuna);
    _syncListActive(idComuna);
    syncUrlState();
}

// Resalta la capa activa en el mapa de forma eficiente
function _highlightCommune(idComuna) {
    if (!geoLayer) return;
    geoLayer.eachLayer(layer => {
        const numGeo = parseInt(layer.feature.properties.CODIGO);
        const isActive = numGeo === idComuna;
        const dd = communeData[numGeo];
        const total = dd ? dd.total : 0;
        layer.setStyle({
            weight:      isActive ? 2.5 : 1,
            color:       isActive ? '#00d4aa' : '#0d1520',
            fillOpacity: isActive ? 0.85 : (total > 0 ? 0.68 : 0.12)
        });
        if (isActive) layer.bringToFront();
    });
}

function _syncListActive(idComuna) {
    document.querySelectorAll('.commune-item').forEach(el => {
        el.classList.toggle('active', parseInt(el.dataset.id) === idComuna);
    });
}

function closeDetail() {
    activeCommune = null;
    document.getElementById('detailPanel').classList.remove('visible');
    if (geoLayer) geoLayer.setStyle(feature => styleFeature(feature));
    document.querySelectorAll('.commune-item').forEach(el => el.classList.remove('active'));
    syncUrlState();
}

/* ─────────────────────────────────────────────────────────────────────────────
   TOOLTIPS INFORMATIVOS (ⓘ)
   Un solo bubble compartido, reposicionado por JS sobre cada .info-icon.
   Funciona con mouse (hover) y con touch/teclado (focus/click), y se
   recalcula en cada apertura para no quedar recortado por el viewport.
   ───────────────────────────────────────────────────────────────────────────── */

function initInfoTips() {
    const bubble = document.createElement('div');
    bubble.className = 'tip-bubble';
    document.body.appendChild(bubble);

    let activeIcon = null;

    function place(icon) {
        bubble.textContent = icon.dataset.tip;
        bubble.classList.add('visible');
        const r = icon.getBoundingClientRect();
        const bw = bubble.offsetWidth;
        const bh = bubble.offsetHeight;
        let left = r.left + r.width / 2 - bw / 2;
        left = Math.max(8, Math.min(left, window.innerWidth - bw - 8));
        let top = r.top - bh - 8;
        if (top < 8) top = r.bottom + 8; // si no cabe arriba, se abre hacia abajo
        bubble.style.left = left + 'px';
        bubble.style.top = top + 'px';
    }

    function hide() {
        bubble.classList.remove('visible');
        if (activeIcon) activeIcon.classList.remove('active');
        activeIcon = null;
    }

    document.querySelectorAll('.info-icon[data-tip]').forEach(icon => {
        icon.setAttribute('tabindex', '0');
        icon.addEventListener('mouseenter', () => { activeIcon = icon; icon.classList.add('active'); place(icon); });
        icon.addEventListener('mouseleave', hide);
        icon.addEventListener('focus', () => { activeIcon = icon; icon.classList.add('active'); place(icon); });
        icon.addEventListener('blur', hide);
        icon.addEventListener('click', e => {
            e.stopPropagation();
            if (activeIcon === icon) { hide(); return; }
            activeIcon = icon; icon.classList.add('active'); place(icon);
        });
    });

    document.addEventListener('click', hide);
    window.addEventListener('scroll', hide, true);
    window.addEventListener('resize', hide);
}

/* ─────────────────────────────────────────────────────────────────────────────
   EXPORTAR COMUNA (PNG / PDF)
   Construye una tarjeta de reporte con tema claro (reutilizando los mismos
   datos y el mismo SVG de evolución mensual — solo cambian las variables de
   color que .export-card redefine) y la saca como imagen (html2canvas,
   cargado perezosamente solo si se usa) o como PDF vía el diálogo de
   impresión nativo del navegador (sin librería extra: "Guardar como PDF" ya
   viene integrado en Chrome/Edge/Firefox).
   ───────────────────────────────────────────────────────────────────────────── */

function buildExportCardHTML(idComuna) {
    const d = communeData[idComuna];
    const nombre = (d && d.nameStr) || `Comuna ${idComuna}`;
    const fecha = new Date().toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' });

    if (!d) {
        return `<div class="export-card">
            <div class="ec-header">
                <div class="ec-title">${nombre}</div>
                <div class="ec-subtitle">Georreferenciación Personas Mayores — Alcaldía de Medellín · ${fecha}</div>
                <div class="ec-filters">Filtros activos: ${describeActiveFilters()}</div>
            </div>
            <div class="ec-empty">Esta comuna no tiene atenciones registradas con los filtros activos.</div>
        </div>`;
    }

    const pctMuj = d.total > 0 ? Math.round(d.mujeres / d.total * 100) : 0;
    const pctHom = d.total > 0 ? Math.round(d.hombres / d.total * 100) : 0;

    const sortedSvc = Object.entries(d.servicios).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const svcRows = sortedSvc.map(([svc, cnt], i) => {
        const pct = Math.round(cnt / d.total * 100);
        const color = SERVICE_COLORS[i % SERVICE_COLORS.length];
        return `<div class="ec-svc-row">
            <div class="ec-svc-dot" style="background:${color}"></div>
            <div class="ec-svc-name">${svc}</div>
            <div class="ec-svc-count">${cnt.toLocaleString('es-CO')} (${pct}%)</div>
        </div>`;
    }).join('');

    const serie = getMonthlySeriesForCommune(idComuna);
    const chartSvg = buildDetailedLineChart(serie);
    const summary = buildChartSummary(serie);

    return `<div class="export-card">
        <div class="ec-header">
            <div class="ec-title">${nombre}</div>
            <div class="ec-subtitle">Georreferenciación Personas Mayores — Alcaldía de Medellín · Generado el ${fecha}</div>
            <div class="ec-filters">Filtros activos: ${describeActiveFilters()}</div>
        </div>

        <div class="ec-kpis">
            <div class="ec-kpi"><div class="ec-kpi-val">${d.total.toLocaleString('es-CO')}</div><div class="ec-kpi-lbl">Total atenciones</div></div>
            <div class="ec-kpi"><div class="ec-kpi-val">${d.edad_promedio || '—'} a</div><div class="ec-kpi-lbl">Edad promedio</div></div>
            <div class="ec-kpi"><div class="ec-kpi-val">${d.mujeres.toLocaleString('es-CO')} (${pctMuj}%)</div><div class="ec-kpi-lbl">Mujeres</div></div>
            <div class="ec-kpi"><div class="ec-kpi-val">${d.hombres.toLocaleString('es-CO')} (${pctHom}%)</div><div class="ec-kpi-lbl">Hombres</div></div>
        </div>

        <div class="ec-section-title">Evolución mensual</div>
        <div class="ec-summary">${summary}</div>
        <div class="ec-chart">${chartSvg}</div>

        <div class="ec-section-title">Servicios</div>
        <div class="ec-svc-list">${svcRows}</div>

        <div class="ec-footer">Cada cifra cuenta atenciones registradas (filas de la base), no personas únicas — una persona atendida varios meses aparece una vez por cada mes reportado.</div>
    </div>`;
}

function exportComunaPDF(idComuna) {
    if (!idComuna) return;
    const container = document.getElementById('exportContainer');
    container.innerHTML = buildExportCardHTML(idComuna);
    // Pequeño respiro para que el navegador termine de pintar el SVG antes
    // de abrir el diálogo de impresión.
    setTimeout(() => window.print(), 80);
}

let _html2canvasPromise = null;
function loadHtml2Canvas() {
    if (window.html2canvas) return Promise.resolve();
    if (_html2canvasPromise) return _html2canvasPromise;
    _html2canvasPromise = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
    return _html2canvasPromise;
}

async function exportComunaPNG(idComuna) {
    if (!idComuna) return;
    const btn = document.getElementById('exportPngBtn');
    const original = btn ? btn.innerHTML : null;
    if (btn) { btn.disabled = true; btn.innerHTML = '⏳ Generando…'; }

    try {
        await loadHtml2Canvas();
        const container = document.getElementById('exportContainer');
        container.innerHTML = buildExportCardHTML(idComuna);
        // Deja que el navegador pinte el SVG antes de capturarlo
        await new Promise(r => setTimeout(r, 60));

        const card = container.querySelector('.export-card');
        const canvas = await html2canvas(card, { backgroundColor: '#ffffff', scale: 2 });
        canvas.toBlob(blob => {
            const nombre = (communeData[idComuna] && communeData[idComuna].nameStr) || `comuna-${idComuna}`;
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `${nombre.replace(/[^\w-]+/g, '_')}.png`;
            a.click();
            URL.revokeObjectURL(a.href);
        });
    } catch (err) {
        console.error('Error exportando PNG:', err);
        alert('No se pudo generar el PNG. Puedes usar la opción PDF como alternativa.');
    } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = original; }
    }
}

/* ─────────────────────────────────────────────────────────────────────────────
   COMMAND PALETTE (Ctrl/Cmd+K)
   Búsqueda rápida de comunas, servicios y proyectos sin pasar por el
   buscador lateral. Ctrl+K / Cmd+K desde cualquier parte para abrir, Escape
   o clic afuera para cerrar, flechas + Enter para navegar y seleccionar.
   ───────────────────────────────────────────────────────────────────────────── */

let _cmdkItems = [];
let _cmdkSelected = -1;

function openCmdk() {
    document.getElementById('cmdkOverlay').classList.add('visible');
    const input = document.getElementById('cmdkInput');
    input.value = '';
    renderCmdkResults('');
    setTimeout(() => input.focus(), 20);
}

function closeCmdk() {
    document.getElementById('cmdkOverlay').classList.remove('visible');
}

function cmdkResultsSource() {
    const comunas = Object.values(communeData)
        .sort((a, b) => a.id - b.id)
        .map(d => ({ type: 'comuna', id: d.id, label: d.nameStr, meta: `${d.total.toLocaleString('es-CO')} atenciones` }));

    const servicios = Object.keys(_indexBySvc)
        .sort((a, b) => _indexBySvc[b].size - _indexBySvc[a].size)
        .map(s => ({ type: 'servicio', label: s, meta: `${_indexBySvc[s].size.toLocaleString('es-CO')}` }));

    const proyectos = Object.keys(_indexByCod)
        .sort((a, b) => _indexByCod[b].size - _indexByCod[a].size)
        .map(cod => ({
            type: 'proyecto',
            cod,
            label: `Proyecto ${cod}`,
            meta: PROJECT_METADATA[cod] ? PROJECT_METADATA[cod].tipo : ''
        }));

    return { comunas, servicios, proyectos };
}

function cmdkItemHTML(item, idx) {
    const icon = item.type === 'comuna' ? '📍' : item.type === 'servicio' ? '🏷' : '📁';
    return `<div class="cmdk-item" data-idx="${idx}" onclick="selectCmdkItem(${idx})">
        <span class="cmdk-item-icon">${icon}</span>
        <span class="cmdk-item-label">${item.label}</span>
        <span class="cmdk-item-meta">${item.meta || ''}</span>
    </div>`;
}

function renderCmdkResults(query) {
    const q = query.trim().toLowerCase();
    const { comunas, servicios, proyectos } = cmdkResultsSource();

    const matchC = comunas.filter(x => x.label.toLowerCase().includes(q)).slice(0, 8);
    const matchS = (q ? servicios.filter(x => x.label.toLowerCase().includes(q)) : servicios).slice(0, q ? 8 : 5);
    const matchP = (q ? proyectos.filter(x => x.label.toLowerCase().includes(q) || x.cod.includes(q)) : proyectos).slice(0, 6);

    _cmdkItems = [...matchC, ...matchS, ...matchP];
    _cmdkSelected = _cmdkItems.length ? 0 : -1;

    const results = document.getElementById('cmdkResults');
    if (_cmdkItems.length === 0) {
        results.innerHTML = '<div class="cmdk-empty">Sin resultados</div>';
        return;
    }

    let html = '';
    if (matchC.length) {
        html += '<div class="cmdk-group-label">Comunas</div>';
        matchC.forEach(item => { html += cmdkItemHTML(item, _cmdkItems.indexOf(item)); });
    }
    if (matchS.length) {
        html += '<div class="cmdk-group-label">Servicios</div>';
        matchS.forEach(item => { html += cmdkItemHTML(item, _cmdkItems.indexOf(item)); });
    }
    if (matchP.length) {
        html += '<div class="cmdk-group-label">Proyectos</div>';
        matchP.forEach(item => { html += cmdkItemHTML(item, _cmdkItems.indexOf(item)); });
    }
    results.innerHTML = html;
    highlightCmdkSelection();
}

function highlightCmdkSelection() {
    document.querySelectorAll('.cmdk-item').forEach(el => {
        el.classList.toggle('selected', parseInt(el.dataset.idx) === _cmdkSelected);
    });
    const sel = document.querySelector('.cmdk-item.selected');
    if (sel) sel.scrollIntoView({ block: 'nearest' });
}

function selectCmdkItem(idx) {
    const item = _cmdkItems[idx];
    if (!item) return;
    closeCmdk();

    if (item.type === 'comuna') {
        showDetail(item.id);
        const layer = layersMap[item.id];
        if (layer) map.flyToBounds(layer.getBounds(), { maxZoom: 14, duration: 0.8 });
    } else if (item.type === 'servicio') {
        document.getElementById('servicioFilter').value = item.label;
        applyFilters();
    } else if (item.type === 'proyecto') {
        document.getElementById('subProyectoFilter').value = item.cod;
        applyFilters();
    }
}

document.getElementById('cmdkInput').addEventListener('input', e => renderCmdkResults(e.target.value));

document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        const overlay = document.getElementById('cmdkOverlay');
        overlay.classList.contains('visible') ? closeCmdk() : openCmdk();
        return;
    }
    if (!document.getElementById('cmdkOverlay').classList.contains('visible')) return;
    if (e.key === 'Escape') { closeCmdk(); return; }
    if (e.key === 'ArrowDown') {
        e.preventDefault();
        if (_cmdkSelected < _cmdkItems.length - 1) _cmdkSelected++;
        highlightCmdkSelection();
    } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        if (_cmdkSelected > 0) _cmdkSelected--;
        highlightCmdkSelection();
    } else if (e.key === 'Enter') {
        e.preventDefault();
        if (_cmdkSelected >= 0) selectCmdkItem(_cmdkSelected);
    }
});

/* ─────────────────────────────────────────────────────────────────────────────
   INICIO
   ───────────────────────────────────────────────────────────────────────────── */

initInfoTips();
cargarExcel();
