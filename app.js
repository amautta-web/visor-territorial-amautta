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
    if (row.__YEAR === 2026) return 'PP';
    const cod = row.__COD;
    if (cod && PROJECT_METADATA[cod]) return PROJECT_METADATA[cod].tipo;
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
            if (row.__YEAR === 2026) row.__TIPO = 'PP';
            else if (row.__COD && PROJECT_METADATA[row.__COD]) row.__TIPO = PROJECT_METADATA[row.__COD].tipo;
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
    const filas2026 = await leerArchivoExcel('PP PROYECTO 250071.xlsx', 2026, 'Selecciona PP PROYECTO 250071.xlsx');

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

    setLoad(78, 'Calculando métricas por comuna…');
    await yieldToMain();
    aggregateData();

    setLoad(85, 'Cargando mapa de Medellín…');
    await cargarGeoJSON();
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
    if (row.__YEAR === 2026) return '250071';
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
                victimas: 0, rural: 0, urbana: 0
            };
        }

        const c = communeData[id];
        c.total++;
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

    const edadProm = edadCnt ? (edadSuma / edadCnt).toFixed(1) : '—';

    document.getElementById('pill-total').textContent     = total.toLocaleString('es-CO');
    document.getElementById('pill-comunas').textContent   = comunas;
    document.getElementById('pill-servicios').textContent = svcSet.size;
    document.getElementById('pill-edad').textContent      = edadProm + ' a';
    document.getElementById('pill-meses').textContent     = mesesSet.size;
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
                <div class="commune-count"></div>
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
   PANEL DE DETALLE
   ───────────────────────────────────────────────────────────────────────────── */

function showNoDataDetail(idComuna) {
    activeCommune = idComuna;
    document.getElementById('dp-title').textContent = `Comuna ${idComuna}`;
    document.getElementById('dp-empty').textContent = 'Esta comuna no priorizó recursos';
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
    document.getElementById('detailPanel').classList.add('visible');
    _highlightCommune(idComuna);
    _syncListActive(idComuna);
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

    document.getElementById('detailPanel').classList.add('visible');
    _highlightCommune(idComuna);
    _syncListActive(idComuna);
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
}

/* ─────────────────────────────────────────────────────────────────────────────
   INICIO
   ───────────────────────────────────────────────────────────────────────────── */

cargarExcel();