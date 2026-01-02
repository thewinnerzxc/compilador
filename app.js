// --- Estado global ---
let allData = [];
let headers = [];
let currentSort = { column: null, direction: 'asc' };
let columnFilters = {};
let globalFilter = '';
let highlightTerm = 'uptodate';

let pageSize = 100;
let currentPage = 1;
let currentFiles = [];

// selección de celdas
let selectedCells = new Map();   // key "row-col" -> {rowIndex, colIndex, text}
let isSelecting = false;
let selectStart = null;

// pan con click derecho
let isPanning = false;
let panStartX = 0;
let panStartY = 0;
let panScrollLeft = 0;
let panScrollTop = 0;

// DOM
const folderInput = document.getElementById('folderInput');
const statusEl = document.getElementById('status');
const tableWrapper = document.getElementById('table-wrapper');
const dataTable = document.getElementById('dataTable');
const controls = document.getElementById('controls');
const globalSearchInput = document.getElementById('globalSearch');
const clearGlobalBtn = document.getElementById('clearGlobalBtn');
const rowCountEl = document.getElementById('rowCount');
const highlightInput = document.getElementById('highlightTerm');
const clearFiltersBtn = document.getElementById('clearFiltersBtn');

const prevPageBtn = document.getElementById('prevPage');
const nextPageBtn = document.getElementById('nextPage');
const pageInfoEl = document.getElementById('pageInfo');
const pageSizeSelect = document.getElementById('pageSizeSelect');

const exportExcelBtn = document.getElementById('exportExcel');
const exportCSVBtn = document.getElementById('exportCSV');
const refreshBtn = document.getElementById('refreshBtn');

const alignUniversBtn = document.getElementById('alignUniversBtn');
const alignSourceBtn = document.getElementById('alignSourceBtn');

// --- Utilidades ---
function debounce(fn, delay) {
    let timer = null;
    return function (...args) {
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}

function excelSerialToDateString(serial) {
    const d = XLSX.SSF.parse_date_code(serial);
    if (!d) return serial;
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const day = d.d;
    const month = months[d.m - 1] || 'Jan';
    const year = d.y;
    return `${day}-${month}-${year}`;
}

function looksLikeExcelDate(num) {
    return typeof num === 'number' && num > 20000 && num < 60000;
}

function escapeHtml(str) {
    return str.replace(/[&<>"']/g, function (c) {
        return ({
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;'
        })[c];
    });
}

// resaltado según highlightTerm
function buildHighlightedHTML(text) {
    const raw = text == null ? '' : text.toString();
    if (!raw) return '';

    const term = highlightTerm.trim();
    if (!term) {
        return escapeHtml(raw);
    }

    const lowerText = raw.toLowerCase();
    const lowerTerm = term.toLowerCase();
    let index = 0;
    const ranges = [];

    while (true) {
        const idx = lowerText.indexOf(lowerTerm, index);
        if (idx === -1) break;
        ranges.push({ start: idx, end: idx + lowerTerm.length });
        index = idx + lowerTerm.length;
    }

    if (!ranges.length) {
        return escapeHtml(raw);
    }

    ranges.sort((a, b) => a.start - b.start);
    const merged = [ranges[0]];
    for (let i = 1; i < ranges.length; i++) {
        const last = merged[merged.length - 1];
        const cur = ranges[i];
        if (cur.start <= last.end) {
            last.end = Math.max(last.end, cur.end);
        } else {
            merged.push(cur);
        }
    }

    const bg = '#c8f7c5';
    const fg = '#005c2e';

    let result = '';
    let pos = 0;
    merged.forEach(r => {
        if (pos < r.start) {
            result += escapeHtml(raw.slice(pos, r.start));
        }
        const chunk = raw.slice(r.start, r.end);
        result += `<span class="highlight-span" style="background-color:${bg};color:${fg};">${escapeHtml(chunk)}</span>`;
        pos = r.end;
    });
    if (pos < raw.length) {
        result += escapeHtml(raw.slice(pos));
    }

    return result;
}

// Helper para parsear fechas formato "D-MMM-YYYY" (e.g. "6-Feb-2026")
// Retorna timestamp (número) o -Infinity si no es válida, para ordenar correctamente.
function parseCustomDate(str) {
    if (!str) return -Infinity;

    // Mapeo meses en inglés (que usa la función excelSerialToDateString)
    const months = {
        'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
        'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
    };

    // Regex simple: DIGITO(S)-LETRAS-DIGITOS
    // E.g. "6-Feb-2026"
    const parts = str.toString().match(/^(\d+)-([A-Za-z]+)-(\d{4})$/);
    if (!parts) {
        // Fallback: intentar parse normal por si acaso viene en otro formato
        const t = Date.parse(str);
        return isNaN(t) ? -Infinity : t;
    }

    const day = parseInt(parts[1], 10);
    const monthStr = parts[2];
    const year = parseInt(parts[3], 10);

    const monthIndex = months[monthStr] !== undefined ? months[monthStr] : -1;

    if (monthIndex === -1) return -Infinity;

    const d = new Date(year, monthIndex, day);
    return d.getTime();
}

// --- Debounced handlers ---
const debouncedGlobalSearch = debounce(handleGlobalSearch, 200);
const debouncedHighlightChange = debounce(handleHighlightChange, 150);

// --- Eventos base ---
folderInput.addEventListener('change', (e) => {
    currentFiles = Array.from(e.target.files);
    processFiles(currentFiles);
});

refreshBtn.addEventListener('click', () => {
    if (!currentFiles || currentFiles.length === 0) {
        alert('Primero selecciona la carpeta con tus archivos de Excel.');
        return;
    }
    processFiles(currentFiles);
});

clearFiltersBtn.addEventListener('click', clearAllFilters);

alignUniversBtn.addEventListener('click', () => scrollToColumnByPrefix('univers'));
alignSourceBtn.addEventListener('click', () => scrollToColumnByPrefix('source'));

// --- Helpers ---
function resetColumnFilters() {
    Object.keys(columnFilters).forEach(col => columnFilters[col] = '');
    const filterInputs = dataTable.querySelectorAll('thead tr:nth-child(2) input');
    filterInputs.forEach(inp => inp.value = '');
}

// limpiar búsqueda global con la X (modificado: pega del portapapeles)
clearGlobalBtn.addEventListener('click', async () => {
    // 1. Limpiar filtros de columnas siempre
    resetColumnFilters();

    try {
        let text = await navigator.clipboard.readText();

        // Lógica inteligente:
        // Si parece número de teléfono (solo dígitos, espacios, +, -, paréntesis), limpiamos todo.
        // Si parece texto (letras), se mantiene tal cual (solo trim).

        // Regex para "solo caracteres de teléfono":
        // ^[\d\s+\-()]*$  -> si machea esto, es candidato a limpieza agresiva.
        // Pero si tiene letras, NO machea.

        if (/^[\d\s+\-()]*$/.test(text) && /\d/.test(text)) {
            // Es un número (tiene dígitos y solo chars de telefono) -> limpiar espacios y +
            text = text.replace(/[+\s\-()]/g, '');
        } else {
            // Es texto o mix -> mantener espacios, solo trim
            text = text.trim();
        }

        globalSearchInput.value = text;
        globalFilter = text; // asignamos directamente para que filtre exacto lo que se ve
    } catch (err) {
        // Fallback
        console.warn('No se pudo acceder al portapapeles:', err);
        globalSearchInput.value = '';
        globalFilter = '';
    }

    currentPage = 1;
    updateTable();
    if (tableWrapper) tableWrapper.scrollLeft = 0;
    globalSearchInput.focus();
});

// --- Handlers lógicos ---
function handleGlobalSearch(e) {
    // AL TIPEAR en búsqueda global, limpiar filtros de columnas automáticamente
    if (e.target.value.length > 0) {
        // Solo si hay algo escrito limpiamos (o siempre? el user dijo "al tipear")
        // Para evitar borrar si solo borras caracteres, mejor lo hacemos siempre que cambie
        // si la intención es "busqueda completa".
        resetColumnFilters();
    }

    globalFilter = e.target.value.trim();
    currentPage = 1;
    updateTable();
    tableWrapper.scrollLeft = 0;
}

function handleHighlightChange(e) {
    highlightTerm = e.target.value;
    updateTable();
}

function handleColumnFilterChange(e) {
    const col = e.target.dataset.column;
    columnFilters[col] = e.target.value.trim();
    currentPage = 1;
    updateTable();
}

function handleSort(column) {
    if (currentSort.column === column) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.column = column;
        currentSort.direction = 'asc';
    }
    currentPage = 1;
    updateSortIndicators();
    updateTable();
}

// 🔘 Limpiar todos los filtros (columnas + búsqueda global)
function clearAllFilters() {
    Object.keys(columnFilters).forEach(col => columnFilters[col] = '');

    const filterInputs = dataTable.querySelectorAll('thead tr:nth-child(2) input');
    filterInputs.forEach(inp => inp.value = '');

    globalFilter = '';
    globalSearchInput.value = '';

    // limpiar selección de celdas
    selectedCells.clear();
    clearVisualSelections();
    isSelecting = false;
    selectStart = null;

    // reset sort
    currentSort = { column: null, direction: 'asc' };
    updateSortIndicators();

    currentPage = 1;
    updateTable();
    updateTable();
    // Al limpiar, alinear a la izquierda
    if (tableWrapper) tableWrapper.scrollLeft = 0;
}

// --- Procesamiento de archivos ---
async function processFiles(files) {
    allData = [];
    headers = [];
    currentSort = { column: null, direction: 'asc' };
    columnFilters = {};
    globalFilter = '';
    dataTable.innerHTML = '';
    tableWrapper.style.display = 'none';
    controls.style.display = 'none';
    rowCountEl.textContent = '';
    currentPage = 1;
    selectedCells.clear();
    clearVisualSelections();
    isSelecting = false;
    selectStart = null;

    if (!files || files.length === 0) {
        statusEl.textContent = 'No se seleccionaron archivos.';
        return;
    }

    const excelFiles = files.filter(f => /\.(xlsx|xls|xlsm)$/i.test(f.name));
    if (excelFiles.length === 0) {
        statusEl.textContent = 'No se encontraron archivos de Excel en la carpeta seleccionada.';
        return;
    }

    statusEl.textContent = `Procesando ${excelFiles.length} archivos de Excel...`;

    for (let file of excelFiles) {
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            const fileTitle = file.name.replace(/\.[^.]+$/, "");

            jsonData.forEach(row => {
                for (const key in row) {
                    const value = row[key];
                    if (looksLikeExcelDate(value)) {
                        row[key] = excelSerialToDateString(value);
                    }
                }
                row['Source'] = fileTitle;

                // Limpieza de Whatsapp
                // Buscamos keys que parezcan whatsapp
                Object.keys(row).forEach(k => {
                    if (k.toLowerCase() === 'whatsapp') {
                        // Quitar + y espacios
                        const val = (row[k] || '').toString();
                        row[k] = val.replace(/[+\s]/g, '');
                    }
                });

                allData.push(row);
            });
        } catch (err) {
            console.error(`Error leyendo ${file.name}`, err);
        }
    }

    if (allData.length === 0) {
        statusEl.textContent = 'No se pudieron leer datos de los archivos.';
        return;
    }

    // --- Lógica de Correlación: Email -> Whatsapp ---
    const emailToWhatsapp = new Map();
    const emailKeysCache = new Map(); // cache para saber qué key tiene el email en cada row

    // Paso 1: Construir mapa
    allData.forEach(row => {
        // Buscar campo email
        let emailKey = Object.keys(row).find(k => /e-?mail/i.test(k));

        // Buscar campo whatsapp
        let waKey = Object.keys(row).find(k => k.toLowerCase() === 'whatsapp');

        if (emailKey && waKey) {
            const email = (row[emailKey] || '').toString().toLowerCase().trim();
            const wa = (row[waKey] || '').toString().trim();
            if (email && wa) {
                // Guardamos en el mapa. Si hay duplicados, el ultimo gana (o el primero? da igual si son consistentes)
                emailToWhatsapp.set(email, wa);
            }
        }
    });

    // Paso 1.5: Incorporar contactos de Supabase (Electron only)
    if (window.electronAPI) {
        try {
            // Mostrar estado temporal
            const oldStatus = statusEl.textContent;
            statusEl.textContent = 'Validando contactos con Supabase...';

            const result = await window.electronAPI.getSupabaseContacts();
            if (result.success && Array.isArray(result.data)) {
                console.log(`Integrando ${result.data.length} contactos de Supabase.`);
                result.data.forEach(c => {
                    const email = (c.email || '').toString().toLowerCase().trim();
                    const wa = (c.whatsapp || '').toString().trim();
                    // Prioridad: Supabase sobrescribe Excel si existe
                    // (Ojo: si Supabase tiene dato antiguo vacío y Excel tiene nuevo, esto borraría?
                    //  Asumimos que Supabase tiene wa válido si viene en el select)
                    if (email && wa) {
                        const cleanWa = wa.replace(/[+\s]/g, '');
                        if (cleanWa) {
                            emailToWhatsapp.set(email, cleanWa);
                        }
                    }
                });
            }
            statusEl.textContent = oldStatus;
        } catch (err) {
            console.error("Error fetching Supabase contacts:", err);
        }
    }


    // Paso 2: Rellenar huecos
    allData.forEach(row => {
        let emailKey = Object.keys(row).find(k => /e-?mail/i.test(k));

        // Asegurarnos que existe columna Whatsapp en el objeto para rellenar
        // Si no existe, podemos crearla si encontramos un match, 
        // pero idealmente deberíamos respetar si la fila tenía la columna o no?
        // Mejor buscamos si ya tiene valor
        let waKey = Object.keys(row).find(k => k.toLowerCase() === 'whatsapp');

        // Si no tiene columna Whatsapp, asignamos 'Whatsapp' o 'WHATSAPP' por defecto?
        // Usaremos 'Whatsapp' capitalizado standard si vamos a insertar.
        if (!waKey) {
            waKey = 'Whatsapp'; // La crearemos
        }

        const currentWa = (row[waKey] || '').toString().trim();

        if (!currentWa && emailKey) {
            const email = (row[emailKey] || '').toString().toLowerCase().trim();
            if (email && emailToWhatsapp.has(email)) {
                row[waKey] = emailToWhatsapp.get(email);
            }
        }
    });

    // Paso 3: Filtrar filas de referencia "Email whatsapp brevo"
    // El usuario pidió que NO aparezcan.
    allData = allData.filter(row => {
        const src = (row['Source'] || '').toLowerCase();
        return !src.includes('email whatsapp brevo');
    });

    // Recalcular headers basado en los datos filtrados y modificados
    const columnSet = new Set();
    allData.forEach(row => {
        Object.keys(row).forEach(key => columnSet.add(key));
    });
    headers = Array.from(columnSet);

    if (headers.includes('Source')) {
        headers = ['Source', ...headers.filter(h => h !== 'Source')];
    }

    // Reordenar: Whatsapp entre Comentarios y Rest.
    // Buscamos si existe la columna Whatsapp (ignorando mayúsculas/minúsculas o asumiendo exacto según Excel)
    // El usuario dijo "Whatsapp"
    const waCol = headers.find(h => h.toLowerCase() === 'whatsapp');
    if (waCol) {
        // Quitamos Whatsapp de donde esté
        let tempHeaders = headers.filter(h => h !== waCol);

        // Buscamos índice de Comentarios
        const idxComentarios = tempHeaders.findIndex(h => h.toLowerCase().includes('comentarios'));

        if (idxComentarios !== -1) {
            // Insertar después de Comentarios
            tempHeaders.splice(idxComentarios + 1, 0, waCol);
            headers = tempHeaders;
        } else {
            // Si no está Comentarios, buscamos Rest.
            const idxRest = tempHeaders.findIndex(h => h.toLowerCase().includes('rest.'));
            if (idxRest !== -1) {
                // Insertar antes de Rest.
                tempHeaders.splice(idxRest, 0, waCol);
                headers = tempHeaders;
            } else {
                // Si no está ninguno, lo dejamos donde estaba o al final? 
                // Mejor ponemos headers como estaba si no encontramos referencias, 
                // pero si queremos forzarlo, podríamos dejarlo.
                // Aquí simplemente si no encuentra referencias, ya se filtró, hay que volverlo a poner
                // al final si se perdió, o reusar logicamente la posición original.
                // Simplemente lo agregamos al final si no se insertó.
                tempHeaders.push(waCol);
                headers = tempHeaders;
            }
        }
    }

    headers.forEach(h => columnFilters[h] = '');

    // FILTRO POR DEFECTO: Pendientes_ = "processing"
    if (headers.includes('Pendientes_')) {
        columnFilters['Pendientes_'] = 'processing';
    }

    // ORDEN POR DEFECTO: Reactivation ASC
    if (headers.includes('Reactivation')) {
        currentSort = { column: 'Reactivation', direction: 'asc' };
    }

    buildTableStructure();

    tableWrapper.style.display = 'block';
    controls.style.display = 'flex';
    globalSearchInput.value = '';
    highlightInput.value = highlightTerm;
    pageSize = parseInt(pageSizeSelect.value, 10) || 100;
    currentPage = 1;

    updateTable();
    // Actualizar inputs de filtros visualmente
    if (columnFilters['Pendientes_']) {
        const pInput = dataTable.querySelector('thead tr:nth-child(2) input[data-column="Pendientes_"]');
        if (pInput) pInput.value = columnFilters['Pendientes_'];
    }

    updateTable();
    // Al cargar, alinear a la izquierda
    if (tableWrapper) tableWrapper.scrollLeft = 0;

    statusEl.textContent = `¡Listo! Se combinaron ${excelFiles.length} archivos y ${allData.length} filas.`;
}

// --- Construcción de tabla ---
function buildTableStructure() {
    dataTable.innerHTML = '';

    const thead = document.createElement('thead');

    const headerRow = document.createElement('tr');
    headers.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        th.dataset.column = col;
        th.classList.add('sortable');

        const indicator = document.createElement('span');
        indicator.className = 'sort-indicator';
        th.appendChild(indicator);

        th.addEventListener('click', () => handleSort(col));

        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    const filterRow = document.createElement('tr');
    headers.forEach(col => {
        const th = document.createElement('th');
        const input = document.createElement('input');
        input.type = 'text';
        input.placeholder = 'Filtrar...';
        input.dataset.column = col;
        input.addEventListener('input', debounce(handleColumnFilterChange, 200));
        th.appendChild(input);
        filterRow.appendChild(th);
    });
    thead.appendChild(filterRow);

    dataTable.appendChild(thead);

    const tbody = document.createElement('tbody');
    dataTable.appendChild(tbody);

    // eventos generales
    globalSearchInput.removeEventListener('input', debouncedGlobalSearch);
    globalSearchInput.addEventListener('input', debouncedGlobalSearch);

    highlightInput.removeEventListener('input', debouncedHighlightChange);
    highlightInput.addEventListener('input', debouncedHighlightChange);

    // Evento paste en búsqueda global: limpiar formato
    // Evento paste en búsqueda global: limpiar formato
    globalSearchInput.addEventListener('paste', (e) => {
        e.preventDefault();
        let text = (e.clipboardData || window.clipboardData).getData('text');

        // Lógica inteligente: Si parece número de teléfono, limpiar. Si es texto, mantener.
        if (/^[\d\s+\-()]*$/.test(text) && /\d/.test(text)) {
            // Es un número -> limpiar espacios y +
            text = text.replace(/[+\s\-()]/g, '');
        } else {
            // Es texto -> mantener espacios internales, solo trim bordes
            // OJO: Si pegamos "Juan Perez", queremos "Juan Perez".
            // Pero si pegamos con saltos de línea, quizás convenga normalizar?
            // "  Juan   Perez  " -> "Juan Perez"?
            // El usuario pidió "que se conserven los espacios" entre nombres.
            // Asumiremos que el clipboard viene bien, solo trim.
            text = text.trim();
        }

        const start = globalSearchInput.selectionStart;
        const end = globalSearchInput.selectionEnd;
        const currentVal = globalSearchInput.value;
        const newVal = currentVal.substring(0, start) + text + currentVal.substring(end);

        globalSearchInput.value = newVal;

        // Actualizar cursor
        globalSearchInput.selectionStart = globalSearchInput.selectionEnd = start + text.length;

        // Disparar evento input para filtrar (esto ahora disparará resetColumnFilters gracias a handleGlobalSearch)
        globalSearchInput.dispatchEvent(new Event('input'));
    });

    prevPageBtn.onclick = () => {
        if (currentPage > 1) {
            currentPage--;
            updateTable();
        }
    };
    nextPageBtn.onclick = () => {
        currentPage++;
        updateTable();
    };
    pageSizeSelect.onchange = () => {
        pageSize = parseInt(pageSizeSelect.value, 10) || 100;
        currentPage = 1;
        updateTable();
    };

    exportExcelBtn.onclick = exportVisibleToExcel;
    exportCSVBtn.onclick = exportVisibleToCSV;

    // selección tipo Excel
    tbody.addEventListener('mousedown', handleTbodyMouseDown);
    tbody.addEventListener('mouseover', handleTbodyMouseOver);
}

// --- Selección de celdas y portapapeles ---
function handleTbodyMouseDown(ev) {
    const td = ev.target.closest('td');
    if (!td) return;

    const tr = td.parentElement;
    if (!tr || tr.parentElement.tagName.toLowerCase() !== 'tbody') return;

    const rowIndex = parseInt(tr.dataset.rowIndex, 10);
    const colIndex = parseInt(td.dataset.colIndex, 10);
    if (isNaN(rowIndex) || isNaN(colIndex)) return;

    // solo botón izquierdo
    if (ev.button !== 0) return;

    const key = `${rowIndex}-${colIndex}`;

    // Ctrl / Cmd + click = toggle single cell
    if (ev.ctrlKey || ev.metaKey) {
        if (selectedCells.has(key)) {
            selectedCells.delete(key);
            td.classList.remove('cell-selected');
        } else {
            selectedCells.set(key, {
                rowIndex,
                colIndex,
                text: td.textContent || ''
            });
            td.classList.add('cell-selected');
        }
        copySelectionToClipboard();
        return;
    }

    // selección rectangular
    isSelecting = true;
    selectStart = { rowIndex, colIndex };
    selectedCells.clear();
    clearVisualSelections();
    applyRectSelection(selectStart, selectStart);

    ev.preventDefault(); // evita selección de texto nativa
}

function handleTbodyMouseOver(ev) {
    if (!isSelecting || !selectStart) return;

    const td = ev.target.closest('td');
    if (!td) return;

    const tr = td.parentElement;
    if (!tr || tr.parentElement.tagName.toLowerCase() !== 'tbody') return;

    const rowIndex = parseInt(tr.dataset.rowIndex, 10);
    const colIndex = parseInt(td.dataset.colIndex, 10);
    if (isNaN(rowIndex) || isNaN(colIndex)) return;

    applyRectSelection(selectStart, { rowIndex, colIndex });
}

document.addEventListener('mouseup', () => {
    if (isSelecting) {
        isSelecting = false;
        selectStart = null;
        copySelectionToClipboard();
    }
});

function applyRectSelection(start, end) {
    const minRow = Math.min(start.rowIndex, end.rowIndex);
    const maxRow = Math.max(start.rowIndex, end.rowIndex);
    const minCol = Math.min(start.colIndex, end.colIndex);
    const maxCol = Math.max(start.colIndex, end.colIndex);

    selectedCells.clear();
    clearVisualSelections();

    const rows = dataTable.querySelectorAll('tbody tr');
    rows.forEach(tr => {
        const rIndex = parseInt(tr.dataset.rowIndex, 10);
        if (isNaN(rIndex)) return;
        if (rIndex < minRow || rIndex > maxRow) return;

        const tds = tr.querySelectorAll('td');
        tds.forEach(td => {
            const cIndex = parseInt(td.dataset.colIndex, 10);
            if (isNaN(cIndex)) return;
            if (cIndex < minCol || cIndex > maxCol) return;

            const key = `${rIndex}-${cIndex}`;
            selectedCells.set(key, {
                rowIndex: rIndex,
                colIndex: cIndex,
                text: td.textContent || ''
            });
            td.classList.add('cell-selected');
        });
    });
}

function clearVisualSelections() {
    const selected = dataTable.querySelectorAll('td.cell-selected');
    selected.forEach(td => td.classList.remove('cell-selected'));
}

function copySelectionToClipboard() {
    if (selectedCells.size === 0) return;

    const cellsArr = Array.from(selectedCells.values());
    cellsArr.sort((a, b) => {
        if (a.rowIndex !== b.rowIndex) return a.rowIndex - b.rowIndex;
        return a.colIndex - b.colIndex;
    });

    const rowsMap = new Map();
    for (const cell of cellsArr) {
        if (!rowsMap.has(cell.rowIndex)) rowsMap.set(cell.rowIndex, []);
        rowsMap.get(cell.rowIndex).push(cell);
    }

    const rowStrings = [];
    for (const [, cells] of rowsMap) {
        cells.sort((a, b) => a.colIndex - b.colIndex);
        const values = cells.map(c => c.text);
        rowStrings.push(values.join('\t'));
    }

    const text = rowStrings.join('\n');

    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).catch(err => {
            console.error('Error copiando al portapapeles:', err);
        });
    } else {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        try {
            document.execCommand('copy');
        } catch (e) {
            console.error('execCommand copy falló:', e);
        }
        document.body.removeChild(textarea);
    }
}

// --- Pan con click derecho en tableWrapper ---
function setupPanning() {
    if (!tableWrapper) return;

    tableWrapper.addEventListener('mousedown', (e) => {
        if (e.button !== 2) return; // solo botón derecho
        isPanning = true;
        panStartX = e.clientX;
        panStartY = e.clientY;
        panScrollLeft = tableWrapper.scrollLeft;
        panScrollTop = tableWrapper.scrollTop;
        tableWrapper.classList.add('panning');
        e.preventDefault();
    });

    tableWrapper.addEventListener('contextmenu', (e) => {
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isPanning) return;
        const dx = e.clientX - panStartX;
        const dy = e.clientY - panStartY;
        tableWrapper.scrollLeft = panScrollLeft - dx;
        tableWrapper.scrollTop = panScrollTop - dy;
    });

    document.addEventListener('mouseup', () => {
        if (isPanning) {
            isPanning = false;
            tableWrapper.classList.remove('panning');
        }
    });
}

setupPanning();

// --- Filtros / orden ---
function updateSortIndicators() {
    const ths = dataTable.querySelectorAll('thead tr:first-child th');
    ths.forEach(th => {
        const span = th.querySelector('.sort-indicator');
        if (span) span.textContent = '';
    });

    if (!currentSort.column) return;

    const activeTh = dataTable.querySelector(
        `thead tr:first-child th[data-column="${CSS.escape(currentSort.column)}"]`
    );
    if (activeTh) {
        const span = activeTh.querySelector('.sort-indicator');
        if (span) {
            span.textContent = currentSort.direction === 'asc' ? ' ▲' : ' ▼';
        }
    }
}

function getFilteredAndSortedData() {
    let filtered = allData.filter(row => {
        for (const [col, filterValue] of Object.entries(columnFilters)) {
            if (!filterValue) continue;

            // Lógica especial para columna "Rest." (rango de colores)
            if (col === 'Rest.') {
                const fVal = filterValue.toLowerCase().trim();
                const cellNum = parseFloat(row[col]);

                // Si el filtro es una palabra color y la celda es numérica, aplicamos rango
                if (!isNaN(cellNum)) {
                    if (fVal === 'rojo') {
                        if (!(cellNum < 1)) return false;
                        continue;
                    } else if (fVal === 'amarillo') {
                        if (!(cellNum >= 1 && cellNum <= 30)) return false;
                        continue;
                    } else if (fVal === 'verde') {
                        if (!(cellNum > 30)) return false;
                        continue;
                    }
                }
                // Si no coincide con palabras clave o no es número, comportamiento default (texto)
            }

            const cellValue = (row[col] ?? '').toString().toLowerCase();
            if (!cellValue.includes(filterValue.toLowerCase())) {
                return false;
            }
        }

        if (globalFilter) {
            const g = globalFilter.toLowerCase();
            let matches = false;
            for (const col of headers) {
                const cellValue = (row[col] ?? '').toString().toLowerCase();
                if (cellValue.includes(g)) {
                    matches = true;
                    break;
                }
            }
            if (!matches) return false;
        }

        return true;
    });

    if (currentSort.column) {
        const col = currentSort.column;
        const dir = currentSort.direction === 'asc' ? 1 : -1;

        filtered.sort((a, b) => {
            let va = a[col];
            let vb = b[col];

            // Revisar si es una columna de fecha para usar parseo especial
            const dateColumns = ['reactivation', 'start', 'end', 'fecha_compra_', 'fecha_actual', 'fecha_hora_actual'];
            if (dateColumns.includes(col.toLowerCase())) {
                const ta = parseCustomDate(va);
                const tb = parseCustomDate(vb);
                return (ta - tb) * dir;
            }

            if (va == null && vb == null) return 0;
            if (va == null) return 1 * dir;
            if (vb == null) return -1 * dir;

            const nA = parseFloat(va);
            const nB = parseFloat(vb);
            const bothNumeric = !isNaN(nA) && !isNaN(nB);

            if (bothNumeric) {
                return (nA - nB) * dir;
            }

            return va.toString().localeCompare(vb.toString(), 'es', { numeric: true }) * dir;
        });
    }

    return filtered;
}

// --- Render tabla ---
function updateTable() {
    const tbody = dataTable.querySelector('tbody');
    tbody.innerHTML = '';

    const data = getFilteredAndSortedData();
    const totalFiltradas = data.length;
    const totalGlobal = allData.length;

    const totalPages = Math.max(1, Math.ceil(totalFiltradas / pageSize));
    if (currentPage > totalPages) currentPage = totalPages;

    const start = (currentPage - 1) * pageSize;
    const end = start + pageSize;
    const pageData = data.slice(start, end);

    // al cambiar contenido, limpiamos selección
    selectedCells.clear();
    clearVisualSelections();
    isSelecting = false;
    selectStart = null;

    // Calcular duplicados en las columnas "Password" y "Codigos" para los datos VISIBLES (filtrados)
    // OJO: ¿Duplicados en toda la tabla o solo en lo filtrado? 
    // Usualmente "identificar duplicados" se refiere al contexto actual.
    // Si queremos duplicados GLOBALES, deberíamos calcular sobre 'allData' o 'data'. 
    // Asumiremos sobre 'data' (lo filtrado y ordenado) para que el feedback sea sobre lo que el usuario trabaja.

    const countMapPassword = new Map();
    const countMapCodigos = new Map();

    // Contamos ocurrencias
    data.forEach(row => {
        const p = (row['Password'] || '').toString().trim();
        const c = (row['Codigos'] || '').toString().trim();
        if (p) countMapPassword.set(p, (countMapPassword.get(p) || 0) + 1);
        if (c) countMapCodigos.set(c, (countMapCodigos.get(c) || 0) + 1);
    });

    pageData.forEach((row, idxOnPage) => {
        const tr = document.createElement('tr');
        const filteredRowIndex = start + idxOnPage;
        tr.dataset.rowIndex = filteredRowIndex;

        headers.forEach((col, colIndex) => {
            const td = document.createElement('td');
            const value = row[col] ?? '';
            td.dataset.colIndex = colIndex;
            td.innerHTML = buildHighlightedHTML(value);

            const txtVal = (value || '').toString().trim();

            // Colorear Duplicados (Password / Codigos)
            if (col === 'Password' && txtVal && countMapPassword.get(txtVal) > 1) {
                td.classList.add('cell-duplicate-red');
            }
            if (col === 'Codigos' && txtVal && countMapCodigos.get(txtVal) > 1) {
                td.classList.add('cell-duplicate-red');
            }

            // Colorear "Pendientes_"
            if (col === 'Pendientes_') {
                const lower = txtVal.toLowerCase();
                if (lower === 'enviado') {
                    td.classList.add('status-enviado');
                } else if (lower.includes('processing')) {
                    td.classList.add('status-processing');
                }
            }

            // Colorear "Rest."
            if (col === 'Rest.') {
                const num = parseFloat(value);
                if (!isNaN(num)) {
                    if (num < 1) td.classList.add('cell-rest-red');
                    else if (num >= 1 && num <= 30) td.classList.add('cell-rest-yellow');
                    else if (num > 30) td.classList.add('cell-rest-green');
                }
            }

            // Colorear "Journal_"
            if (col === 'Journal_') {
                const lower = txtVal.toLowerCase();
                if (lower.includes('uptodate')) td.classList.add('cell-journal-uptodate');
                else if (lower.includes('bmj')) td.classList.add('cell-journal-bmj');
                else if (lower.includes('dynamed')) td.classList.add('cell-journal-dynamed');
                else if (lower.includes('clinicalkey')) td.classList.add('cell-journal-clinicalkey');
            }

            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    rowCountEl.textContent =
        `Filas en esta página: ${pageData.length} | Filtradas: ${totalFiltradas} | Total: ${totalGlobal}`;
    pageInfoEl.textContent = `Página ${currentPage} / ${totalPages}`;

    prevPageBtn.disabled = currentPage <= 1;
    nextPageBtn.disabled = currentPage >= totalPages;
}

// --- Scroll a columna por prefijo ---
function scrollToColumnByPrefix(prefix) {
    const wrapper = tableWrapper;
    if (!wrapper) return;

    const ths = dataTable.querySelectorAll('thead tr:first-child th');
    if (!ths.length) return;

    let targetTh = null;
    const p = prefix.toLowerCase();
    ths.forEach(th => {
        const txt = th.textContent.trim().toLowerCase();
        if (!targetTh && txt.startsWith(p)) {
            targetTh = th;
        }
    });

    if (!targetTh) return;
    wrapper.scrollLeft = targetTh.offsetLeft;
}

// --- Export ---
function exportVisibleToExcel() {
    const data = getFilteredAndSortedData();
    if (!data.length) {
        alert('No hay datos para exportar.');
        return;
    }

    const exportRows = data.map(row => {
        const obj = {};
        headers.forEach(h => { obj[h] = row[h] ?? ''; });
        return obj;
    });

    const ws = XLSX.utils.json_to_sheet(exportRows, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Datos');
    XLSX.writeFile(wb, 'tabla_combinada_filtrada.xlsx');
}

function exportVisibleToCSV() {
    const data = getFilteredAndSortedData();
    if (!data.length) {
        alert('No hay datos para exportar.');
        return;
    }

    const exportRows = data.map(row => {
        const obj = {};
        headers.forEach(h => { obj[h] = row[h] ?? ''; });
        return obj;
    });

    const ws = XLSX.utils.json_to_sheet(exportRows, { header: headers });
    const csv = XLSX.utils.sheet_to_csv(ws);

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'tabla_combinada_filtrada.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// --- Integración con Electron (Escritorio) ---
if (window.electronAPI) {
    console.log("Modo Electron detectado");

    // Crear botón "Cargar desde OneDrive"
    const btnParam = document.createElement('button');
    btnParam.textContent = "🔄 Cargar desde OneDrive";
    btnParam.style.backgroundColor = "#2ecc71"; // Verde distinto
    btnParam.style.marginLeft = "10px";

    btnParam.onclick = async () => {
        statusEl.textContent = "Cargando archivos desde OneDrive (Local)...";
        try {
            const result = await window.electronAPI.getOneDriveFiles();
            if (!result.success) {
                alert("Error cargando archivos: " + result.error);
                statusEl.textContent = "Error al cargar archivos via Electron.";
                return;
            }

            if (result.count === 0) {
                alert(`No se encontraron archivos Excel en: ${result.path}`);
                statusEl.textContent = "Carpeta vacía o sin Excel.";
                return;
            }

            console.log(`Recibidos ${result.count} archivos desde ${result.path}`);

            // Adaptar los datos de Electron al formato que espera processFiles
            // processFiles espera objetos con .name y .arrayBuffer()
            const pseudoFiles = result.files.map(f => ({
                name: f.name,
                arrayBuffer: async () => f.data
            }));

            // Actualizar referencia global (aunque no son File reales, funcionan igual para lectura)
            currentFiles = pseudoFiles;
            await processFiles(pseudoFiles);

            statusEl.textContent = `¡Carga automática completada! (${result.count} archivos)`;

        } catch (err) {
            console.error(err);
            statusEl.textContent = "Error inesperado en modo escritorio.";
        }
    };

    // Insertar el botón en la UI
    const uploadRight = document.querySelector('.upload-right');
    if (uploadRight) {
        // Insertar al principio de los botones
        uploadRight.insertBefore(btnParam, uploadRight.firstChild);
    }
}
