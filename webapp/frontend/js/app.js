const API_BASE = '';

document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initHealthCheck();
    initMoverArchivo();
    initExportarCorreos();
    initCorreoMasivo();
    initDividirPDF();
    initCrearCarpetas();
});

function initTabs() {
    const tabs = document.querySelectorAll('#mainTabs .nav-link');
    const panes = document.querySelectorAll('.tab-pane');

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');

            const target = tab.dataset.tab;
            panes.forEach(pane => {
                pane.classList.add('d-none');
                pane.classList.remove('active');
            });

            const targetPane = document.getElementById(`tab-${target}`);
            targetPane.classList.remove('d-none');
            targetPane.classList.add('active');
        });
    });
}

async function initHealthCheck() {
    const badge = document.getElementById('healthBadge');
    try {
        const res = await fetch(`${API_BASE}/api/health`);
        const data = await res.json();

        if (data.outlook && data.word) {
            badge.textContent = 'Outlook y Word OK';
            badge.className = 'badge bg-success';
        } else if (data.outlook || data.word) {
            const missing = [];
            if (!data.outlook) missing.push('Outlook');
            if (!data.word) missing.push('Word');
            badge.textContent = `Falta: ${missing.join(', ')}`;
            badge.className = 'badge bg-warning';
        } else {
            badge.textContent = 'Outlook y Word no disponibles';
            badge.className = 'badge bg-danger';
        }
    } catch (e) {
        badge.textContent = 'Servidor no disponible';
        badge.className = 'badge bg-danger';
    }
}

function showProgress(id) {
    document.getElementById(id).classList.remove('d-none');
}

function hideProgress(id) {
    document.getElementById(id).classList.add('d-none');
}

function renderSummary(containerId, data) {
    const container = document.getElementById(containerId);
    container.innerHTML = `
        <div class="summary-badges">
            <span class="badge bg-success">${data.success_count} exitosos</span>
            <span class="badge bg-danger">${data.error_count} errores</span>
            <span class="badge bg-secondary">${data.total_rows} total</span>
        </div>
        ${data.download_url ? `<a href="${API_BASE}${data.download_url}" class="btn btn-outline-primary btn-sm download-btn">Descargar Excel con resultados</a>` : ''}
    `;
}

function renderTable(containerId, columns, rows) {
    const container = document.getElementById(containerId);

    let html = '<div class="table-container"><table class="table table-sm results-table"><thead><tr>';
    columns.forEach(col => {
        html += `<th>${col.label}</th>`;
    });
    html += '</tr></thead><tbody>';

    rows.forEach(row => {
        const rowClass = row.success ? 'row-success' : 'row-error';
        html += `<tr class="${rowClass}">`;
        columns.forEach(col => {
            let val = row[col.key];
            if (col.key === 'success') {
                val = val ? 'OK' : 'ERROR';
            }
            if (val === null || val === undefined) val = '';
            html += `<td>${val}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table></div>';
    container.innerHTML = html;
}

function initMoverArchivo() {
    const fileInput = document.getElementById('mover-file');
    const btn = document.getElementById('mover-btn');

    fileInput.addEventListener('change', () => {
        btn.disabled = !fileInput.files.length;
    });

    btn.addEventListener('click', async () => {
        if (!fileInput.files.length) return;

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);

        btn.disabled = true;
        showProgress('mover-progress');
        document.getElementById('mover-summary').innerHTML = '';
        document.getElementById('mover-results').innerHTML = '';

        try {
            const res = await fetch(`${API_BASE}/api/mover-archivo`, { method: 'POST', body: formData });
            const data = await res.json();

            renderSummary('mover-summary', data);
            renderTable('mover-results', [
                { key: 'row', label: 'Fila' },
                { key: 'source', label: 'Origen' },
                { key: 'destination', label: 'Destino' },
                { key: 'success', label: 'Estado' },
                { key: 'error', label: 'Error' }
            ], data.rows);
        } catch (e) {
            document.getElementById('mover-summary').innerHTML = `<div class="alert alert-danger">Error: ${e.message}</div>`;
        } finally {
            hideProgress('mover-progress');
            btn.disabled = false;
        }
    });
}

function initExportarCorreos() {
    const fileInput = document.getElementById('exportar-file');
    const folderInput = document.getElementById('exportar-folder');
    const btn = document.getElementById('exportar-btn');

    function checkReady() {
        btn.disabled = !(fileInput.files.length && folderInput.value.trim());
    }

    fileInput.addEventListener('change', checkReady);
    folderInput.addEventListener('input', checkReady);

    btn.addEventListener('click', async () => {
        if (!fileInput.files.length || !folderInput.value.trim()) return;

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('output_folder', folderInput.value.trim());

        btn.disabled = true;
        showProgress('exportar-progress');
        document.getElementById('exportar-summary').innerHTML = '';
        document.getElementById('exportar-results').innerHTML = '';

        try {
            const res = await fetch(`${API_BASE}/api/exportar-correos`, { method: 'POST', body: formData });
            const data = await res.json();

            if (data.error) {
                document.getElementById('exportar-summary').innerHTML = `<div class="alert alert-danger">${data.error}</div>`;
            } else {
                renderSummary('exportar-summary', data);
                renderTable('exportar-results', [
                    { key: 'row', label: 'Fila' },
                    { key: 'destinatario', label: 'Destinatario' },
                    { key: 'asunto', label: 'Asunto' },
                    { key: 'pdf_path', label: 'PDF' },
                    { key: 'success', label: 'Estado' },
                    { key: 'error', label: 'Error' }
                ], data.rows);
            }
        } catch (e) {
            document.getElementById('exportar-summary').innerHTML = `<div class="alert alert-danger">Error: ${e.message}</div>`;
        } finally {
            hideProgress('exportar-progress');
            btn.disabled = false;
        }
    });
}

function initCorreoMasivo() {
    const fileInput = document.getElementById('correo-file');
    const folderInput = document.getElementById('correo-folder');
    const btn = document.getElementById('correo-btn');

    function checkReady() {
        btn.disabled = !(fileInput.files.length && folderInput.value.trim());
    }

    fileInput.addEventListener('change', checkReady);
    folderInput.addEventListener('input', checkReady);

    btn.addEventListener('click', async () => {
        if (!fileInput.files.length || !folderInput.value.trim()) return;

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('pdf_folder', folderInput.value.trim());
        formData.append('columna_nombre', document.getElementById('correo-col-nombre').value);
        formData.append('columna_correo', document.getElementById('correo-col-correo').value);

        btn.disabled = true;
        showProgress('correo-progress');
        document.getElementById('correo-summary').innerHTML = '';
        document.getElementById('correo-results').innerHTML = '';

        try {
            const res = await fetch(`${API_BASE}/api/correo-masivo`, { method: 'POST', body: formData });
            const data = await res.json();

            renderSummary('correo-summary', data);
            renderTable('correo-results', [
                { key: 'row', label: 'Fila' },
                { key: 'correo', label: 'Correo' },
                { key: 'archivo', label: 'Archivo' },
                { key: 'success', label: 'Estado' },
                { key: 'error', label: 'Error' }
            ], data.rows);
        } catch (e) {
            document.getElementById('correo-summary').innerHTML = `<div class="alert alert-danger">Error: ${e.message}</div>`;
        } finally {
            hideProgress('correo-progress');
            btn.disabled = false;
        }
    });
}

function initDividirPDF() {
    const fileInput = document.getElementById('dividir-file');
    const wordInput = document.getElementById('dividir-word');
    const btn = document.getElementById('dividir-btn');

    function checkReady() {
        btn.disabled = !(fileInput.files.length && wordInput.files.length);
    }

    fileInput.addEventListener('change', checkReady);
    wordInput.addEventListener('change', checkReady);

    btn.addEventListener('click', async () => {
        if (!fileInput.files.length || !wordInput.files.length) return;

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('word_document', wordInput.files[0]);

        btn.disabled = true;
        showProgress('dividir-progress');
        document.getElementById('dividir-summary').innerHTML = '';
        document.getElementById('dividir-results').innerHTML = '';

        try {
            const res = await fetch(`${API_BASE}/api/dividir-resoluciones`, { method: 'POST', body: formData });
            const data = await res.json();

            renderSummary('dividir-summary', data);
            renderTable('dividir-results', [
                { key: 'chunk', label: '#' },
                { key: 'pages', label: 'Páginas' },
                { key: 'filename', label: 'Nombre PDF' },
                { key: 'success', label: 'Estado' },
                { key: 'error', label: 'Error' }
            ], data.rows);
        } catch (e) {
            document.getElementById('dividir-summary').innerHTML = `<div class="alert alert-danger">Error: ${e.message}</div>`;
        } finally {
            hideProgress('dividir-progress');
            btn.disabled = false;
        }
    });
}

function initCrearCarpetas() {
    const fileInput = document.getElementById('carpetas-file');
    const pathInput = document.getElementById('carpetas-path');
    const btn = document.getElementById('carpetas-btn');

    function checkReady() {
        btn.disabled = !(fileInput.files.length && pathInput.value.trim());
    }

    fileInput.addEventListener('change', checkReady);
    pathInput.addEventListener('input', checkReady);

    btn.addEventListener('click', async () => {
        if (!fileInput.files.length || !pathInput.value.trim()) return;

        const formData = new FormData();
        formData.append('file', fileInput.files[0]);
        formData.append('base_path', pathInput.value.trim());
        formData.append('columna', document.getElementById('carpetas-col').value);

        btn.disabled = true;
        showProgress('carpetas-progress');
        document.getElementById('carpetas-summary').innerHTML = '';
        document.getElementById('carpetas-results').innerHTML = '';

        try {
            const res = await fetch(`${API_BASE}/api/crear-carpetas`, { method: 'POST', body: formData });
            const data = await res.json();

            renderSummary('carpetas-summary', data);
            renderTable('carpetas-results', [
                { key: 'row', label: 'Fila' },
                { key: 'folder_name', label: 'Carpeta' },
                { key: 'full_path', label: 'Ruta completa' },
                { key: 'success', label: 'Estado' },
                { key: 'error', label: 'Error' }
            ], data.rows);
        } catch (e) {
            document.getElementById('carpetas-summary').innerHTML = `<div class="alert alert-danger">Error: ${e.message}</div>`;
        } finally {
            hideProgress('carpetas-progress');
            btn.disabled = false;
        }
    });
}
