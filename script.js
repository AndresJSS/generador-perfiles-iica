let perfilesData = [];
let perfilesFiltrados = [];

const officeSelect = document.getElementById('officeSelect');
const btnDownload = document.getElementById('btnDownload');
const pdfContainer = document.getElementById('pdf-container');
const emptyState = document.getElementById('empty-state');

officeSelect.addEventListener('change', () => {
    renderPages();
    window.scrollTo({ top: 0, behavior: 'smooth' });
});

btnDownload.addEventListener('click', exportarPDF);

window.addEventListener('DOMContentLoaded', async () => {
    try {
        const response = await fetch('./funcionarios.xlsx');
        if (!response.ok) throw new Error("Archivo no encontrado");
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const sheetName = workbook.SheetNames.includes("perfil_profesional") ? "perfil_profesional" : workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        perfilesData = XLSX.utils.sheet_to_json(sheet);
        
        if (perfilesData.length > 0) {
            const sedesSet = new Set();
            perfilesData.forEach(person => {
                const sede = obtenerValor(person, ['representacion_sede', 'representacion', 'sede']);
                if (sede) sedesSet.add(sede.trim());
            });

            const sedesArray = Array.from(sedesSet).sort();
            officeSelect.innerHTML = ''; 
            sedesArray.forEach(sede => {
                const option = document.createElement('option');
                option.value = sede;
                option.textContent = sede;
                officeSelect.appendChild(option);
            });

            btnDownload.disabled = false;
            btnDownload.innerHTML = 'Exportar a PDF';
            emptyState.style.display = 'none';
            
            renderPages(); 
        }
    } catch (error) {
        console.error("Error cargando el Excel automático:", error);
        emptyState.innerHTML = `<div class="text-red-500 font-semibold">❌ Error: No se pudo cargar el archivo 'funcionarios.xlsx'.<br>Asegúrate de que está subido en la misma carpeta.</div>`;
        btnDownload.innerHTML = 'Error de conexión';
    }
});

function obtenerValor(row, posiblesNombres) {
    for (let key of Object.keys(row)) {
        const cleanKey = key.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ /g, "_");
        if (posiblesNombres.includes(cleanKey)) {
            return row[key] || "";
        }
    }
    return "";
}

function esRepresentantePrincipal(cargo) {
    if (!cargo) return false;
    const c = cargo.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    if (c.includes("asistente") || c.includes("assistant") || c.includes("adjunto") || c.includes("deputy") || c.includes("apoyo")) {
        return false;
    }
    
    return c.includes("representante") || c.includes("representative");
}

function renderPages() {
    if (perfilesData.length === 0) return;
    
    pdfContainer.innerHTML = "";
    
    const representacionSeleccionada = officeSelect.value;
    
    const hoy = new Date();
    const fecha = `${hoy.getDate().toString().padStart(2, '0')}/${(hoy.getMonth() + 1).toString().padStart(2, '0')}/${hoy.getFullYear()}`;
    
    const tituloDocumento = "Dashboard de matriz de capacidades técnicas";
    const subtituloDocumento = `Representación del IICA ${representacionSeleccionada}`;

    perfilesFiltrados = perfilesData.filter(person => {
        const sede = obtenerValor(person, ['representacion_sede', 'representacion', 'sede']);
        return sede.trim().toLowerCase() === representacionSeleccionada.trim().toLowerCase();
    });

    perfilesFiltrados.sort((a, b) => {
        const cargoA = obtenerValor(a, ['tipo_de_cargo', 'cargo', 'puesto']);
        const cargoB = obtenerValor(b, ['tipo_de_cargo', 'cargo', 'puesto']);
        
        const aEsJefe = esRepresentantePrincipal(cargoA);
        const bEsJefe = esRepresentantePrincipal(cargoB);
        
        if (aEsJefe && !bEsJefe) return -1; 
        if (!aEsJefe && bEsJefe) return 1;  
        return 0; 
    });

    if (perfilesFiltrados.length === 0) {
        btnDownload.disabled = true;
        pdfContainer.innerHTML = `
            <div class="bg-white p-10 rounded-xl border border-dashed border-slate-300 text-center text-slate-500 w-full max-w-[210mm]">
                No se encontraron funcionarios asignados a <b>${representacionSeleccionada}</b>.
            </div>`;
        return;
    }

    btnDownload.disabled = false;
    const totalPaginas = Math.ceil(perfilesFiltrados.length / 2);

    for (let i = 0; i < perfilesFiltrados.length; i += 2) {
        const chunk = perfilesFiltrados.slice(i, i + 2);
        const numPagina = Math.floor(i / 2) + 1;
        
        const page = document.createElement('div');
        page.className = "pdf-page";

        // Se eliminó la línea azul claro (el acento absoluto)
        // Se ajustaron los márgenes (margin-bottom a 6mm) para que las tarjetas no choquen con el pie de página
        let htmlContent = `
            <div style="position: absolute; top: 0; left: 0; right: 0; height: 4mm; background-color: #003366;"></div>
            
            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 6mm; border-bottom: 1.2mm solid #003366; padding-bottom: 4mm; position: relative;">
                
                <div style="display: flex; align-items: center; gap: 4mm;">
                    <img src="logo.png" alt="Logo" style="max-height: 16mm; max-width: 36mm; object-fit: contain;" 
                         onerror="this.style.display='none'">
                    
                    <div style="border-left: 1px solid #CBD5E1; padding-left: 4mm; white-space: nowrap;">
                        <h1 class="text-primary" style="font-size: 14.5pt; font-weight: bold; margin: 0 0 1mm 0;">${tituloDocumento}</h1>
                        <h2 class="text-muted" style="font-size: 10.5pt; font-weight: normal; margin: 0;">${subtituloDocumento}</h2>
                    </div>
                </div>

                <div class="text-primary" style="font-size: 10.5pt; font-weight: bold; margin-bottom: 1.5mm; white-space: nowrap;">
                    ${fecha}
                </div>
            </div>
            
            <!-- Se redujo el gap (espacio entre tarjetas) a 6mm para evitar cortes -->
            <div style="flex: 1; display: flex; flex-direction: column; gap: 6mm;">
        `;

        chunk.forEach(person => {
            const nombre = obtenerValor(person, ['nombre_completo', 'nombre', 'funcionario']);
            const correo = obtenerValor(person, ['correo_electronico', 'correo', 'email']);
            const cargo = obtenerValor(person, ['tipo_de_cargo', 'cargo', 'puesto']);
            const nacionalidad = obtenerValor(person, ['nacionalidad', 'pais']);
            const resumen = obtenerValor(person, ['resumen_ejecutivo', 'resumen', 'biografia']);
            const foto = obtenerValor(person, ['foto', 'imagen']);

            let metaTexto = [];
            if (nacionalidad) metaTexto.push(`<div class="meta-line"><b>Nacionalidad:</b> ${nacionalidad}</div>`);
            if (correo) metaTexto.push(`<div class="meta-line"><b>Correo:</b> ${correo}</div>`);
            const metaHtml = metaTexto.join('');

            const photoHtml = foto 
                ? `<img src="${foto}" class="photo-img" onerror="this.outerHTML='<div class=\\'photo-placeholder\\'><div class=\\'silhouette\\'></div><div class=\\'silhouette-body\\'></div></div>'">`
                : `<div class="photo-placeholder"><div class="silhouette"></div><div class="silhouette-body"></div></div>`;

            htmlContent += `
                <div class="profile-card">
                    <div class="photo-container">${photoHtml}</div>
                    <div class="text-content">
                        <h3 class="text-primary text-name">${nombre || 'Sin nombre'}</h3>
                        <p class="text-secondary text-role">${cargo || 'Cargo no especificado'}</p>
                        <p class="text-muted text-meta">${metaHtml}</p>
                        <div class="divider"></div>
                        <p class="text-body text-bio">${resumen ? resumen.replace(/\n/g, '<br>') : ''}</p>
                    </div>
                </div>
            `;
        });

        htmlContent += `</div>`;

        htmlContent += `
            <div style="border-top: 0.6mm solid #E2E8F0; padding-top: 4mm; display: flex; justify-content: space-between; font-size: 8pt;" class="text-muted">
                <span>${subtituloDocumento}</span>
                <span>Página ${numPagina} de ${totalPaginas}</span>
            </div>
        `;

        page.innerHTML = htmlContent;
        pdfContainer.appendChild(page);
    }
}

function exportarPDF() {
    if (perfilesFiltrados.length === 0) return;
    
    const representacion = officeSelect.value.replace(/\s+/g, '_');
    const container = document.getElementById('pdf-container');
    const pages = document.querySelectorAll('.pdf-page');
    
    const oldContainerClass = container.className;
    container.className = "block"; 
    
    pages.forEach(p => { 
        p.style.boxShadow = 'none'; 
        p.style.margin = '0'; 
        p.style.pageBreakAfter = 'avoid'; 
        p.style.height = '296mm'; 
        p.style.overflow = 'hidden';
    });

    const opt = {
        margin:       0,
        filename:     `Perfiles_${representacion}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, scrollY: 0, scrollX: 0 },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    document.getElementById('control-panel').style.opacity = '0.5';
    const textoOriginal = btnDownload.innerHTML;
    btnDownload.innerHTML = 'Generando PDF (Espere)...';

    html2pdf().set(opt).from(container).save().then(() => {
        container.className = oldContainerClass;
        pages.forEach(p => { 
            p.style.boxShadow = '0 10px 25px rgba(0,0,0,0.1)'; 
            p.style.margin = '0 auto 24px auto'; 
            p.style.pageBreakAfter = 'always';
            p.style.height = '297mm';
            p.style.overflow = 'visible';
        });
        
        document.getElementById('control-panel').style.opacity = '1';
        btnDownload.innerHTML = textoOriginal;
    }).catch(err => {
        console.error("Error al exportar:", err);
        btnDownload.innerHTML = 'Error al exportar';
    });
}