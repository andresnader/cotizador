// --- 1. CONFIGURACIÓN DE GOOGLE API ---

// ID de Cliente (El tuyo)
const CLIENT_ID = "1005993441268-4sle7juetquq55efkqehm252ebtml4tv.apps.googleusercontent.com";

// Define los "scopes" o permisos que tu app necesita.
const SCOPES = [
    'https://www.googleapis.com/auth/userinfo.email', // Para ver el email del usuario
    'https://www.googleapis.com/auth/userinfo.profile', // Para ver el nombre del usuario
    'https://www.googleapis.com/auth/spreadsheets', // Para leer/escribir en Google Sheets
    'https://www.googleapis.com/auth/drive' // Para crear/modificar archivos en Google Drive (Docs)
].join(' ');

// *** IDs DE TUS ARCHIVOS (ACTUALIZADOS) ***

// ID de tu Hoja de Clientes
const CLIENTS_SHEET_ID = '1JnowPVvio2tSjNSmQH6z8drqejmiwwVthmkFj2eGf2g';
// Rango de Clientes: Columnas A (RUC) a F (Email)
const CLIENTS_SHEET_RANGE = 'Hoja 1!A2:F'; 

// ID de tu Hoja de Servicios/Productos
const SERVICES_SHEET_ID = '19FdGfh7wtznlOHsnp5Ntd9ZF2fSZLd11yUD40VP0VNI';
// Rango de Servicios: Columnas A (id) a E (updated_at)
const SERVICES_SHEET_RANGE = 'Hoja 1!A2:E'; 

// IDs INSERTADOS DESDE LAS IMÁGENES
// ID de tu Hoja de Historial de Cotizaciones
const QUOTES_SHEET_ID = '153s8Lhum68zConWSbbFQtbR3dYzkafHy-4YQJ1HpvLk';
// Asegúrate que esta hoja tenga una pestaña llamada 'Hoja 1' y las columnas correctas
const QUOTES_SHEET_RANGE = 'Hoja 1!A2:K'; // Asume A=ID, B=Numero, C=Fecha Emisión, D=Fecha Validez, E=Cliente JSON, F=Items JSON, G=Total, H=Estado, I=Notas, J=Google Doc ID, K=Google PDF ID

// ID de tu Plantilla de Contrato en Google Docs
const CONTRACT_TEMPLATE_ID = '1MNE19Ymp40dTZ4Ulv31An3EWVRc8SW-ktD2H-dOFSa4';

// Carpeta donde se almacenarán los contratos generados
const CONTRACTS_DRIVE_FOLDER_ID = '1doqaU7jvk5dGH1TTAJfniGU9boZn5JrB';
// Carpeta donde se almacenarán las cotizaciones (Docs y PDFs)
const QUOTES_DRIVE_FOLDER_ID = '1mTsrPMrcLzSDapMs4m3fG_d5Q8EguEoX';

// Hoja donde se almacenará la configuración personalizada por usuario
const COMPANY_CONFIG_SHEET_ID = 'ID_DE_TU_HOJA_DE_CONFIGURACION';
const COMPANY_CONFIG_SHEET_TAB = 'Configuracion';
const COMPANY_CONFIG_SHEET_RANGE = `${COMPANY_CONFIG_SHEET_TAB}!A2:C`;

let tokenClient;
// let gapiInited = false; // Ya no se usan como flags globales simples
// let gisInited = false;

// --- DECLARACIÓN DE VARIABLES GLOBALES DE ELEMENTOS DEL DOM ---
let loginOverlay, appContainer, authStatus, loginButton, logoutButton, loadingSpinner;
let clientForm, clientList, clientRowIdInput, clientNameInput, clientRUCInput, clientContactInput;
let serviceForm, serviceList, serviceRowIdInput, serviceNameInput, serviceDescriptionInput, servicePriceInput, serviceIdInput;
let quotesHistoryList;
let quoteClientSelect, quoteRUCInput, serviceSelect, servicePriceOverrideInput, quoteItemsBody, quoteNumberInput, quoteNotesInput, quoteDate;
let generateQuotePdfBtn, generateContractBtn;
let configForm, companyNameInput, companyAddressInput, companyContactInput, companyWebsiteInput, companyWhatsappInput, logoUploadInput, logoPreview, primaryColorInput, accentColorInput;
let contractClientSelect, contractQuoteSelect, contractText, saveContractBtn;
let portalClientSelect, clientPortalContent;
let remoteConfigStatus, remoteConfigMeta, loadCloudConfigBtn, saveCloudConfigBtn, autoSyncCloudToggle;

let signedInUser = null;

// --- INICIO: LÓGICA DE MODAL PERSONALIZADO ---
let modalOverlay, modalTitle, modalBody, modalBtnConfirm, modalBtnCancel;
let confirmCallback = null;

function initModal() {
    modalOverlay = document.getElementById('custom-modal-overlay');
    modalTitle = document.getElementById('custom-modal-title');
    modalBody = document.getElementById('custom-modal-body');
    modalBtnConfirm = document.getElementById('custom-modal-btn-confirm');
    modalBtnCancel = document.getElementById('custom-modal-btn-cancel');

    if (modalOverlay) {
        modalOverlay.addEventListener('click', (e) => {
            if (e.target === modalOverlay) {
                hideModal();
            }
        });
    }
    if (modalBtnCancel) {
        modalBtnCancel.addEventListener('click', hideModal);
    }
    if (modalBtnConfirm) {
        modalBtnConfirm.addEventListener('click', () => {
            if (typeof confirmCallback === 'function') {
                confirmCallback();
            }
            hideModal();
        });
    }
}

function hideModal() {
    if (modalOverlay) modalOverlay.classList.remove('visible');
    confirmCallback = null; // Limpiar callback
}

/**
 * Muestra una alerta no bloqueante.
 * @param {string} message El mensaje a mostrar.
 * @param {string} [title='Alerta'] El título opcional del modal.
 */
function showAlert(message, title = 'Alerta') {
    if (!modalOverlay || !modalTitle || !modalBody || !modalBtnConfirm || !modalBtnCancel) {
        console.error("Modal no inicializado, usando alert() como fallback.");
        alert(message); // Fallback si el modal no existe
        return;
    }
    modalTitle.innerText = title;
    modalBody.innerText = message;
    modalBtnConfirm.innerText = 'Aceptar';
    modalBtnCancel.style.display = 'none'; // Ocultar botón de cancelar
    confirmCallback = null; // Es una alerta, no hay callback de confirmación
    modalOverlay.classList.add('visible');
}

/**
 * Muestra una confirmación no bloqueante.
 * @param {string} message El mensaje de confirmación.
 * @param {function} onConfirm El callback que se ejecutará si el usuario confirma.
 * @param {string} [title='Confirmación'] El título opcional del modal.
 */
function showConfirm(message, onConfirm, title = 'Confirmación') {
    if (!modalOverlay || !modalTitle || !modalBody || !modalBtnConfirm || !modalBtnCancel) {
        console.error("Modal no inicializado, cancelando acción.");
        // No podemos usar confirm() aquí porque esperamos un callback.
        // Lo más seguro es no ejecutar la acción.
        return;
    }
    modalTitle.innerText = title;
    modalBody.innerText = message;
    modalBtnConfirm.innerText = 'Confirmar';
    modalBtnCancel.style.display = 'inline-block'; // Mostrar botón de cancelar
    confirmCallback = onConfirm; // Asignar el callback
    modalOverlay.classList.add('visible');
}
// --- FIN: LÓGICA DE MODAL PERSONALIZADO ---

function arrayBufferToBase64(buffer) {
    if (!buffer) return '';
    if (typeof buffer === 'string') {
        return btoa(buffer);
    }
    let binary = '';
    const bytes = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : new Uint8Array(buffer.buffer || buffer);
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
        const chunk = bytes.subarray(i, i + chunkSize);
        binary += String.fromCharCode.apply(null, chunk);
    }
    return btoa(binary);
}

function escapeHtml(value) {
    if (value === null || value === undefined) return '';
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function formatMultilineText(value) {
    if (!value) return '';
    return escapeHtml(value).replace(/\r?\n/g, '<br>');
}

let cachedPdfLibraries = null;

function resolveHtml2canvasFunction() {
    if (typeof window.html2canvas === 'function') {
        return window.html2canvas;
    }
    if (window.html2canvas && typeof window.html2canvas.default === 'function') {
        window.html2canvas = window.html2canvas.default;
        return window.html2canvas;
    }
    return null;
}

function resolveJsPdfConstructor() {
    if (window.jspdf && typeof window.jspdf.jsPDF === 'function') {
        return window.jspdf.jsPDF;
    }
    return null;
}

async function ensurePdfLibrariesAvailable() {
    if (cachedPdfLibraries) {
        return cachedPdfLibraries;
    }

    const html2canvasFn = resolveHtml2canvasFunction();
    if (typeof html2canvasFn !== 'function') {
        throw new Error('La librería html2canvas no está disponible. Verifica la conexión a internet o el script CDN.');
    }

    const jsPDFConstructor = resolveJsPdfConstructor();
    if (typeof jsPDFConstructor !== 'function') {
        throw new Error('La librería jsPDF no está disponible. Verifica la conexión a internet o el script CDN.');
    }

    cachedPdfLibraries = { html2canvasFn, jsPDFConstructor };
    return cachedPdfLibraries;
}

async function waitForImagesToLoad(container) {
    const images = Array.from(container.querySelectorAll('img'));
    if (images.length === 0) return;
    await Promise.all(images.map(img => {
        if (img.complete && img.naturalWidth !== 0) return Promise.resolve();
        return new Promise(resolve => {
            img.onload = resolve;
            img.onerror = resolve; // Resolver incluso si falla para evitar bloqueo
        });
    }));
}

async function generateQuotePdfBlob(quoteRecord, pdfLibraries) {
    const { html2canvasFn, jsPDFConstructor } = pdfLibraries || await ensurePdfLibrariesAvailable();
    const printArea = document.getElementById('print-area');
    if (!printArea) {
        throw new Error('No se encontró el contenedor de impresión para generar el PDF.');
    }

    const previousHtml = printArea.innerHTML;
    const previousWidth = printArea.style.width;
    try {
        printArea.style.width = '794px'; // Aproximadamente ancho carta/A4
        const settingsForPdf = quoteRecord.companySettingsForPrint || companySettings;
        printArea.innerHTML = buildPrintableHtml({ ...quoteRecord, companySettings: settingsForPdf });
        await waitForImagesToLoad(printArea);

        const canvas = await html2canvasFn(printArea, {
            scale: 2,
            useCORS: true,
            backgroundColor: '#ffffff'
        });

        const imageData = canvas.toDataURL('image/png');
        const pdf = new jsPDFConstructor({ orientation: 'portrait', unit: 'pt', format: 'a4' });
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();

        const ratio = Math.min(pageWidth / canvas.width, pageHeight / canvas.height);
        const imgWidth = canvas.width * ratio;
        const imgHeight = canvas.height * ratio;
        const marginX = (pageWidth - imgWidth) / 2;
        const marginY = (pageHeight - imgHeight) / 2;

        pdf.addImage(imageData, 'PNG', marginX, marginY, imgWidth, imgHeight, undefined, 'FAST');
        return pdf.output('blob');
    } finally {
        printArea.innerHTML = previousHtml;
        printArea.style.width = previousWidth;
    }
}

async function uploadPdfBlobToDrive(pdfBlob, pdfName, targetFolderId = QUOTES_DRIVE_FOLDER_ID) {
    if (!pdfBlob) {
        throw new Error('No se generó ningún archivo PDF para subir.');
    }

    const token = (typeof gapi?.client?.getToken === 'function') ? gapi.client.getToken() : gapi?.auth?.getToken?.();
    const accessToken = token?.access_token;
    if (!accessToken) {
        throw new Error('No se encontró un token de acceso válido para subir el PDF a Drive.');
    }

    const metadata = {
        name: pdfName,
        mimeType: 'application/pdf'
    };
    if (targetFolderId) {
        metadata.parents = [targetFolderId];
    }

    const boundary = '-------314159265358979323846';
    const delimiter = `\r\n--${boundary}\r\n`;
    const closeDelimiter = `\r\n--${boundary}--`;
    const pdfBuffer = await pdfBlob.arrayBuffer();
    const base64Data = arrayBufferToBase64(pdfBuffer);

    const multipartRequestBody =
        delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
        JSON.stringify(metadata) +
        delimiter + 'Content-Type: application/pdf\r\nContent-Transfer-Encoding: base64\r\n\r\n' +
        base64Data +
        closeDelimiter;

    const uploadResponse = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true', {
        method: 'POST',
        headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': `multipart/related; boundary=${boundary}`
        },
        body: multipartRequestBody
    });

    if (!uploadResponse.ok) {
        const errorText = await uploadResponse.text();
        throw new Error(`No se pudo guardar el PDF en Drive: ${uploadResponse.status} ${errorText}`);
    }

    const uploadResult = await uploadResponse.json();
    console.log('PDF guardado en Drive con ID:', uploadResult.id, 'padres:', uploadResult.parents);
    return uploadResult?.id || null;
}

async function trashDriveFile(fileId) {
    if (!fileId || !gapi?.client?.drive) return;
    try {
        await gapi.client.drive.files.update({
            fileId,
            resource: { trashed: true },
            supportsAllDrives: true
        });
    } catch (error) {
        console.error('Error enviando archivo a la papelera en Drive:', error);
    }
}


// --- FUNCIÓN PARA ASIGNAR ELEMENTOS DEL DOM ---
function showLoginOverlay() {
    if (!loginOverlay) return;
    loginOverlay.classList.remove('d-none');
    loginOverlay.classList.add('d-flex');
    loginOverlay.style.display = 'flex';
}

function hideLoginOverlay() {
    if (!loginOverlay) return;
    loginOverlay.classList.remove('d-flex');
    loginOverlay.classList.add('d-none');
    loginOverlay.style.display = 'none';
}

function assignDOMElements() {
    // ... (asignaciones iguales que antes)...
    loginOverlay = document.getElementById('login-overlay');
    appContainer = document.getElementById('app-container');
    authStatus = document.getElementById('auth-status');
    loginButton = document.getElementById('google-login-btn');
    logoutButton = document.getElementById('google-logout-btn');
    loadingSpinner = document.getElementById('loading-spinner');

    clientForm = document.getElementById('clientForm');
    clientList = document.getElementById('clientList');
    clientRowIdInput = document.getElementById('clientRowId'); 
    clientNameInput = document.getElementById('clientName');
    clientRUCInput = document.getElementById('clientRUC');
    clientContactInput = document.getElementById('clientContact');

    serviceForm = document.getElementById('serviceForm');
    serviceList = document.getElementById('serviceList');
    serviceRowIdInput = document.getElementById('serviceRowId');
    serviceIdInput = document.getElementById('serviceId'); // Asignar serviceIdInput
    serviceNameInput = document.getElementById('serviceName');
    serviceDescriptionInput = document.getElementById('serviceDescription');
    servicePriceInput = document.getElementById('servicePrice');

    quotesHistoryList = document.getElementById('quotesHistoryList');

    quoteClientSelect = document.getElementById('quoteClient');
    quoteRUCInput = document.getElementById('quoteRUC');
    serviceSelect = document.getElementById('serviceSelect');
    servicePriceOverrideInput = document.getElementById('servicePriceOverride');
    quoteItemsBody = document.getElementById('quoteItems');
    quoteNumberInput = document.getElementById('quoteNumber');
    quoteNotesInput = document.getElementById('quoteNotes');
    quoteDate = document.getElementById('quoteDate'); // Asignar quoteDate
    generateQuotePdfBtn = document.getElementById('generate-quote-pdf-btn');
    generateContractBtn = document.getElementById('generate-contract-btn');

    configForm = document.getElementById('configForm');
    companyNameInput = document.getElementById('companyName');
    companyAddressInput = document.getElementById('companyAddress');
    companyContactInput = document.getElementById('companyContact');
    companyWebsiteInput = document.getElementById('companyWebsite');
    companyWhatsappInput = document.getElementById('companyWhatsapp');
    logoUploadInput = document.getElementById('logoUpload');
    logoPreview = document.getElementById('logoPreview');
    primaryColorInput = document.getElementById('primaryColor');
    accentColorInput = document.getElementById('accentColor');
    remoteConfigStatus = document.getElementById('remote-config-status');
    remoteConfigMeta = document.getElementById('remote-config-meta');
    loadCloudConfigBtn = document.getElementById('load-cloud-config-btn');
    saveCloudConfigBtn = document.getElementById('save-cloud-config-btn');
    autoSyncCloudToggle = document.getElementById('autoSyncCloud');

    contractClientSelect = document.getElementById('contract-client-select');
    contractQuoteSelect = document.getElementById('contract-quote-select');
    contractText = document.getElementById('contract-text');
    saveContractBtn = document.getElementById('save-contract-btn');

    portalClientSelect = document.getElementById('portal-client-select');
    clientPortalContent = document.getElementById('client-portal-content');

    // --- AÑADIR EVENT LISTENERS ---
     // Añadir listeners SOLO si el elemento existe para evitar errores
    if (loginButton) loginButton.addEventListener('click', handleAuthClick);
    if (logoutButton) logoutButton.addEventListener('click', handleSignoutClick);
    if (clientForm) clientForm.addEventListener('submit', handleClientFormSubmit);
    if (serviceForm) serviceForm.addEventListener('submit', handleServiceFormSubmit);
    if (saveContractBtn) saveContractBtn.addEventListener('click', handleSaveContractClick);
    if (configForm) configForm.addEventListener('submit', handleConfigFormSubmit);
    if (logoUploadInput) logoUploadInput.addEventListener('change', handleLogoUploadChange);
    if (quoteClientSelect) quoteClientSelect.addEventListener('change', handleQuoteClientChange);
    if (serviceSelect) serviceSelect.addEventListener('change', handleServiceSelectChange);
    if (contractClientSelect) contractClientSelect.addEventListener('change', handleContractClientChange);
    if (contractQuoteSelect) contractQuoteSelect.addEventListener('change', handleContractQuoteChange);
    if (portalClientSelect) portalClientSelect.addEventListener('change', handlePortalClientChange);
    if (autoSyncCloudToggle) {
        autoSyncCloudToggle.checked = autoSyncCloud;
        autoSyncCloudToggle.addEventListener('change', handleAutoSyncToggleChange);
    }
    if (saveCloudConfigBtn) saveCloudConfigBtn.addEventListener('click', handleManualCloudSave);
    if (loadCloudConfigBtn) loadCloudConfigBtn.addEventListener('click', handleManualCloudLoad);

    syncServicePriceOverrideField();
    initializeRemoteConfigPanel();
    updateQuoteActionButtonsState();
}

// --- 2. LÓGICA DE AUTENTICACIÓN (AUTH) - REESTRUCTURADA ---

// Función principal para inicializar GAPI y GIS
async function initializeApp() {
    console.log("initializeApp: Iniciando...");
    if (authStatus) authStatus.innerText = 'Cargando Google API Client (GAPI)...';

    // Cargar GAPI Client
    await new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://apis.google.com/js/api.js';
        script.async = true;
        script.defer = true;
        script.onload = () => {
            console.log("GAPI script cargado.");
            // Ahora cargar la biblioteca 'client' de GAPI
            gapi.load('client', {
                 callback: async () => {
                     console.log("GAPI 'client' library cargada.");
                     try {
                         await gapi.client.init({
                             // No se necesita apiKey si solo usamos OAuth
                             discoveryDocs: [
                                 'https://sheets.googleapis.com/$discovery/rest?version=v4',
                                 'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest',
                                 'https://docs.googleapis.com/$discovery/rest?version=v1'
                             ],
                         });
                         console.log("GAPI client inicializado.");
                         resolve(); // Resuelve la promesa GAPI
                     } catch (err) {
                          console.error("Error inicializando GAPI client:", err);
                          if (authStatus) authStatus.innerText = `Error GAPI: ${err.message || 'Desconocido'}`;
                          reject(err);
                      }
                 },
                 onerror: (err) => {
                      console.error("Error cargando GAPI 'client' library:", err);
                      if (authStatus) authStatus.innerText = 'Error cargando GAPI client library.';
                      reject(err);
                  }
             });
        };
         script.onerror = (err) => {
             console.error("Error cargando script api.js:", err);
             if (authStatus) authStatus.innerText = 'Error al cargar GAPI script.';
             reject(err);
         };
        document.body.appendChild(script);
    });

    // GAPI está listo, ahora inicializar GIS
     if (authStatus) authStatus.innerText = 'Cargando Google Sign-In (GIS)...';
     console.log("Inicializando GIS...");

    await new Promise((resolve, reject) => {
         const script = document.createElement('script');
         script.src = 'https://accounts.google.com/gsi/client';
         script.async = true;
         script.defer = true;
         script.onload = () => {
             console.log("GIS script cargado.");
             try {
                 tokenClient = google.accounts.oauth2.initTokenClient({
                     client_id: CLIENT_ID,
                     scope: SCOPES,
                     callback: tokenCallback, // Función que se llama después de que el usuario inicia sesión
                 });
                 console.log("GIS tokenClient inicializado.");
                 resolve(); // Resuelve la promesa GIS
             } catch (err) {
                  console.error("Error inicializando GIS tokenClient:", err);
                  if (authStatus) authStatus.innerText = `Error GIS: ${err.message || 'Desconocido'}`;
                  reject(err);
              }
         };
         script.onerror = (err) => {
             console.error("Error cargando script gsi/client:", err);
             if (authStatus) authStatus.innerText = 'Error al cargar GIS script.';
             reject(err);
         };
         document.body.appendChild(script);
    });

    // Ambas bibliotecas están listas
    console.log("GAPI y GIS listos.");
    if (authStatus) authStatus.innerText = "Listo para iniciar sesión.";
    if (loginButton) loginButton.disabled = false; // Habilitar el botón
}

// Se llama cuando el usuario hace clic en "Iniciar Sesión"
function handleAuthClick() {
    if (loadingSpinner) loadingSpinner.style.display = 'block';
    if (authStatus) authStatus.innerText = 'Abriendo ventana de Google...';
    if (tokenClient) {
        console.log("Solicitando token de acceso...");
        // Solicitar token. El callback 'tokenCallback' se ejecutará si tiene éxito.
        tokenClient.requestAccessToken({prompt: 'consent'}); // Forzar consentimiento la primera vez o si expira
    } else {
        console.error("handleAuthClick: TokenClient no está listo.");
        if (authStatus) authStatus.innerText = 'Error: Cliente de Google no inicializado.';
         if (loadingSpinner) loadingSpinner.style.display = 'none';
    }
}

// Se llama cuando el usuario hace clic en "Cerrar Sesión"
function handleSignoutClick() {
    const token = gapi.client?.getToken(); // Usar optional chaining
    if (token && token.access_token) {
        console.log("Revocando token...");
        google.accounts.oauth2.revoke(token.access_token, () => {
            console.log("Token revocado.");
            gapi.client.setToken(null); // Limpiar token de GAPI
            // Ocultar app, mostrar login
            if (appContainer) appContainer.style.display = 'none';
            showLoginOverlay();
            // Limpiar email
            const userEmailEl = document.getElementById('user-email');
            if (userEmailEl) userEmailEl.innerText = '';
            // Limpiar datos de la aplicación
            clearAppData();
            signedInUser = null;
            resetRemoteConfigState();
            if (authStatus) authStatus.innerText = 'Sesión cerrada. Listo para iniciar sesión.';
            if (loginButton) loginButton.disabled = false; // Asegurarse que el botón esté habilitado
        });
    } else {
         console.warn("handleSignoutClick: No hay token para revocar o GAPI no listo.");
         // Aún así, forzar el estado de cierre de sesión visualmente
         if (appContainer) appContainer.style.display = 'none';
         showLoginOverlay();
        const userEmailEl = document.getElementById('user-email');
        if (userEmailEl) userEmailEl.innerText = '';
        clearAppData();
        signedInUser = null;
        resetRemoteConfigState();
        if (authStatus) authStatus.innerText = 'Listo para iniciar sesión.';
        if (loginButton) loginButton.disabled = false;
    }
}

// Función para limpiar datos al cerrar sesión
function clearAppData() {
     clients = [];
     services = [];
     quotesHistory = [];
     contracts = [];
    quoteItemsData = [];
    currentQuoteContextId = null;
     if(clientList) clientList.innerHTML = ''; // Limpiar tablas visualmente
     if(serviceList) serviceList.innerHTML = '';
     if(quotesHistoryList) quotesHistoryList.innerHTML = '';
     if(quoteItemsBody) quoteItemsBody.innerHTML = '';
     // Resetear selects si existen
     if(quoteClientSelect) quoteClientSelect.innerHTML = '<option value="">-- Seleccione --</option>';
     if(serviceSelect) serviceSelect.innerHTML = '<option value="">-- Seleccione --</option>';
     // Podrías resetear más elementos si es necesario
    console.log("Datos de la aplicación limpiados.");
    updateQuoteActionButtonsState();
}

// Callback principal después de que el usuario inicia sesión exitosamente.
async function tokenCallback(tokenResponse) {
     console.log("tokenCallback recibido:", tokenResponse);
    if (tokenResponse.error) {
        console.error("Error en la respuesta del token:", tokenResponse.error);
        if (authStatus) authStatus.innerText = `Error: ${tokenResponse.error_description || tokenResponse.error}`;
        if (loadingSpinner) loadingSpinner.style.display = 'none';
        if (loginButton) loginButton.disabled = false; // Habilitar botón si falla
        return;
    }

    // Guardar el token en GAPI para usar las APIs
    gapi.client.setToken(tokenResponse);
    console.log("Token de acceso establecido en GAPI.");

    if (authStatus) authStatus.innerText = 'Autenticado. Obteniendo info de usuario...';

    signedInUser = { email: '', name: '' };

    // Obtener info del usuario (opcional pero útil)
    try {
        const userInfoResponse = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
            headers: { 'Authorization': `Bearer ${tokenResponse.access_token}` }
        });
        if (!userInfoResponse.ok) {
             throw new Error(`Error ${userInfoResponse.status}: ${await userInfoResponse.text()}`);
         }
        const userData = await userInfoResponse.json();
        console.log("Información del usuario:", userData);
        signedInUser = {
            email: userData.email || '',
            name: userData.name || userData.given_name || userData.family_name || ''
        };
        const userEmailEl = document.getElementById('user-email');
         if (userEmailEl) userEmailEl.innerText = userData.email || 'Usuario';
    } catch (err) {
        console.warn("No se pudo obtener info del usuario:", err);
         const userEmailEl = document.getElementById('user-email');
         // Poner un placeholder si falla
         if (userEmailEl) userEmailEl.innerText = 'Usuario (email no disponible)';
    }

    refreshRemoteConfigUiState();

    // Ocultar login, mostrar app
    hideLoginOverlay();
    if (appContainer) appContainer.style.display = 'block';
    if (loadingSpinner) loadingSpinner.style.display = 'none'; // Ocultar spinner

    // Cargar datos iniciales desde Google Sheets/Drive
    if (authStatus) authStatus.innerText = 'Cargando datos iniciales...';
    await initializeDataFromGoogle();
}

// Función para verificar si falta alguna configuración
function checkSheetConfig() {
    let missing = [];
    if (!CLIENTS_SHEET_ID || CLIENTS_SHEET_ID === 'ID_DE_TU_HOJA_DE_CLIENTES') missing.push("Clientes");
    if (!SERVICES_SHEET_ID || SERVICES_SHEET_ID === 'ID_DE_TU_HOJA_DE_SERVICIOS') missing.push("Servicios");
    if (!QUOTES_SHEET_ID || QUOTES_SHEET_ID === 'ID_DE_TU_HOJA_DE_COTIZACIONES') missing.push("Historial Cotizaciones");
    if (!CONTRACT_TEMPLATE_ID || CONTRACT_TEMPLATE_ID === 'ID_DE_TU_PLANTILLA_DE_CONTRATO_EN_GOOGLE_DOCS') missing.push("Plantilla Contrato");

    if (missing.length > 0) {
         console.warn(message);
         showAlert(message, "Error de Configuración"); // REEMPLAZO DE ALERT
         return false; // Indicar que falta configuración
    }
    return true; // Configuración completa
}

// --- 3. LÓGICA DE DATOS (GOOGLE SHEETS) ---

// Constantes y Estado de la Aplicación
const IVA_RATE = 0.15;
const DEFAULT_COMPANY_SETTINGS = {
    name: 'Tu Empresa S.A.',
    address: 'Tu Dirección, Guayaquil',
    contact: 'tuemail@empresa.com',
    logo: 'https://placehold.co/200x100/eef2ff/4f46e5?text=Tu+Logo',
    primaryColor: '#1a202c',
    accentColor: '#4f46e5',
    website: '',
    whatsapp: ''
};

let clients = [];
let services = [];
let quotesHistory = [];
let contracts = []; // Podría usarse para cachear contratos cargados
let quoteItemsData = [];
let quoteCounter = 1;

let autoSyncCloud = false;
try {
    autoSyncCloud = localStorage.getItem('companySettingsAutoSyncCloud') === 'true';
} catch (error) {
    console.warn('No se pudo leer la preferencia de sincronización automática:', error);
}

let companySettings = readCompanySettingsFromLocalStorage();
let companySettingsRowIndex = null;
let remoteConfigLastLoadedAt = null;
let remoteConfigLastSavedAt = null;
let remoteConfigBusy = false;
let currentQuoteContextId = null;

// Función principal de inicialización de datos (llamada después del login)
async function initializeDataFromGoogle() {
    // No continuar si falta configuración esencial (Sheets IDs)
    if (!checkSheetConfig()) { 
        if (authStatus) authStatus.innerText = 'Error: Configuración incompleta.';
        return; 
    } 

    if (authStatus) authStatus.innerText = 'Cargando Clientes...';
    await loadClientsFromSheet();

    if (authStatus) authStatus.innerText = 'Cargando Servicios...';
    await loadServicesFromSheet();

    if (authStatus) authStatus.innerText = 'Cargando Historial...';
    await loadQuotesHistoryFromSheet(); // Esto también actualiza quoteCounter

     if (authStatus) authStatus.innerText = 'Datos cargados. Listo.';

    // Renderizar UI con los datos cargados
    renderClients();
    renderServices();
    renderQuoteItems(); // Renderiza la tabla de items (vacía inicialmente)
    updateQuoteNumberPreview(); // Calcula el número de cotización
    loadSettings(); // Carga config. local (nombre empresa, logo, etc.)
    refreshRemoteConfigUiState();
    await loadCompanySettingsFromCloud({ silent: true });
    renderQuotesHistory(); // Renderiza la tabla de historial

    // Asegurar que la primera pestaña esté activa visualmente
    const firstTabButton = document.querySelector('.tab-button');
    const firstTabContent = document.querySelector('.tab-content');
    if (firstTabButton && firstTabContent) {
         document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
         document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
         firstTabButton.classList.add('active');
         firstTabContent.classList.add('active');
     } else {
          console.warn("No se encontró la primera pestaña o contenido para activar.");
      }
}

// --- LÓGICA DE CLIENTES (GOOGLE SHEETS) ---

async function loadClientsFromSheet() {
     // Verificar que gapi.client y gapi.client.sheets existan
     if (!gapi?.client?.sheets) { 
         console.error("loadClientsFromSheet: API de Sheets no está lista."); 
         showAlert("Error: La API de Google Sheets no está lista. Intenta recargar.", "Error de API"); // REEMPLAZO DE ALERT
         return; 
     }
    try {
        console.log(`Cargando clientes desde ${CLIENTS_SHEET_ID} rango ${CLIENTS_SHEET_RANGE}`);
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: CLIENTS_SHEET_ID,
            range: CLIENTS_SHEET_RANGE,
        });
        console.log("Respuesta de Sheets (Clientes):", response.result);
        // Mapeo cuidadoso, verificando si response.result.values existe
        clients = (response.result.values || []).map((row, index) => {
            // Añadir chequeos por si una fila es más corta de lo esperado
            const ruc = row[0] || '';
            const name = row[1] || '';
            const contact = row[5] || row[4] || ''; // Prioriza email (F) sobre teléfono (E)
            return {
                rowId: index + 2, // Fila real en la hoja (A2 es fila 2)
                id: ruc || `client_row_${index + 2}`, // Usar RUC como ID si existe, sino un ID basado en fila
                code: ruc,             
                name: name,             
                ruc: ruc,            
                contact: contact    
            };
        }).filter(client => client.ruc || client.name); // Filtrar filas completamente vacías

        console.log(`Clientes cargados: ${clients.length}`);
    } catch (err) {
        console.error("Error cargando clientes:", err);
        const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
        showAlert(`Error al cargar clientes: ${errorMsg}\nAsegúrate de que la hoja exista, tengas permiso y el rango sea correcto.`, "Error de Carga"); // REEMPLAZO DE ALERT
        clients = []; // Resetear si falla
    }
}

// Renderizar Clientes 
function renderClients() {
     if (!clientList) {console.warn("Elemento clientList no encontrado para renderizar."); return;} 
    clientList.innerHTML = ''; // Limpiar tabla
    if (clients.length === 0) {
        clientList.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-muted">No hay clientes para mostrar. Añade uno o revisa la conexión con Google Sheets.</td></tr>`;
    } else {
        // Ordenar por nombre antes de renderizar
        [...clients].sort((a,b) => (a.name || '').localeCompare(b.name || '')).forEach(client => {
            const row = document.createElement('tr');
            // Usamos client.id (RUC) para editar y client.rowId para borrar
            row.innerHTML = `
                <td class="px-3 py-2 text-nowrap font-monospace">${client.ruc || 'N/A'}</td>
                <td class="px-3 py-2 text-nowrap">${client.name || 'N/A'}</td>
                <td class="px-3 py-2 text-nowrap">${client.contact || ''}</td>
                <td class="px-3 py-2 text-end">
                    <button onclick="editClient('${client.id}')" class="btn btn-sm btn-outline-primary me-2" title="Editar Cliente"><i class="fas fa-edit"></i></button>
                    <button onclick="deleteClient('${client.rowId}')" class="btn btn-sm btn-outline-danger" title="Eliminar Cliente (limpia la fila)"><i class="fas fa-trash"></i></button>
                </td>`;
            clientList.appendChild(row);
        });
    }
    updateClientDropdowns(); // Actualizar los <select> que usan clientes
}

// Guardar/Actualizar Cliente 
async function handleClientFormSubmit(e) {
    e.preventDefault();
     // Chequeo robusto de elementos
     if (!clientRowIdInput || !clientRUCInput || !clientNameInput || !clientContactInput || !clientForm || !gapi?.client?.sheets) {
         showAlert("Error: Faltan elementos del formulario o la API de Sheets no está lista.", "Error de Formulario"); // REEMPLAZO
         return; 
     }
    const rowId = clientRowIdInput.value; // ID de fila para saber si es update o append
    const ruc = clientRUCInput.value.trim();
    const name = clientNameInput.value.trim();
    const contact = clientContactInput.value.trim();

    if (!ruc || !name) {
         showAlert("El RUC/CI y la Razón Social son obligatorios.", "Datos Incompletos"); // REEMPLAZO
         return;
     }

    // Mapeo a columnas A-F
    const isEmail = contact.includes('@');
    const clientData = [
        ruc,       // Col A: Identificación No.
        name,      // Col B: Razón Social
        name,      // Col C: Nombre Comercial (default)
        '',        // Col D: Dirección (vacía)
        isEmail ? '' : contact, // Col E: Teléfono
        isEmail ? contact : ''  // Col F: E-mail
    ];

    // Deshabilitar botón mientras guarda
    const submitButton = clientForm.querySelector('button[type="submit"]');
    if (submitButton) submitButton.disabled = true;
    if (authStatus) authStatus.innerText = rowId ? 'Actualizando cliente...' : 'Guardando nuevo cliente...';

    try {
        if (rowId) {
            // --- ACTUALIZAR (UPDATE) ---
            const range = `Hoja 1!A${rowId}:F${rowId}`;
            console.log(`Actualizando cliente en ${CLIENTS_SHEET_ID} rango ${range}`);
            await gapi.client.sheets.spreadsheets.values.update({
                spreadsheetId: CLIENTS_SHEET_ID,
                range: range,
                valueInputOption: 'USER_ENTERED',
                resource: { values: [clientData] }
            });
            showAlert('Cliente actualizado.'); // REEMPLAZO
        } else {
            // --- CREAR (APPEND) ---
             console.log(`Añadiendo cliente a ${CLIENTS_SHEET_ID}`);
            await gapi.client.sheets.spreadsheets.values.append({
                spreadsheetId: CLIENTS_SHEET_ID,
                range: 'Hoja 1!A:F', // Apuntar al rango completo para append
                valueInputOption: 'USER_ENTERED',
                insertDataOption: 'INSERT_ROWS', 
                resource: { values: [clientData] }
            });
            showAlert('Nuevo cliente guardado.'); // REEMPLAZO
        }

        clientForm.reset();
        clientRowIdInput.value = ''; // Limpiar ID de fila
        await loadClientsFromSheet(); // Recargar datos
        renderClients(); // Actualizar tabla
        if (authStatus) authStatus.innerText = 'Listo.';

    } catch (err) {
        console.error("Error guardando cliente:", err);
        const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
        showAlert(`Error al guardar cliente: ${errorMsg}\nRevisa la consola.`, "Error al Guardar"); // REEMPLAZO
         if (authStatus) authStatus.innerText = 'Error al guardar.';
    } finally {
         // Rehabilitar botón
         if (submitButton) submitButton.disabled = false;
     }
}

// Editar Cliente (carga datos en el formulario)
function editClient(id) {
    const client = clients.find(c => c.id === id);
    // Chequeo robusto
    if (!client || !clientRowIdInput || !clientNameInput || !clientRUCInput || !clientContactInput) {
         console.error("editClient: No se encontró el cliente o faltan elementos del form. ID:", id);
         showAlert("No se encontró el cliente para editar."); // REEMPLAZO
         return; 
    }
    console.log("Editando cliente:", client);
    clientRowIdInput.value = client.rowId; // Guardar la fila para el update
    clientNameInput.value = client.name;
    clientRUCInput.value = client.ruc;
    clientContactInput.value = client.contact;
     // Cambiar a la pestaña de clientes
     const tabButton = document.querySelector('.tab-button[onclick*="clientes"]');
     if (tabButton) {
         changeTab({currentTarget: tabButton}, 'clientes');
     }
    clientNameInput.focus(); // Poner foco en el primer campo editable
    window.scrollTo(0, 0); // Ir al inicio de la página para ver el formulario
}

// Borrar Cliente (limpia la fila)
async function deleteClient(rowId) {
     if (!gapi?.client?.sheets) {
         showAlert("Error: API de Sheets no lista para borrar.", "Error de API"); // REEMPLAZO
         return;
     }

     // REEMPLAZO DE CONFIRM
     showConfirm(`¿Estás seguro de ELIMINAR los datos del cliente en la fila ${rowId}?\nEsta acción limpiará las celdas en Google Sheets.`, async () => {
        if (authStatus) authStatus.innerText = `Eliminando fila ${rowId}...`;

        try {
            const range = `Hoja 1!A${rowId}:F${rowId}`; // Rango a limpiar
             console.log(`Borrando datos en ${CLIENTS_SHEET_ID} rango ${range}`);
            await gapi.client.sheets.spreadsheets.values.clear({
                spreadsheetId: CLIENTS_SHEET_ID,
                range: range,
            });
            showAlert('Datos del cliente eliminados de la fila.'); // REEMPLAZO

            // Quitar de la lista local y recargar/renderizar
            clients = clients.filter(c => c.rowId !== parseInt(rowId)); // Actualizar array local
            renderClients(); // Redibujar tabla
            if (authStatus) authStatus.innerText = 'Listo.';

        } catch (err) {
            console.error("Error borrando cliente:", err);
            const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
            showAlert(`Error al eliminar cliente: ${errorMsg}\nRevisa la consola.`, "Error al Eliminar"); // REEMPLAZO
             if (authStatus) authStatus.innerText = 'Error al eliminar.';
             // Opcional: Recargar todo si falla para asegurar consistencia
             // await loadClientsFromSheet(); renderClients();
         }
     });
}

// --- LÓGICA DE SERVICIOS (GOOGLE SHEETS) ---

async function loadServicesFromSheet() {
     if (!gapi?.client?.sheets) { console.error("loadServicesFromSheet: API de Sheets no lista."); return; }
    try {
        console.log(`Cargando servicios desde ${SERVICES_SHEET_ID} rango ${SERVICES_SHEET_RANGE}`);
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SERVICES_SHEET_ID,
            range: SERVICES_SHEET_RANGE, // A2:E
        });
        console.log("Respuesta de Sheets (Servicios):", response.result);
        services = (response.result.values || []).map((row, index) => {
            const id = row[0] || '';
            const name = row[2] || ''; // Col C: description
            const priceStr = row[3] || '0'; // Col D: price
            // Limpieza robusta del precio
            const price = parseFloat(String(priceStr).replace(/[^0-9.-]+/g,"")) || 0; 
            return {
                rowId: index + 2,
                id: id || `service_row_${index + 2}`, // Usar ID de Col A si existe
                name: name,             
                description: name, // Usar descripción como ambos por ahora       
                price: price 
            };
        }).filter(service => service.id || service.name); // Filtrar filas vacías

        console.log(`Servicios cargados: ${services.length}`);
    } catch (err) {
        console.error("Error cargando servicios:", err);
        const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
        showAlert(`Error al cargar servicios: ${errorMsg}\nAsegúrate de que la hoja exista, tengas permiso y el rango sea correcto.`, "Error de Carga"); // REEMPLAZO
        services = []; // Resetear si falla
    }
}

function renderServices() {
     if (!serviceList) {console.warn("Elemento serviceList no encontrado."); return;}
    serviceList.innerHTML = ''; // Limpiar
     if (services.length === 0) {
         serviceList.innerHTML = `<tr><td colspan="3" class="text-center py-4 text-muted">No hay servicios para mostrar. Añade uno o revisa la conexión con Google Sheets.</td></tr>`;
     } else {
         // Ordenar por nombre
         [...services].sort((a,b) => (a.name || '').localeCompare(b.name || '')).forEach(service => {
            const row = document.createElement('tr');
            // Usar service.id para editar, service.rowId para borrar
            row.innerHTML = `
                <td class="px-3 py-2 text-nowrap">${service.name || 'N/A'}</td>
                <td class="px-3 py-2 text-nowrap">$${(service.price || 0).toFixed(2)}</td>
                <td class="px-3 py-2 text-end">
                    <button onclick="editService('${service.id}')" class="btn btn-sm btn-outline-primary me-2" title="Editar Servicio"><i class="fas fa-edit"></i></button>
                    <button onclick="deleteService('${service.rowId}')" class="btn btn-sm btn-outline-danger" title="Eliminar Servicio (limpia la fila)"><i class="fas fa-trash"></i></button>
                </td>`;
            serviceList.appendChild(row);
        });
    }
    updateServiceDropdown(); // Actualizar el <select> del cotizador
}

async function handleServiceFormSubmit(e) {
    e.preventDefault();
     // Chequeo robusto
     if (!serviceRowIdInput || !serviceIdInput || !serviceNameInput || !servicePriceInput || !serviceForm || !gapi?.client?.sheets) {
         showAlert("Error: Faltan elementos del formulario de servicios o API no lista.", "Error de Formulario"); // REEMPLAZO
         return;
     }
    const rowId = serviceRowIdInput.value;
    const name = serviceNameInput.value.trim();
    const price = parseFloat(servicePriceInput.value);

     if (!name || isNaN(price) || price < 0) {
         showAlert("El nombre y un precio válido (mayor o igual a 0) son obligatorios.", "Datos Incompletos"); // REEMPLAZO
         return;
      }

    // Mapeo a columnas A-E
    const serviceIdValue = rowId ? serviceIdInput.value : `service_${new Date().getTime()}`; // Generar nuevo ID si no existe
    const serviceData = [
        serviceIdValue,         // Col A: id
        '',                     // Col B: code (vacío)
        name,                   // Col C: description
        price,                  // Col D: price
        new Date().toISOString() // Col E: updated_at
    ];

    // Deshabilitar botón
    const submitButton = serviceForm.querySelector('button[type="submit"]');
    if (submitButton) submitButton.disabled = true;
    if (authStatus) authStatus.innerText = rowId ? 'Actualizando servicio...' : 'Guardando nuevo servicio...';

    try {
        if (rowId) {
            // --- ACTUALIZAR (UPDATE) ---
            const range = `Hoja 1!A${rowId}:E${rowId}`;
            console.log(`Actualizando servicio en ${SERVICES_SHEET_ID} rango ${range}`);
            await gapi.client.sheets.spreadsheets.values.update({
                spreadsheetId: SERVICES_SHEET_ID,
                range: range,
                valueInputOption: 'USER_ENTERED',
                resource: { values: [serviceData] }
            });
            showAlert('Servicio actualizado.'); // REEMPLAZO
        } else {
            // --- CREAR (APPEND) ---
             console.log(`Añadiendo servicio a ${SERVICES_SHEET_ID}`);
            await gapi.client.sheets.spreadsheets.values.append({
                spreadsheetId: SERVICES_SHEET_ID,
                range: 'Hoja 1!A:E', 
                valueInputOption: 'USER_ENTERED',
                insertDataOption: 'INSERT_ROWS',
                resource: { values: [serviceData] }
            });
            showAlert('Nuevo servicio guardado.'); // REEMPLAZO
        }

        serviceForm.reset();
        serviceRowIdInput.value = '';
        serviceIdInput.value = ''; // Limpiar ID oculto
        await loadServicesFromSheet(); // Recargar
        renderServices(); // Actualizar tabla
        if (authStatus) authStatus.innerText = 'Listo.';

    } catch (err) {
        console.error("Error guardando servicio:", err);
        const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
        showAlert(`Error al guardar servicio: ${errorMsg}\nRevisa la consola.`, "Error al Guardar"); // REEMPLAZO
         if (authStatus) authStatus.innerText = 'Error al guardar.';
    } finally {
         if (submitButton) submitButton.disabled = false; // Rehabilitar
     }
}

function editService(id) {
    const service = services.find(s => s.id === id);
    // Chequeo robusto
    if (!service || !serviceRowIdInput || !serviceIdInput || !serviceNameInput || !serviceDescriptionInput || !servicePriceInput) {
        console.error("editService: No se encontró servicio o faltan elementos del form. ID:", id);
        showAlert("No se encontró el servicio para editar."); // REEMPLAZO
         return;
    }
    console.log("Editando servicio:", service);
    serviceRowIdInput.value = service.rowId; // Guardar fila para update
    serviceIdInput.value = service.id; // Guardar ID existente
    serviceNameInput.value = service.name; // description -> name input
    serviceDescriptionInput.value = service.description; // Llenar campo oculto (por si acaso)
    servicePriceInput.value = service.price;
     // Cambiar a pestaña
     const tabButton = document.querySelector('.tab-button[onclick*="servicios"]');
     if (tabButton) {
         changeTab({currentTarget: tabButton}, 'servicios');
     }
    serviceNameInput.focus(); // Foco
    window.scrollTo(0, 0); // Ir arriba
}

async function deleteService(rowId) {
     if (!gapi?.client?.sheets) {
         showAlert("Error: API de Sheets no lista para borrar.", "Error de API"); // REEMPLAZO
         return;
     }

     // REEMPLAZO DE CONFIRM
     showConfirm(`¿Estás seguro de ELIMINAR los datos del servicio en la fila ${rowId}?\nEsto limpiará las celdas en Google Sheets.`, async () => {
        if (authStatus) authStatus.innerText = `Eliminando fila ${rowId}...`;

        try {
            const range = `Hoja 1!A${rowId}:E${rowId}`; // Rango A-E
             console.log(`Borrando datos en ${SERVICES_SHEET_ID} rango ${range}`);
            await gapi.client.sheets.spreadsheets.values.clear({
                spreadsheetId: SERVICES_SHEET_ID,
                range: range,
            });
            showAlert('Datos del servicio eliminados de la fila.'); // REEMPLAZO

            // Actualizar localmente
            services = services.filter(s => s.rowId !== parseInt(rowId)); 
            renderServices(); // Redibujar
            if (authStatus) authStatus.innerText = 'Listo.';

        } catch (err) {
            console.error("Error borrando servicio:", err);
            const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
            showAlert(`Error al eliminar servicio: ${errorMsg}\nRevisa la consola.`, "Error al Eliminar"); // REEMPLAZO
             if (authStatus) authStatus.innerText = 'Error al eliminar.';
             // Opcional: Recargar todo
             // await loadServicesFromSheet(); renderServices();
         }
     });
}

// --- LÓGICA DE HISTORIAL (GOOGLE SHEETS) ---

async function loadQuotesHistoryFromSheet() {
     // Verificar si el ID es el placeholder o si la API no está lista
     if (!QUOTES_SHEET_ID || QUOTES_SHEET_ID === 'ID_DE_TU_HOJA_DE_COTIZACIONES' || !gapi?.client?.sheets) { 
         console.warn("loadQuotesHistoryFromSheet: ID de hoja no configurado o API no lista. Saltando carga.");
         quotesHistory = []; // Asegurar que esté vacío
         quoteCounter = 1;
         renderQuotesHistory(); // Mostrar mensaje de "no cargado"
         return; 
     }

    try {
        console.log(`Cargando historial desde ${QUOTES_SHEET_ID} rango ${QUOTES_SHEET_RANGE}`);
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: QUOTES_SHEET_ID,
            range: QUOTES_SHEET_RANGE, // A2:J
        });
        console.log("Respuesta de Sheets (Historial):", response.result);

        quotesHistory = (response.result.values || []).map((row, index) => {
            // Validar longitud mínima de fila (hasta Estado - Col H)
            if (row.length < 8) { 
                 console.warn(`Fila ${index + 2} incompleta en Historial (longitud ${row.length}):`, row);
                 return null; // Ignorar fila incompleta
             }
            try {
                // Parsear JSON con cuidado
                const clientData = JSON.parse(row[4] || '{}');
                const itemsData = JSON.parse(row[5] || '[]');

                // Validar que items sea un array
                 if (!Array.isArray(itemsData)) {
                     console.warn(`Items en fila ${index + 2} no es un array:`, row[5]);
                     throw new Error("Formato de items inválido."); // Lanzar error para que se filtre
                  }

                return {
                    rowId: index + 2,
                    id: row[0] || `quote_row_${index + 2}`, // Col A
                    number: row[1] || 'N/A', // Col B
                    issueDate: row[2] || 'N/A', // Col C
                    validityDate: row[3] || 'N/A', // Col D
                    client: clientData, // Col E (JSON)
                    items: itemsData, // Col F (JSON Array)
                    total: parseFloat(row[6]) || 0, // Col G
                    status: row[7] || 'Pendiente', // Col H
                    notes: row[8] || '', // Col I
                    googleDocId: row[9] || null, // Col J
                    googlePdfId: row[10] || null, // Col K
                    companySettingsForPrint: null
                }
            } catch(e) {
                 // Capturar errores de JSON.parse u otros
                 console.error(`Error procesando fila ${index + 2} de Historial:`, e.message, "Fila:", row);
                 return null; // Marcar fila como inválida
             }
        }).filter(q => q !== null); // Filtrar filas nulas (incompletas o con error)

        quoteCounter = quotesHistory.length + 1; // Contador basado en filas válidas
         console.log(`Historial cargado: ${quotesHistory.length} registros válidos.`);
    } catch (err) {
         console.error("Error fatal cargando historial:", err);
         const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
         showAlert(`Error CRÍTICO al cargar el historial: ${errorMsg}\nRevisa permisos, ID de hoja y rango. La app podría no funcionar correctamente.`, "Error Crítico"); // REEMPLAZO
         quotesHistory = []; // Resetear si falla gravemente
         quoteCounter = 1;
     }
     // Renderizar siempre para mostrar estado (vacío, error o datos)
    renderQuotesHistory();
    updateQuoteActionButtonsState();
}

// Renderizar tabla de historial
function renderQuotesHistory() {
    if (!quotesHistoryList) {console.warn("Elemento quotesHistoryList no encontrado."); return;}
    quotesHistoryList.innerHTML = ''; // Limpiar
    if (quotesHistory.length === 0) {
        quotesHistoryList.innerHTML = `<tr><td colspan="6" class="text-center py-4 text-muted">No hay cotizaciones registradas en Google Sheets o no se pudieron cargar.</td></tr>`;
    } else {
        // Mostrar las más recientes primero
        [...quotesHistory].reverse().forEach(quote => {
            const row = document.createElement('tr');
            // Usar fallbacks si client o name no existen
            const clientName = quote.client?.name || 'Cliente Desconocido';
            const pdfLinkHtml = quote.googlePdfId
                ? `<a href="https://drive.google.com/file/d/${quote.googlePdfId}/view" target="_blank" class="btn btn-sm btn-outline-danger me-2" title="Abrir PDF en Google Drive"><i class="fas fa-file-pdf"></i></a>`
                : '';

            row.innerHTML = `
                <td class="px-3 py-2 text-nowrap font-monospace">${quote.number || 'N/A'}</td>
                <td class="px-3 py-2 text-nowrap">${clientName}</td>
                <td class="px-3 py-2 text-nowrap">${quote.issueDate || 'N/A'}</td>
                <td class="px-3 py-2 text-nowrap fw-semibold">$${(quote.total || 0).toFixed(2)}</td>
                <td class="px-3 py-2 text-nowrap">
                    <select onchange="updateQuoteStatus('${quote.id}', this.value, ${quote.rowId})"
                            class="form-select form-select-sm status-${quote.status || 'Pendiente'}"
                            aria-label="Estado de la cotización ${quote.number}">
                        <option value="Pendiente" ${quote.status === 'Pendiente' ? 'selected' : ''}>Pendiente</option>
                        <option value="Aceptada" ${quote.status === 'Aceptada' ? 'selected' : ''}>Aceptada</option>
                        <option value="Rechazada" ${quote.status === 'Rechazada' ? 'selected' : ''}>Rechazada</option>
                    </select>
                </td>
                <td class="px-3 py-2 text-end">
                    <button onclick="loadQuoteForEdit('${quote.id}')" class="btn btn-sm btn-outline-secondary me-2" title="Cargar cotización para editarla"><i class="fas fa-pencil-alt"></i></button>
                    <button onclick="reprintQuote('${quote.id}')" class="btn btn-sm btn-outline-primary me-2" title="Ver / Reimprimir cotización"><i class="fas fa-eye"></i></button>
                    ${pdfLinkHtml}
                    <button onclick="deleteQuote('${quote.id}', ${quote.rowId}, '${quote.googleDocId || ''}', '${quote.googlePdfId || ''}')" class="btn btn-sm btn-outline-danger ms-2" title="Eliminar cotización"><i class="fas fa-trash"></i></button>
                </td>`;
            quotesHistoryList.appendChild(row);
        });
    }
    updateQuoteActionButtonsState();
}

// Actualizar estado de cotización en Google Sheet
async function updateQuoteStatus(quoteId, newStatus, rowId) {
    // Revalidar ID y API
    if (!QUOTES_SHEET_ID || QUOTES_SHEET_ID === 'ID_DE_TU_HOJA_DE_COTIZACIONES' || !rowId || !gapi?.client?.sheets) {
        showAlert("Error: ID de Hoja de Cotizaciones no configurado, fila inválida o API no lista.", "Error de Configuración"); // REEMPLAZO
        // Revertir visualmente el select si es posible
         const selectElement = event?.target; // Intentar obtener el select que disparó el evento
          if (selectElement) {
               const originalQuote = quotesHistory.find(q => q.id === quoteId);
               if (originalQuote) selectElement.value = originalQuote.status;
           }
         return;
     }

     if (authStatus) authStatus.innerText = `Actualizando estado fila ${rowId}...`;

    try {
        // Col H = Estado (Índice 7 en base 0)
        const range = `Hoja 1!H${rowId}`; 
         console.log(`Actualizando estado en ${QUOTES_SHEET_ID} rango ${range} a ${newStatus}`);
        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: QUOTES_SHEET_ID,
            range: range,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[newStatus]] }
        });

        // Actualizar localmente si la API tuvo éxito
        const quoteIndex = quotesHistory.findIndex(q => q.id === quoteId);
        if (quoteIndex > -1) {
            quotesHistory[quoteIndex].status = newStatus;
            renderQuotesHistory(); // Redibujar tabla historial
            if (authStatus) authStatus.innerText = 'Estado actualizado.';
        } else {
             console.warn("Cotización no encontrada localmente tras actualizar estado:", quoteId);
             // Si no se encontró localmente (raro), forzar recarga completa
             if (authStatus) authStatus.innerText = 'Estado actualizado en Sheet, recargando historial local...';
             await loadQuotesHistoryFromSheet(); 
         }
    } catch(err) {
        console.error("Error actualizando estado:", err);
        const errorMsg = err.result?.error?.message || err.message || "Error desconocido.";
        showAlert(`Error al actualizar estado: ${errorMsg}\nRevisa la consola.`, "Error al Actualizar"); // REEMPLAZO
         if (authStatus) authStatus.innerText = 'Error al actualizar estado.';
          // Revertir visualmente el select si falla la API
          const selectElement = event?.target;
          if (selectElement) {
               const originalQuote = quotesHistory.find(q => q.id === quoteId);
               if (originalQuote) selectElement.value = originalQuote.status;
           }
     }
}

// Cargar datos de una cotización antigua en el formulario
function loadQuoteForEdit(quoteId) {
    const quote = quotesHistory.find(q => q.id === quoteId);
    if (!quote) {
         showAlert("Cotización no encontrada en el historial cargado."); // REEMPLAZO
         return;
     }

     // Chequeos de elementos DOM
     if (!quoteClientSelect || !quoteNotesInput || !quoteItemsBody) {
         console.error("loadQuoteForEdit: Faltan elementos del formulario.");
         return;
      }
     console.log("Cargando para editar:", quote);

    // Seleccionar cliente (si existe en la lista actual)
    if(quote.client && quote.client.id && quoteClientSelect.querySelector(`option[value="${quote.client.id}"]`)) {
         quoteClientSelect.value = quote.client.id;
    } else {
        console.warn("Cliente de la cotización no encontrado en la lista actual:", quote.client?.id);
         quoteClientSelect.value = ''; // Resetear si no está
     }

    // Cargar items (asegurando que sea un array)
    quoteItemsData = Array.isArray(quote.items) ? JSON.parse(JSON.stringify(quote.items)) : []; // Clonar items
    quoteNotesInput.value = quote.notes || ''; // Cargar notas

    // Disparar evento change en cliente para actualizar RUC y número preview
    const event = new Event('change');
    quoteClientSelect.dispatchEvent(event); // Esto llama a handleQuoteClientChange -> updateQuoteNumberPreview

    renderQuoteItems(); // Dibujar tabla de items

    // Cambiar a la pestaña del cotizador
    const tabButton = document.querySelector('.tab-button[onclick*="cotizador"]');
    if (tabButton) {
        changeTab({currentTarget: tabButton}, 'cotizador');
    }
    showAlert(`Cotización ${quote.number} cargada.\nPuedes modificarla y generar una NUEVA cotización.\nLos cambios NO sobrescribirán la original.`, "Cotización Cargada"); // REEMPLAZO
    window.scrollTo(0, 0); // Ir arriba
    currentQuoteContextId = quote.id;
    updateQuoteActionButtonsState();
}


// --- LÓGICA DE CONTRATOS ---
async function handleSaveContractClick() {
    const quoteId = contractQuoteSelect ? contractQuoteSelect.value : null;
    const text = contractText ? contractText.value : null;
    if (!quoteId || text === null || text.trim() === '') {
        showAlert('Selecciona una cotización aceptada y escribe el contenido del contrato.', 'Datos incompletos');
        return;
    }

    const quote = quotesHistory.find(q => q.id === quoteId);
    if (!quote) {
        showAlert('Cotización seleccionada no encontrada.');
        return;
    }

    if (quote.googleDocId) {
        showConfirm('Esta cotización ya tiene un contrato guardado en Google Drive. ¿Deseas abrirlo en una nueva pestaña?', () => {
            window.open(`https://docs.google.com/document/d/${quote.googleDocId}/edit`, '_blank');
        }, 'Contrato existente');
        return;
    }

    if (authStatus) authStatus.innerText = 'Creando documento de contrato...';
    saveContractBtn.disabled = true;

    try {
        const newDocId = await createContractDocumentForQuote(quote, text.trim());
        showAlert(`Contrato creado exitosamente en Google Drive.\nID: ${newDocId}\nSe ha guardado el enlace en el historial.`, 'Contrato creado');
        if (authStatus) authStatus.innerText = 'Contrato guardado.';
        currentQuoteContextId = quote.id;
        updateQuoteActionButtonsState();
    } catch (err) {
        console.error('Error guardando contrato:', err);
        const errorMsg = err.result?.error?.message || err.message || 'Error desconocido.';
        showAlert(`Error al guardar el contrato en Google Drive/Sheets: ${errorMsg}\nRevisa la consola.`, 'Error al guardar');
        if (authStatus) authStatus.innerText = 'Error al guardar contrato.';
    } finally {
        saveContractBtn.disabled = false;
    }
}

function buildContractBodyFromQuote(quote) {
    if (!quote) return '';

    const clientName = quote.client?.name || 'Cliente';
    const issueDate = quote.issueDate || 'N/A';
    const validityDate = quote.validityDate && quote.validityDate !== 'Indefinida'
        ? quote.validityDate
        : '';
    const items = Array.isArray(quote.items) ? quote.items : [];
    const companyForContract = quote.companySettingsForPrint || companySettings || DEFAULT_COMPANY_SETTINGS;

    const itemLines = items.map(item => {
        const name = item.name || 'Servicio';
        const quantity = parseFloat(item.quantity) || 1;
        const price = parseFloat(item.price) || 0;
        const total = (quantity * price).toFixed(2);
        return `• ${name} — Cantidad: ${quantity} — Total: $${total}`;
    });

    let body = `Este contrato se genera a partir de la cotización ${quote.number || 'sin número'} emitida el ${issueDate} y aceptada por ${clientName}.`;
    if (validityDate) {
        body += `\nLa vigencia acordada se extiende hasta ${validityDate}.`;
    }
    if (itemLines.length > 0) {
        body += `\n\nDetalle de servicios y/o productos contratados:\n${itemLines.join('\n')}`;
    }
    if (quote.notes) {
        body += `\n\nNotas acordadas:\n${quote.notes}`;
    }
    body += `\n\nLas partes se comprometen a cumplir con los términos y condiciones establecidos por ${companyForContract.name || 'la empresa proveedora'}, incluyendo plazos, entregas y formas de pago acordadas.`;
    body += `\n\nFirmas de conformidad:\n\n_______________________________    _______________________________\nRepresentante ${companyForContract.name || 'Proveedor'}    ${clientName}`;

    return body;
}

async function createContractDocumentForQuote(quote, bodyText = '') {
    if (!quote) {
        throw new Error('No se encontró la cotización a procesar.');
    }
    if (!quote.rowId) {
        throw new Error('No se pudo determinar la fila de Google Sheets asociada a la cotización.');
    }
    if (!gapi?.client?.drive || !gapi?.client?.docs) {
        throw new Error('Las APIs de Google Drive o Docs no están listas.');
    }
    if (!CONTRACT_TEMPLATE_ID || CONTRACT_TEMPLATE_ID === 'ID_DE_TU_PLANTILLA_DE_CONTRATO_EN_GOOGLE_DOCS') {
        throw new Error('Falta configurar el ID de la plantilla de contrato.');
    }

    const clientName = quote.client?.name || 'Cliente';
    const docTitle = `Contrato - ${quote.number || 'Cotización'} - ${clientName}`;

    const copyPayload = {
        fileId: CONTRACT_TEMPLATE_ID,
        resource: { name: docTitle }
    };
    if (CONTRACTS_DRIVE_FOLDER_ID) {
        copyPayload.resource.parents = [CONTRACTS_DRIVE_FOLDER_ID];
    }

    const copyResponse = await gapi.client.drive.files.copy({
        ...copyPayload,
        supportsAllDrives: true,
        fields: 'id, parents'
    });
    const newDocId = copyResponse.result.id;
    const assignedParents = copyResponse.result.parents || [];
    console.log('Plantilla copiada, nuevo ID de contrato:', newDocId, 'con padres:', assignedParents);

    const basicRequests = [
        { replaceAllText: { containsText: { text: '{{CLIENTE_NOMBRE}}', matchCase: false }, replaceText: quote.client?.name || '' } },
        { replaceAllText: { containsText: { text: '{{CLIENTE_RUC}}', matchCase: false }, replaceText: quote.client?.ruc || '' } },
        { replaceAllText: { containsText: { text: '{{COTIZACION_NUMERO}}', matchCase: false }, replaceText: quote.number || '' } },
        { replaceAllText: { containsText: { text: '{{FECHA_EMISION}}', matchCase: false }, replaceText: quote.issueDate || '' } },
        { replaceAllText: { containsText: { text: '{{TOTAL}}', matchCase: false }, replaceText: `$${(quote.total || 0).toFixed(2)}` } }
    ];

    await gapi.client.docs.documents.batchUpdate({
        documentId: newDocId,
        resource: { requests: basicRequests }
    });

    if (typeof bodyText === 'string') {
        const contentRequest = [{
            replaceAllText: {
                containsText: { text: '{{CONTENIDO_CONTRATO}}', matchCase: false },
                replaceText: bodyText
            }
        }];
        try {
            await gapi.client.docs.documents.batchUpdate({
                documentId: newDocId,
                resource: { requests: contentRequest }
            });
        } catch (error) {
            console.warn('No se pudo reemplazar {{CONTENIDO_CONTRATO}} en la plantilla:', error);
        }
    }

    const range = `Hoja 1!J${quote.rowId}`;
    await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: QUOTES_SHEET_ID,
        range,
        valueInputOption: 'USER_ENTERED',
        resource: { values: [[newDocId]] }
    });

    quote.googleDocId = newDocId;
    currentQuoteContextId = quote.id;
    renderQuotesHistory();

    return newDocId;
}

async function generateContractForCurrentQuote() {
    if (!generateContractBtn) {
        return;
    }

    if (!currentQuoteContextId) {
        showAlert('Genera o carga una cotización aceptada desde el historial antes de crear el contrato.', 'Cotización no seleccionada');
        return;
    }

    const quote = quotesHistory.find(q => q.id === currentQuoteContextId);
    if (!quote) {
        showAlert('La cotización seleccionada ya no está disponible. Actualiza el historial e inténtalo nuevamente.', 'Cotización no encontrada');
        currentQuoteContextId = null;
        updateQuoteActionButtonsState();
        return;
    }

    if (quote.status !== 'Aceptada') {
        showAlert('Marca la cotización como "Aceptada" en el historial para habilitar la generación del contrato.', 'Cotización pendiente');
        updateQuoteActionButtonsState();
        return;
    }

    if (quote.googleDocId) {
        showConfirm('Esta cotización ya tiene un contrato guardado en Google Drive. ¿Deseas abrirlo en una nueva pestaña?', () => {
            window.open(`https://docs.google.com/document/d/${quote.googleDocId}/edit`, '_blank');
        }, 'Contrato existente');
        return;
    }

    if (!quote.rowId) {
        showAlert('No se encontró la fila de Google Sheets para esta cotización. Refresca el historial antes de continuar.', 'Referencia incompleta');
        return;
    }

    setButtonLoadingState(generateContractBtn, true, 'Generando contrato...');
    if (authStatus) authStatus.innerText = 'Generando contrato en Google Drive...';

    try {
        const bodyText = buildContractBodyFromQuote(quote);
        const newDocId = await createContractDocumentForQuote(quote, bodyText);
        showAlert(`Contrato creado exitosamente en Google Drive.\nID: ${newDocId}\nPuedes editarlo desde tu cuenta.`, 'Contrato creado');
        if (authStatus) authStatus.innerText = 'Contrato guardado en Google Drive.';
        window.open(`https://docs.google.com/document/d/${newDocId}/edit`, '_blank');
    } catch (error) {
        console.error('Error al generar el contrato desde el resumen:', error);
        const errorMsg = error.result?.error?.message || error.message || 'Error desconocido.';
        showAlert(`No se pudo generar el contrato: ${errorMsg}\nVerifica los permisos o la plantilla e inténtalo nuevamente.`, 'Error al generar contrato');
        if (authStatus) authStatus.innerText = 'Error al generar contrato.';
    } finally {
        setButtonLoadingState(generateContractBtn, false);
        updateQuoteActionButtonsState();
    }
}

function updateQuoteActionButtonsState() {
    if (!generateContractBtn) {
        return;
    }

    if (generateContractBtn.classList.contains('is-loading')) {
        return;
    }

    let disabled = true;
    let tooltip = 'Genera o carga una cotización y márcala como aceptada para habilitar esta acción.';

    if (currentQuoteContextId) {
        const quote = quotesHistory.find(q => q.id === currentQuoteContextId);
        if (quote) {
            if (quote.status === 'Aceptada') {
                if (!quote.googleDocId) {
                    disabled = false;
                    tooltip = 'Genera el contrato en Google Docs para la cotización aceptada seleccionada.';
                } else {
                    tooltip = 'Esta cotización ya tiene un contrato guardado en Google Drive.';
                }
            } else {
                tooltip = 'Marca la cotización como "Aceptada" en el historial para generar el contrato.';
            }
        } else {
            tooltip = 'Selecciona nuevamente la cotización desde el historial.';
        }
    }

    generateContractBtn.disabled = disabled;
    generateContractBtn.title = tooltip;
}


// --- LÓGICA DE GENERACIÓN Y GUARDADO DE COTIZACIÓN ---
function setButtonLoadingState(buttonOrId, isLoading, labelWhenLoading = 'Procesando...') {
    const button = typeof buttonOrId === 'string' ? document.getElementById(buttonOrId) : buttonOrId;
    if (!button) return;

    if (isLoading) {
        if (!button.dataset.originalHtml) {
            button.dataset.originalHtml = button.innerHTML;
        }
        button.disabled = true;
        button.classList.add('is-loading');
        const loadingLabel = labelWhenLoading || 'Procesando...';
        button.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>${loadingLabel}`;
    } else {
        button.classList.remove('is-loading');
        if (button.dataset.originalHtml) {
            button.innerHTML = button.dataset.originalHtml;
            delete button.dataset.originalHtml;
        }
        button.disabled = false;
    }
}

async function generatePrintableQuote() {
    let configOk = true;
    if (!QUOTES_SHEET_ID || QUOTES_SHEET_ID === 'ID_DE_TU_HOJA_DE_COTIZACIONES') {
        showAlert("Falta configurar el ID de la Hoja de Historial (QUOTES_SHEET_ID).", "Error de Configuración");
        configOk = false;
    }
    if (!configOk) return;

    if (!quoteClientSelect || !quoteItemsData || !quoteNotesInput || !quoteDate) {
        showAlert("Error: Faltan elementos del formulario de cotización.", "Error de Formulario");
        return;
    }
    const client = clients.find(c => c.id === quoteClientSelect.value);
    if (!client) { showAlert('Por favor, seleccione un cliente.'); return; }
    if (quoteItemsData.length === 0) { showAlert('Por favor, añada al menos un servicio a la cotización.'); return; }

    const now = new Date();
    const dateStr = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
    const timeStr = `${now.getHours().toString().padStart(2, '0')}${now.getMinutes().toString().padStart(2, '0')}`;
    const sequential = (quotesHistory.length + 1).toString().padStart(4, '0');
    const finalQuoteNumber = `${client.code}-${dateStr}${timeStr}-${sequential}`;

    const quoteDateValue = quoteDate.value;
    const validityDate = quoteDateValue ? new Date(quoteDateValue + 'T00:00:00Z').toLocaleDateString('es-EC') : 'Indefinida';
    const issueDate = now.toLocaleDateString('es-EC');

    const subtotal = quoteItemsData.reduce((acc, item) => {
        const itemPrice = parseFloat(item.price) || 0;
        const itemQuantity = parseFloat(item.quantity) || 1;
        return acc + (itemPrice * itemQuantity);
    }, 0);
    const iva = parseFloat((subtotal * IVA_RATE).toFixed(2));
    const total = parseFloat((subtotal + iva).toFixed(2));

    const quoteRecord = {
        id: `quote_${now.getTime()}`,
        number: finalQuoteNumber,
        issueDate: issueDate,
        validityDate: validityDate,
        client: JSON.parse(JSON.stringify(client)),
        items: JSON.parse(JSON.stringify(quoteItemsData)),
        subtotal: subtotal,
        iva: iva,
        total: total,
        notes: quoteNotesInput.value,
        status: 'Pendiente',
        googleDocId: null,
        googlePdfId: null
    };
    quoteRecord.companySettingsForPrint = { ...companySettings };

    setButtonLoadingState('generate-quote-pdf-btn', true, 'Generando PDF...');
    if (authStatus) authStatus.innerText = 'Generando PDF de la cotización...';

    try {
        const pdfLibraries = await ensurePdfLibrariesAvailable();
        setButtonLoadingState('generate-quote-pdf-btn', true, 'Generando PDF...');
        if (authStatus) authStatus.innerText = 'Generando PDF de la cotización...';
        const pdfBlob = await generateQuotePdfBlob(quoteRecord, pdfLibraries);

        setButtonLoadingState('generate-quote-pdf-btn', true, 'Guardando en Google Drive...');
        if (authStatus) authStatus.innerText = 'Guardando PDF en Google Drive...';
        const pdfFileName = `Cotizacion-${quoteRecord.number}.pdf`;
        const pdfIdFromDrive = await uploadPdfBlobToDrive(pdfBlob, pdfFileName, QUOTES_DRIVE_FOLDER_ID);
        quoteRecord.googlePdfId = pdfIdFromDrive;

        let successMessage = 'PDF de cotización generado y guardado en Google Drive.';
        if (pdfIdFromDrive) successMessage += `\nID de PDF: ${pdfIdFromDrive}`;
        showAlert(`${successMessage}\nSe guardará el enlace en el historial.`, 'Cotización Guardada');

        const sheetRowData = [
            quoteRecord.id,
            quoteRecord.number,
            quoteRecord.issueDate,
            quoteRecord.validityDate,
            JSON.stringify(quoteRecord.client),
            JSON.stringify(quoteRecord.items),
            quoteRecord.total,
            quoteRecord.status,
            quoteRecord.notes,
            quoteRecord.googleDocId || '',
            quoteRecord.googlePdfId || ''
        ];

        setButtonLoadingState('generate-quote-pdf-btn', true, 'Registrando en Google Sheets...');
        if (authStatus) authStatus.innerText = 'Guardando registro en Google Sheets...';
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: QUOTES_SHEET_ID,
            range: 'Hoja 1!A:K',
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: [sheetRowData] }
        });

        const newRowIndex = quotesHistory.length + 2;
        quoteRecord.rowId = newRowIndex;
        quotesHistory.push(quoteRecord);
        renderQuotesHistory();
        quoteCounter = quotesHistory.length + 1;

        if (pdfIdFromDrive) {
            window.open(`https://drive.google.com/file/d/${pdfIdFromDrive}/view`, '_blank');
        }
        if (authStatus) authStatus.innerText = 'Cotización generada y guardada.';

        updateQuoteNumberPreview();
        quoteItemsData = [];
        if (quoteNotesInput) quoteNotesInput.value = '';
        if (serviceSelect) serviceSelect.value = '';
        syncServicePriceOverrideField();
        renderQuoteItems();

        currentQuoteContextId = quoteRecord.id;
        updateQuoteActionButtonsState();

    } catch (err) {
        console.error('Error generando/guardando cotización:', err);
        const errorMsg = err.result?.error?.message || err.message || 'Error desconocido.';
        showAlert(`Error al generar/guardar la cotización: ${errorMsg}\nRevisa la consola.`, 'Error al Guardar');
        if (authStatus) authStatus.innerText = 'Error al generar/guardar.';
    } finally {
        setButtonLoadingState('generate-quote-pdf-btn', false);
        updateQuoteActionButtonsState();
    }
}

// --- LÓGICA DE PESTAÑAS ---
function changeTab(event, tabName) {
    const panels = document.querySelectorAll('.tab-content');
    panels.forEach(panel => panel?.classList?.remove('active'));

    const buttons = document.querySelectorAll('.tab-button');
    buttons.forEach(button => button?.classList?.remove('active'));

    const tabElement = document.getElementById(tabName);
    if (tabElement?.classList) {
        tabElement.classList.add('active');
    }

    let triggerButton = event?.currentTarget;
    if (!triggerButton || !triggerButton.classList) {
        triggerButton = document.querySelector(`.tab-button[data-tab-target="${tabName}"]`) ||
            document.querySelector(`.tab-button[onclick*="'${tabName}'"]`);
    }
    if (triggerButton?.classList) {
        triggerButton.classList.add('active');
    }

    if (tabName === 'contratos') renderContractsPage();
    if (tabName === 'portal') renderClientPortalPage();
}

// --- LÓGICA DEL COTIZADOR ---
function updateQuoteNumberPreview() {
    // Chequeo de existencia de elementos
    if (!quoteClientSelect || !quoteNumberInput) return; 

    const selectedClientId = quoteClientSelect.value;
    const client = clients.find(c => c.id === selectedClientId);
    const clientCode = client ? client.code : 'CLIENTE'; // client.code es el RUC/ID
    // Usa el contador actual del historial + 1
    const sequential = (quotesHistory.length + 1).toString().padStart(4, '0'); 
    const now = new Date();
    // Formato YYYYMMDD
    const dateStr = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
    quoteNumberInput.value = `${clientCode}-${dateStr}-${sequential}`;
}

function updateClientDropdowns() {
    const selects = [quoteClientSelect, contractClientSelect, portalClientSelect];
    selects.forEach(select => {
         if (!select) return; // Chequear si el select existe en el DOM
        const currentValue = select.value; // Guardar valor actual si existe
        select.innerHTML = '<option value="">-- Seleccione un cliente --</option>'; // Opción por defecto más clara
        // Ordenar clientes alfabéticamente por nombre
        [...clients].sort((a, b) => (a.name || '').localeCompare(b.name || '')).forEach(client => { // Añadir fallback para name
            const option = document.createElement('option');
            option.value = client.id; // client.id es el RUC/ID
            option.textContent = `${client.name || 'Sin Nombre'} (${client.ruc || 'Sin RUC'})`; // Añadir fallback
            select.appendChild(option);
        });
         // Intentar restaurar valor si es posible
         if (currentValue && select.querySelector(`option[value="${currentValue}"]`)) {
             select.value = currentValue;
         }
    });
}

// Manejador para el cambio en el select de cliente del cotizador
function handleQuoteClientChange() {
     if (!quoteClientSelect || !quoteRUCInput) return; // Chequeo
    const selectedClientId = quoteClientSelect.value;
    const client = clients.find(c => c.id === selectedClientId);
    quoteRUCInput.value = client ? client.ruc : ''; // Poner RUC en el campo RUC
    updateQuoteNumberPreview(); // Actualizar número de cotización
}

// Actualizar el dropdown de servicios
function updateServiceDropdown() {
     if (!serviceSelect) return; // Chequeo
     const currentValue = serviceSelect.value;
    serviceSelect.innerHTML = '<option value="">-- Seleccione un servicio --</option>';
    // Ordenar servicios alfabéticamente
    [...services].sort((a, b) => (a.name || '').localeCompare(b.name || '')).forEach(service => { // Añadir fallback
        const option = document.createElement('option');
        option.value = service.id; // service.id es de Col A
        option.textContent = `${service.name || 'Sin Nombre'} - $${(service.price || 0).toFixed(2)}`; // Añadir fallback
        serviceSelect.appendChild(option);
    });
     if (currentValue && serviceSelect.querySelector(`option[value="${currentValue}"]`)) {
         serviceSelect.value = currentValue;
     }
    syncServicePriceOverrideField(true);
}

function handleServiceSelectChange() {
    syncServicePriceOverrideField(true);
}

function syncServicePriceOverrideField(forceDefault = false) {
    if (!servicePriceOverrideInput) return;
    const hasSelection = !!serviceSelect && !!serviceSelect.value;

    if (!hasSelection) {
        servicePriceOverrideInput.value = '';
        servicePriceOverrideInput.placeholder = '0.00';
        servicePriceOverrideInput.disabled = true;
        return;
    }

    const selectedService = services.find(s => s.id === serviceSelect.value);
    const servicePrice = parseFloat(selectedService?.price);

    servicePriceOverrideInput.disabled = false;

    if (
        forceDefault ||
        servicePriceOverrideInput.value === '' ||
        servicePriceOverrideInput.disabled
    ) {
        if (Number.isFinite(servicePrice)) {
            servicePriceOverrideInput.value = servicePrice.toFixed(2);
        } else {
            servicePriceOverrideInput.value = '0.00';
        }
    }
}

// Añadir servicio a la tabla temporal de la cotización
function addServiceToQuote() {
     if (!serviceSelect) return; // Chequeo
    const serviceId = serviceSelect.value;
    if (!serviceId) { showAlert('Por favor, seleccione un servicio.'); return; } // REEMPLAZO
    const service = services.find(s => s.id === serviceId);
    if (!service) { showAlert('Servicio seleccionado no encontrado en la lista cargada.'); return; } // REEMPLAZO

    let overridePrice = null;
    if (servicePriceOverrideInput && !servicePriceOverrideInput.disabled) {
        const rawPrice = (servicePriceOverrideInput.value || '').trim();
        if (rawPrice !== '') {
            const parsedOverride = parseFloat(rawPrice);
            if (!Number.isFinite(parsedOverride)) {
                showAlert('Ingrese un precio válido para el servicio seleccionado.');
                servicePriceOverrideInput.focus();
                return;
            }
            if (parsedOverride < 0) {
                showAlert('El precio debe ser un número mayor o igual a 0.');
                servicePriceOverrideInput.focus();
                return;
            }
            overridePrice = parseFloat(parsedOverride.toFixed(2));
        }
    }

    const parsedServicePrice = parseFloat(service.price);
    let defaultPrice = 0;
    if (Number.isFinite(parsedServicePrice)) {
        defaultPrice = parseFloat(parsedServicePrice.toFixed(2));
    }
    const finalPrice = overridePrice !== null ? overridePrice : defaultPrice;

    if (servicePriceOverrideInput && !servicePriceOverrideInput.disabled) {
        servicePriceOverrideInput.value = finalPrice.toFixed(2);
    }

    // Buscar si ya existe en la cotización actual
    const existingItemIndex = quoteItemsData.findIndex(item => item.id === serviceId);
    if (existingItemIndex > -1) {
        // Si existe, incrementar cantidad
        quoteItemsData[existingItemIndex].quantity++;
        if (overridePrice !== null) {
            quoteItemsData[existingItemIndex].price = finalPrice;
        }
    } else {
        // Si no existe, añadirlo con cantidad 1
        quoteItemsData.push({
            id: service.id,
            name: service.name,
            description: service.description, // Aunque no se muestra, guardarlo
            price: finalPrice,
            quantity: 1
        });
    }
    renderQuoteItems(); // Redibujar la tabla de items
}

// Dibujar la tabla de items de la cotización actual
function renderQuoteItems() {
     if (!quoteItemsBody) return; // Chequeo
    quoteItemsBody.innerHTML = ''; // Limpiar tabla
    if (quoteItemsData.length === 0) {
        quoteItemsBody.innerHTML = `<tr><td colspan="5" class="text-center py-4 text-muted">Añada servicios a la cotización.</td></tr>`;
    } else {
        quoteItemsData.forEach((item, index) => {
            const row = document.createElement('tr');
            const itemPrice = parseFloat(item.price) || 0; // Fallback numérico
            const itemQuantity = parseFloat(item.quantity) || 1; // Fallback numérico
            const totalPrice = (itemPrice * itemQuantity).toFixed(2);
            // Input de cantidad ahora llama a updateQuantity con 'index' y 'this'
            row.innerHTML = `
                <td class="align-middle">
                    <div class="fw-semibold text-dark">${item.name || 'N/A'}</div>
                </td>
                <td class="align-middle" style="width: 110px;">
                    <input type="number" min="1" value="${itemQuantity}" onchange="updateQuantity(${index}, this)" class="form-control form-control-sm text-center">
                </td>
                <td class="align-middle" style="width: 130px;">
                    <input type="number" min="0" step="0.01" value="${itemPrice.toFixed(2)}" onchange="updatePrice(${index}, this)" class="form-control form-control-sm text-end">
                </td>
                <td class="align-middle text-end fw-semibold">$${totalPrice}</td>
                <td class="align-middle text-end">
                    <button onclick="removeQuoteItem(${index})" class="btn btn-sm btn-outline-danger" title="Quitar Item"><i class="fas fa-trash"></i></button>
                </td>`;
            quoteItemsBody.appendChild(row);
        });
    }
    calculateTotals(); // Recalcular totales cada vez que se redibuja
}

// Actualizar cantidad de un item
function updateQuantity(index, inputElement) {
     if (!inputElement) return; // Chequeo
     // Validar que el índice es válido
     if (index < 0 || index >= quoteItemsData.length) {
         console.error("Índice inválido para actualizar cantidad:", index);
         inputElement.value = 1; // Resetear input a 1 si el índice es malo
         return;
     }
    const newQuantity = parseInt(inputElement.value);
    // Validar que la cantidad sea un número positivo
    if (isNaN(newQuantity) || newQuantity < 1) {
         showAlert("La cantidad debe ser un número mayor o igual a 1."); // REEMPLAZO
         inputElement.value = quoteItemsData[index].quantity; // Restaurar valor anterior válido
         return;
    }
    quoteItemsData[index].quantity = newQuantity;
    renderQuoteItems(); // Redibujar para actualizar total del item y resumen
}

function updatePrice(index, inputElement) {
    if (!inputElement) return;
    if (index < 0 || index >= quoteItemsData.length) {
        console.error('Índice inválido para actualizar precio:', index);
        inputElement.value = quoteItemsData[index]?.price?.toFixed?.(2) || '0.00';
        return;
    }
    const newPrice = parseFloat(inputElement.value);
    if (isNaN(newPrice) || newPrice < 0) {
        showAlert('El precio debe ser un número mayor o igual a 0.');
        inputElement.value = (quoteItemsData[index].price || 0).toFixed(2);
        return;
    }
    quoteItemsData[index].price = parseFloat(newPrice.toFixed(2));
    inputElement.value = quoteItemsData[index].price.toFixed(2);
    renderQuoteItems();
}

// Quitar un item de la cotización
function removeQuoteItem(index) {
     if (index < 0 || index >= quoteItemsData.length) {
         console.error("Índice inválido para quitar item:", index);
         return;
     }
    // Quitar el item del array
    const removedItem = quoteItemsData.splice(index, 1); 
    console.log("Item quitado:", removedItem[0]?.name);
    renderQuoteItems(); // Redibujar la tabla
}

// Calcular y mostrar Subtotal, IVA y Total
function calculateTotals() {
    // Usar fallbacks en reduce
    const subtotal = quoteItemsData.reduce((acc, item) => {
        const itemPrice = parseFloat(item.price) || 0;
        const itemQuantity = parseFloat(item.quantity) || 1;
        return acc + (itemPrice * itemQuantity);
    }, 0);
    const iva = parseFloat((subtotal * IVA_RATE).toFixed(2));
    const total = parseFloat((subtotal + iva).toFixed(2));

    // Actualizar elementos del DOM (con chequeos)
    const subtotalEl = document.getElementById('subtotal');
    const ivaEl = document.getElementById('iva');
    const totalEl = document.getElementById('total');
    if(subtotalEl) subtotalEl.textContent = `$${subtotal.toFixed(2)}`;
    if(ivaEl) ivaEl.textContent = `$${iva.toFixed(2)}`;
    if(totalEl) totalEl.textContent = `$${total.toFixed(2)}`;
}


// --- LÓGICA DE CONFIGURACIÓN (local y nube) ---
function loadSettings() {
    if (!companyNameInput || !companyAddressInput || !companyContactInput || !companyWebsiteInput || !companyWhatsappInput || !logoPreview || !primaryColorInput || !accentColorInput) {
        console.warn("Faltan elementos del formulario de configuración.");
        return;
    }

    companySettings = { ...DEFAULT_COMPANY_SETTINGS, ...companySettings };

    companyNameInput.value = companySettings.name || DEFAULT_COMPANY_SETTINGS.name;
    companyAddressInput.value = companySettings.address || DEFAULT_COMPANY_SETTINGS.address;
    companyContactInput.value = companySettings.contact || DEFAULT_COMPANY_SETTINGS.contact;
    companyWebsiteInput.value = companySettings.website || DEFAULT_COMPANY_SETTINGS.website;
    companyWhatsappInput.value = companySettings.whatsapp || DEFAULT_COMPANY_SETTINGS.whatsapp;
    logoPreview.src = companySettings.logo || DEFAULT_COMPANY_SETTINGS.logo;
    primaryColorInput.value = companySettings.primaryColor || DEFAULT_COMPANY_SETTINGS.primaryColor;
    accentColorInput.value = companySettings.accentColor || DEFAULT_COMPANY_SETTINGS.accentColor;

    if (autoSyncCloudToggle) autoSyncCloudToggle.checked = autoSyncCloud;

    applyAccentColor();
    updateRemoteConfigMeta();
}

function applyAccentColor() {
    const accentColor = companySettings.accentColor || DEFAULT_COMPANY_SETTINGS.accentColor;
    const totalLabelEl = document.getElementById('total-label');
    const totalEl = document.getElementById('total');
    if (totalLabelEl) totalLabelEl.style.color = accentColor;
    if (totalEl) totalEl.style.color = accentColor;
}

function persistCompanySettingsLocally() {
    try {
        localStorage.setItem('companySettings', JSON.stringify(companySettings));
        return true;
    } catch (error) {
        console.error('Error guardando configuración en localStorage:', error);
        return false;
    }
}

function updateCompanySettingsFromForm() {
    if (!companyNameInput || !companyAddressInput || !companyContactInput || !companyWebsiteInput || !companyWhatsappInput || !primaryColorInput || !accentColorInput) {
        showAlert('Error: Faltan elementos del formulario de configuración para guardar.', 'Error de Formulario');
        return false;
    }

    companySettings = {
        ...DEFAULT_COMPANY_SETTINGS,
        ...companySettings,
        name: companyNameInput.value.trim(),
        address: companyAddressInput.value.trim(),
        contact: companyContactInput.value.trim(),
        website: companyWebsiteInput.value.trim(),
        whatsapp: companyWhatsappInput.value.trim(),
        primaryColor: primaryColorInput.value,
        accentColor: accentColorInput.value
    };

    if (!persistCompanySettingsLocally()) {
        showAlert('Hubo un error al guardar la configuración local.');
        return false;
    }

    loadSettings();
    return true;
}

function handleLogoUploadChange(e) {
    if (!logoPreview) return;
    if (!e.target.files || e.target.files.length === 0) return;

    const file = e.target.files[0];
    if (!file.type.startsWith('image/')) {
        showAlert('Por favor, selecciona un archivo de imagen válido.');
        e.target.value = '';
        return;
    }
    if (file.size > 2 * 1024 * 1024) {
        showAlert('El archivo es muy grande (máx 2MB).');
        e.target.value = '';
        return;
    }

    const reader = new FileReader();
    reader.onload = async (event) => {
        const logoBase64 = event.target.result;
        logoPreview.src = logoBase64;
        companySettings = { ...DEFAULT_COMPANY_SETTINGS, ...companySettings, logo: logoBase64 };

        if (!persistCompanySettingsLocally()) {
            showAlert('Hubo un error al guardar el logo. El almacenamiento local podría estar lleno.');
            return;
        }

        console.log('Logo guardado en localStorage.');
        updateRemoteConfigMeta();

        if (autoSyncCloud) {
            const synced = await saveCompanySettingsToCloud({ silent: true });
            if (!synced) {
                console.warn('No se pudo sincronizar el logo con la nube de inmediato.');
            }
        } else {
            refreshRemoteConfigUiState();
        }
    };
    reader.onerror = (error) => {
        console.error('Error leyendo archivo de logo:', error);
        showAlert('Error al leer el archivo de logo seleccionado.');
        e.target.value = '';
    };
    reader.readAsDataURL(file);
}

async function handleConfigFormSubmit(e) {
    e.preventDefault();
    const updated = updateCompanySettingsFromForm();
    if (!updated) return;

    showAlert('Configuración guardada localmente.');

    if (autoSyncCloud) {
        const synced = await saveCompanySettingsToCloud({ silent: true });
        if (!synced) {
            showAlert('Los cambios se guardaron en este dispositivo, pero no se pudieron sincronizar con la nube. Reintenta más tarde.', 'Advertencia');
        }
    } else {
        refreshRemoteConfigUiState();
    }
}

async function handleManualCloudSave() {
    const updated = updateCompanySettingsFromForm();
    if (!updated) return;
    await saveCompanySettingsToCloud({ silent: false });
}

async function handleManualCloudLoad() {
    await loadCompanySettingsFromCloud({ silent: false });
}

function handleAutoSyncToggleChange(event) {
    autoSyncCloud = !!event.target.checked;
    try {
        localStorage.setItem('companySettingsAutoSyncCloud', autoSyncCloud ? 'true' : 'false');
    } catch (error) {
        console.warn('No se pudo guardar la preferencia de sincronización automática:', error);
    }
    updateRemoteConfigMeta();
    if (autoSyncCloud && companySettingsRowIndex) {
        refreshRemoteConfigUiState();
    }
}

function initializeRemoteConfigPanel() {
    refreshRemoteConfigUiState();
    updateRemoteConfigMeta();
}

function getCurrentUserEmail() {
    const email = signedInUser?.email || '';
    return typeof email === 'string' ? email.trim() : '';
}

function isConfigSheetConfigured() {
    return Boolean(COMPANY_CONFIG_SHEET_ID && COMPANY_CONFIG_SHEET_ID !== 'ID_DE_TU_HOJA_DE_CONFIGURACION');
}

function updateRemoteConfigButtonsState() {
    const disable = remoteConfigBusy || !isConfigSheetConfigured() || !getCurrentUserEmail();
    if (saveCloudConfigBtn) saveCloudConfigBtn.disabled = disable;
    if (loadCloudConfigBtn) loadCloudConfigBtn.disabled = disable;
}

function updateRemoteConfigStatus(message, state = 'info', options = {}) {
    if (!remoteConfigStatus) return;
    const { isLoading = false } = options;
    const classMap = {
        success: 'alert-success',
        info: 'alert-info',
        warning: 'alert-warning',
        danger: 'alert-danger',
        error: 'alert-danger'
    };
    const alertClass = classMap[state] || 'alert-secondary';
    remoteConfigStatus.className = `alert ${alertClass} mb-3`;
    if (isLoading) {
        remoteConfigStatus.innerHTML = `<div class="d-flex align-items-center gap-2"><div class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></div><span>${escapeHtml(message)}</span></div>`;
    } else {
        remoteConfigStatus.innerHTML = `<span>${escapeHtml(message)}</span>`;
    }
}

function updateRemoteConfigMeta() {
    if (!remoteConfigMeta) return;
    const metaParts = [];
    if (companySettingsRowIndex) metaParts.push(`Fila remota ${companySettingsRowIndex}`);
    if (remoteConfigLastLoadedAt instanceof Date && !Number.isNaN(remoteConfigLastLoadedAt.getTime())) {
        metaParts.push(`Última carga ${remoteConfigLastLoadedAt.toLocaleString()}`);
    }
    if (remoteConfigLastSavedAt instanceof Date && !Number.isNaN(remoteConfigLastSavedAt.getTime())) {
        metaParts.push(`Última sincronización ${remoteConfigLastSavedAt.toLocaleString()}`);
    }
    metaParts.push(autoSyncCloud ? 'Sincronización automática activa' : 'Sincronización automática desactivada');
    remoteConfigMeta.textContent = metaParts.join(' · ');
}

function setRemoteConfigBusy(isBusy, message) {
    remoteConfigBusy = isBusy;
    if (isBusy && message) {
        updateRemoteConfigStatus(message, 'info', { isLoading: true });
    }
    updateRemoteConfigButtonsState();
}

function refreshRemoteConfigUiState() {
    updateRemoteConfigButtonsState();

    if (!remoteConfigStatus) return;

    if (!isConfigSheetConfigured()) {
        updateRemoteConfigStatus('Define COMPANY_CONFIG_SHEET_ID en js/app.js para habilitar la sincronización en la nube.', 'warning');
        updateRemoteConfigMeta();
        return;
    }

    if (!getCurrentUserEmail()) {
        updateRemoteConfigStatus('Inicia sesión para sincronizar tu configuración con la nube.', 'info');
        updateRemoteConfigMeta();
        return;
    }

    if (remoteConfigBusy) {
        updateRemoteConfigMeta();
        return;
    }

    if (companySettingsRowIndex) {
        updateRemoteConfigStatus('Configuración conectada a la nube. Guarda para sincronizar cambios.', 'success');
    } else {
        updateRemoteConfigStatus('Aún no hay una copia remota para tu cuenta. Guarda para crearla.', 'info');
    }
    updateRemoteConfigMeta();
}

function extractRowIndexFromRange(range) {
    if (!range) return null;
    const match = range.match(/([A-Z]+)(\d+)/);
    if (match && match[2]) return Number(match[2]);
    const digits = range.match(/\d+/g);
    if (digits && digits.length > 0) {
        return Number(digits[digits.length - 1]);
    }
    return null;
}

function readCompanySettingsFromLocalStorage() {
    try {
        const raw = localStorage.getItem('companySettings');
        if (!raw) return { ...DEFAULT_COMPANY_SETTINGS };
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') {
            return { ...DEFAULT_COMPANY_SETTINGS, ...parsed };
        }
    } catch (error) {
        console.warn('No se pudo leer la configuración local:', error);
    }
    return { ...DEFAULT_COMPANY_SETTINGS };
}

async function loadCompanySettingsFromCloud({ silent = false } = {}) {
    if (!isConfigSheetConfigured()) {
        if (!silent) {
            showAlert('Configura COMPANY_CONFIG_SHEET_ID en js/app.js para habilitar la sincronización en la nube.', 'Configuración incompleta');
        }
        refreshRemoteConfigUiState();
        return false;
    }

    if (!gapi?.client?.sheets) {
        updateRemoteConfigStatus('La API de Google Sheets no está lista. Intenta recargar.', 'danger');
        if (!silent) {
            showAlert('La API de Google Sheets no está lista. Recarga la página e inténtalo de nuevo.', 'Error de API');
        }
        return false;
    }

    const email = getCurrentUserEmail();
    if (!email) {
        updateRemoteConfigStatus('No se pudo determinar el correo de la sesión para sincronizar la configuración.', 'warning');
        if (!silent) {
            showAlert('Inicia sesión con una cuenta que entregue el correo electrónico para sincronizar la configuración.', 'Correo no disponible');
        }
        return false;
    }

    setRemoteConfigBusy(true, 'Buscando configuración en la nube...');
    let success = false;

    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: COMPANY_CONFIG_SHEET_ID,
            range: COMPANY_CONFIG_SHEET_RANGE,
        });

        const values = response.result.values || [];
        const lowerEmail = email.toLowerCase();
        let foundRow = null;
        let foundIndex = null;

        for (let i = 0; i < values.length; i += 1) {
            const rowEmail = (values[i][0] || '').toLowerCase();
            if (rowEmail === lowerEmail) {
                foundRow = values[i];
                foundIndex = i + 2; // Datos comienzan en la fila 2
                break;
            }
        }

        if (!foundRow) {
            companySettingsRowIndex = null;
            remoteConfigLastLoadedAt = new Date();
            const statusType = silent ? 'info' : 'warning';
            updateRemoteConfigStatus('No se encontró una configuración guardada en la nube para tu cuenta. Guarda para crearla.', statusType);
            updateRemoteConfigMeta();
            if (!silent) {
                showAlert('No se encontró una configuración guardada en la nube para tu cuenta. Puedes guardarla ahora para crearla.', 'Sin configuración remota');
            }
            return false;
        }

        let remoteSettings = null;
        if (foundRow[1]) {
            try {
                const parsedSettings = JSON.parse(foundRow[1]);
                if (parsedSettings && typeof parsedSettings === 'object') {
                    remoteSettings = parsedSettings;
                }
            } catch (error) {
                console.error('No se pudo interpretar la configuración remota:', error);
                if (!silent) {
                    showAlert('La configuración remota existe pero no se pudo interpretar. Se conservarán los valores locales.', 'Error de formato');
                }
            }
        }

        if (remoteSettings) {
            companySettings = { ...DEFAULT_COMPANY_SETTINGS, ...remoteSettings };
        }
        persistCompanySettingsLocally();
        loadSettings();
        companySettingsRowIndex = foundIndex;
        remoteConfigLastLoadedAt = new Date();

        if (foundRow[2]) {
            const parsedDate = new Date(foundRow[2]);
            if (!Number.isNaN(parsedDate.getTime())) {
                remoteConfigLastSavedAt = parsedDate;
            }
        }

        updateRemoteConfigStatus('Configuración recuperada desde la nube.', 'success');
        updateRemoteConfigMeta();
        if (!silent) {
            showAlert('La configuración de la empresa se actualizó con la copia en la nube.', 'Configuración sincronizada');
        }
        success = true;
    } catch (error) {
        console.error('Error al cargar configuración en la nube:', error);
        updateRemoteConfigStatus('No se pudo cargar la configuración en la nube.', 'danger');
        if (!silent) {
            showAlert('No se pudo cargar la configuración desde la nube. Inténtalo nuevamente.', 'Error al cargar');
        }
    } finally {
        setRemoteConfigBusy(false);
        if (!success) {
            updateRemoteConfigMeta();
        }
    }

    return success;
}

async function saveCompanySettingsToCloud({ silent = false } = {}) {
    if (!isConfigSheetConfigured()) {
        updateRemoteConfigStatus('Configura COMPANY_CONFIG_SHEET_ID en js/app.js para habilitar la sincronización en la nube.', 'warning');
        if (!silent) {
            showAlert('Configura COMPANY_CONFIG_SHEET_ID en js/app.js para habilitar la sincronización en la nube.', 'Configuración incompleta');
        }
        return false;
    }

    if (!gapi?.client?.sheets) {
        updateRemoteConfigStatus('La API de Google Sheets no está lista. Intenta recargar.', 'danger');
        if (!silent) {
            showAlert('La API de Google Sheets no está lista. Recarga la página e inténtalo nuevamente.', 'Error de API');
        }
        return false;
    }

    const email = getCurrentUserEmail();
    if (!email) {
        updateRemoteConfigStatus('No se pudo determinar el correo de la sesión para sincronizar la configuración.', 'warning');
        if (!silent) {
            showAlert('Inicia sesión con una cuenta que entregue el correo electrónico para sincronizar la configuración.', 'Correo no disponible');
        }
        return false;
    }

    setRemoteConfigBusy(true, 'Sincronizando configuración en la nube...');
    let success = false;

    try {
        const payload = JSON.stringify({ ...DEFAULT_COMPANY_SETTINGS, ...companySettings });
        const nowIso = new Date().toISOString();
        let targetRow = companySettingsRowIndex;

        if (!targetRow) {
            const lookupResponse = await gapi.client.sheets.spreadsheets.values.get({
                spreadsheetId: COMPANY_CONFIG_SHEET_ID,
                range: COMPANY_CONFIG_SHEET_RANGE,
            });
            const values = lookupResponse.result.values || [];
            const lowerEmail = email.toLowerCase();
            for (let i = 0; i < values.length; i += 1) {
                if ((values[i][0] || '').toLowerCase() === lowerEmail) {
                    targetRow = i + 2;
                    break;
                }
            }
        }

        const rowValues = [[email, payload, nowIso]];

        if (targetRow) {
            const range = `${COMPANY_CONFIG_SHEET_TAB}!A${targetRow}:C${targetRow}`;
            await gapi.client.sheets.spreadsheets.values.update({
                spreadsheetId: COMPANY_CONFIG_SHEET_ID,
                range,
                valueInputOption: 'RAW',
                resource: { values: rowValues }
            });
            companySettingsRowIndex = targetRow;
        } else {
            const appendResponse = await gapi.client.sheets.spreadsheets.values.append({
                spreadsheetId: COMPANY_CONFIG_SHEET_ID,
                range: `${COMPANY_CONFIG_SHEET_TAB}!A:C`,
                valueInputOption: 'RAW',
                insertDataOption: 'INSERT_ROWS',
                resource: { values: rowValues }
            });
            const updatedRange = appendResponse.result?.updates?.updatedRange || '';
            const parsedRow = extractRowIndexFromRange(updatedRange);
            if (parsedRow) {
                companySettingsRowIndex = parsedRow;
            }
        }

        const now = new Date();
        remoteConfigLastSavedAt = now;
        remoteConfigLastLoadedAt = now;
        updateRemoteConfigStatus('Configuración sincronizada con la nube.', 'success');
        updateRemoteConfigMeta();
        if (!silent) {
            showAlert('La configuración se guardó correctamente en la nube.', 'Sincronización exitosa');
        }
        success = true;
    } catch (error) {
        console.error('Error al guardar configuración en la nube:', error);
        updateRemoteConfigStatus('No se pudo guardar la configuración en la nube.', 'danger');
        if (!silent) {
            showAlert('No se pudo guardar la configuración en la nube. Intenta nuevamente.', 'Error al sincronizar');
        }
    } finally {
        setRemoteConfigBusy(false);
        if (!success) {
            updateRemoteConfigMeta();
        }
    }

    return success;
}

function resetRemoteConfigState() {
    companySettingsRowIndex = null;
    remoteConfigLastLoadedAt = null;
    remoteConfigLastSavedAt = null;
    remoteConfigBusy = false;
    updateRemoteConfigButtonsState();
    refreshRemoteConfigUiState();
    updateRemoteConfigMeta();
}

// --- LÓGICA DE CONTRATOS Y PORTAL ---
 function renderContractsPage() {
     // Disparar evento change en el select de cliente para actualizar las cotizaciones
     if (contractClientSelect) {
         handleContractClientChange(); // Llamar directamente a la función que actualiza
     }
}

// Manejador de cambio en select de cliente (pestaña Contratos)
function handleContractClientChange() {
     // Chequeos robustos
     if (!contractClientSelect || !contractQuoteSelect || !contractText || !saveContractBtn) {
         console.warn("Faltan elementos en la pestaña de Contratos.");
         return;
      }
    const clientId = contractClientSelect.value;
    contractQuoteSelect.innerHTML = '<option value="">-- Seleccione cotización --</option>'; // Limpiar y poner default
    contractText.value = ''; // Limpiar área de texto
    saveContractBtn.disabled = true; // Deshabilitar botón de guardar

    if (clientId) {
        // Filtrar cotizaciones del cliente seleccionado que estén 'Aceptada'
        const clientQuotes = quotesHistory.filter(q => q && q.client && q.client.id === clientId && q.status === 'Aceptada');
        if (clientQuotes.length > 0) {
            clientQuotes.forEach(q => {
                const option = document.createElement('option');
                option.value = q.id;
                option.textContent = q.number; // Mostrar número de cotización
                contractQuoteSelect.appendChild(option);
            });
        } else {
             contractQuoteSelect.innerHTML = '<option value="">-- No hay cotizaciones aceptadas --</option>';
         }
    }
}

function deleteQuote(quoteId, rowId, googleDocId = '', googlePdfId = '') {
    if (!quoteId || !rowId) {
        showAlert('No se pudo determinar qué cotización eliminar.', 'Error al eliminar');
        return;
    }
    if (!QUOTES_SHEET_ID || QUOTES_SHEET_ID === 'ID_DE_TU_HOJA_DE_COTIZACIONES') {
        showAlert('Falta configurar el ID de la hoja de historial antes de eliminar.', 'Error de Configuración');
        return;
    }
    if (!gapi?.client?.sheets) {
        showAlert('La API de Google Sheets no está lista. Intenta nuevamente en unos segundos.', 'API no disponible');
        return;
    }

    const sanitizedDocId = (googleDocId && googleDocId !== 'null' && googleDocId !== 'undefined') ? googleDocId : '';
    const sanitizedPdfId = (googlePdfId && googlePdfId !== 'null' && googlePdfId !== 'undefined') ? googlePdfId : '';

    showConfirm('¿Deseas eliminar esta cotización del historial?\nEsta acción limpiará la fila en Google Sheets y enviará a la papelera los archivos relacionados en Drive.', async () => {
        try {
            if (authStatus) authStatus.innerText = `Eliminando cotización en la fila ${rowId}...`;
            const range = `Hoja 1!A${rowId}:K${rowId}`;
            await gapi.client.sheets.spreadsheets.values.clear({
                spreadsheetId: QUOTES_SHEET_ID,
                range
            });

            if (sanitizedDocId) {
                await trashDriveFile(sanitizedDocId);
            }
            if (sanitizedPdfId) {
                await trashDriveFile(sanitizedPdfId);
            }

            quotesHistory = quotesHistory.filter(q => q.id !== quoteId);
            if (currentQuoteContextId === quoteId) {
                currentQuoteContextId = null;
            }
            quoteCounter = quotesHistory.length + 1;
            renderQuotesHistory();
            showAlert('Cotización eliminada correctamente.', 'Cotización eliminada');
        } catch (error) {
            console.error('Error al eliminar la cotización:', error);
            const errorMsg = error.result?.error?.message || error.message || 'Error desconocido.';
            showAlert(`No se pudo eliminar la cotización: ${errorMsg}`, 'Error al eliminar');
        } finally {
            if (authStatus) authStatus.innerText = 'Listo.';
        }
    }, 'Eliminar cotización');
}

 // Manejador de cambio en select de cotización (pestaña Contratos)
 function handleContractQuoteChange() {
     // Chequeos
     if (!contractQuoteSelect || !contractText || !saveContractBtn) return;

     const quoteId = contractQuoteSelect.value;
     contractText.value = ''; // Limpiar siempre al cambiar
     saveContractBtn.disabled = !quoteId; // Habilitar botón solo si se selecciona una cotización válida

     if (quoteId) {
         contractText.placeholder = "Escribe aquí el contenido del contrato o carga desde Doc existente...";
         // TODO: Si la cotización tiene googleDocId, ofrecer cargar su contenido
         // const quote = quotesHistory.find(q => q.id === quoteId);
         // if(quote && quote.googleDocId) { ... }
     } else {
         contractText.placeholder = "Selecciona una cotización aceptada para redactar el contrato...";
      }
 }

// Renderizar página del portal cliente
function renderClientPortalPage() {
     if (portalClientSelect) {
         handlePortalClientChange();
     }
}

// Manejador de cambio en select de cliente (pestaña Portal)
function handlePortalClientChange() {
     // Chequeos
     if (!portalClientSelect || !clientPortalContent) {
         console.warn("Faltan elementos en la pestaña Portal Cliente.");
         return;
      }

    const clientId = portalClientSelect.value;
    clientPortalContent.innerHTML = '<p class="text-muted fst-italic">Cargando...</p>';

    if (!clientId) {
        clientPortalContent.innerHTML = '<p class="text-muted">Seleccione un cliente para ver su historial.</p>';
        return;
    }

    // Filtrar cotizaciones válidas del cliente
    const clientQuotes = quotesHistory.filter(q => q && q.client && q.client.id === clientId);

    if (clientQuotes.length === 0) {
        clientPortalContent.innerHTML = '<p class="text-muted">Este cliente no tiene cotizaciones registradas.</p>';
        return;
    }

    // Construir HTML
    let contentHtml = '';
    [...clientQuotes].reverse().forEach(quote => {
        const pdfLinkHtml = quote.googlePdfId
            ? `<a href="https://drive.google.com/file/d/${quote.googlePdfId}/view" target="_blank" class="link-danger small ms-2"><i class="fas fa-file-pdf me-1"></i>Ver PDF</a>`
            : '';
        contentHtml += `
            <div class="mb-4 pb-3 border-bottom">
                <div class="d-flex flex-wrap justify-content-between align-items-center gap-2 mb-1">
                    <h3 class="h6 fw-semibold text-dark mb-0">${quote.number || 'N/A'}</h3>
                    <span class="status-badge status-${quote.status || 'Pendiente'}">${quote.status || 'Pendiente'}</span>
                </div>
                <p class="small text-secondary mb-0">
                    <span class="me-3">Emitida: ${quote.issueDate || 'N/A'}</span> |
                    <span class="mx-3">Total: <b class="text-dark">$${(quote.total || 0).toFixed(2)}</b></span>
                    ${pdfLinkHtml}
                </p>
            </div>
        `;
    });
    clientPortalContent.innerHTML = contentHtml; 
}

// --- LÓGICA DE IMPRESIÓN ---
function buildPrintableHtml(quoteData) {
    if (!quoteData || !quoteData.companySettings || !quoteData.client || !Array.isArray(quoteData.items)) {
        console.error("buildPrintableHtml: Datos incompletos o inválidos:", quoteData);
        return '<p style="color: red; font-weight: bold; text-align: center; padding: 20px;">Error: No se pueden generar los datos de la cotización para imprimir.</p>';
    }

    const {
        companySettings: cs,
        client,
        items,
        number,
        issueDate,
        validityDate,
        subtotal,
        iva,
        total,
        notes
    } = quoteData;

    const primaryColor = cs.primaryColor || '#0ea5e9';
    const accentColor = cs.accentColor || '#22c55e';

    const companyName = cs.name || 'Nombre Empresa N/A';
    const companyAddress = cs.address || '';
    const companyContact = cs.contact || '';
    const companyWebsite = typeof cs.website === 'string' ? cs.website.trim() : '';
    const companyWhatsapp = typeof cs.whatsapp === 'string' ? cs.whatsapp.trim() : '';

    const clientName = client.name || 'Cliente N/A';
    const clientRuc = client.ruc || 'RUC/CI N/A';
    const clientContact = client.contact || '';

    const quoteNumber = number || 'N/A';
    const quoteIssueDate = issueDate || 'N/A';
    const quoteValidityDate = validityDate || 'N/A';
    const quoteSubtotal = typeof subtotal === 'number' ? subtotal : parseFloat(subtotal) || 0;
    const quoteIva = typeof iva === 'number' ? iva : parseFloat(iva) || 0;
    const quoteTotal = typeof total === 'number' ? total : parseFloat(total) || 0;
    const quoteNotes = notes || '';

    const logoUrl = typeof cs.logo === 'string' ? cs.logo.trim() : '';
    const companyInitials = companyName
        .split(/\s+/)
        .filter(Boolean)
        .slice(0, 2)
        .map(word => word.charAt(0).toUpperCase())
        .join('') || 'LOGO';

    const logoHtml = logoUrl
        ? `<img src="${escapeHtml(logoUrl)}" alt="Logo ${escapeHtml(companyName)}" class="quote-logo">`
        : `<div class="quote-logo-placeholder">${escapeHtml(companyInitials)}</div>`;

    const itemsHtml = items.map(item => {
        const itemName = item?.name ? escapeHtml(item.name) : 'Servicio/Producto N/A';
        const itemDescription = item?.description ? `<p class="quote-item-desc">${formatMultilineText(item.description)}</p>` : '';
        const parsedQty = Number(item?.quantity);
        const itemQty = Number.isFinite(parsedQty) && parsedQty > 0 ? parsedQty : 1;
        const parsedPrice = Number(item?.price);
        const itemPriceValue = Number.isFinite(parsedPrice) && parsedPrice >= 0 ? parsedPrice : 0;
        const itemTotalPrice = itemPriceValue * itemQty;
        return `
            <tr>
                <td>
                    <p class="quote-item-name">${itemName}</p>
                    ${itemDescription}
                </td>
                <td class="text-center">${itemQty}</td>
                <td class="text-end">$${itemPriceValue.toFixed(2)}</td>
                <td class="text-end">$${itemTotalPrice.toFixed(2)}</td>
            </tr>
        `;
    }).join('');

    const websiteDisplay = companyWebsite ? companyWebsite.replace(/^(https?:\/\/)?(www\.)?/, '') : '';
    const websiteHref = companyWebsite ? (companyWebsite.startsWith('http') ? companyWebsite : `https://${companyWebsite}`) : '';
    const whatsappHref = companyWhatsapp ? `https://wa.me/${companyWhatsapp.replace(/\D/g, '')}` : '';

    const footerLinks = [];
    if (websiteHref) {
        footerLinks.push(`<a href="${escapeHtml(websiteHref)}" target="_blank" rel="noopener">${escapeHtml(websiteDisplay)}</a>`);
    }
    if (companyWhatsapp) {
        footerLinks.push(`<a href="${escapeHtml(whatsappHref)}" target="_blank" rel="noopener"><i class="fab fa-whatsapp"></i> ${escapeHtml(companyWhatsapp)}</a>`);
    }

    const notesHtml = quoteNotes
        ? `<div class="quote-notes"><h4>Notas adicionales</h4><p>${formatMultilineText(quoteNotes)}</p></div>`
        : '';

    return `
        <div class="quote-print" style="--quote-primary: ${primaryColor}; --quote-accent: ${accentColor};">
            <header class="quote-print__header">
                <div class="quote-print__brand">
                    ${logoHtml}
                    <div>
                        <p class="quote-info-strong">${escapeHtml(companyName)}</p>
                        ${companyAddress ? `<p class="quote-info-text">${formatMultilineText(companyAddress)}</p>` : ''}
                        ${companyContact ? `<p class="quote-info-text">${formatMultilineText(companyContact)}</p>` : ''}
                    </div>
                </div>
                <div class="quote-print__heading">
                    <div class="quote-print__badge">Cotización</div>
                    <p class="quote-print__number">Nº:<strong>${escapeHtml(quoteNumber)}</strong></p>
                </div>
            </header>
            <section class="quote-info-grid">
                <article class="quote-info-card">
                    <h3>Cliente</h3>
                    <p class="quote-info-strong">${escapeHtml(clientName)}</p>
                    <p class="quote-info-meta">RUC/C.I.: ${escapeHtml(clientRuc)}</p>
                    ${clientContact ? `<p class="quote-info-text">${formatMultilineText(clientContact)}</p>` : ''}
                </article>
                <article class="quote-info-card">
                    <h3>De</h3>
                    <p class="quote-info-strong">${escapeHtml(companyName)}</p>
                    ${companyAddress ? `<p class="quote-info-text">${formatMultilineText(companyAddress)}</p>` : ''}
                    ${companyContact ? `<p class="quote-info-text">${formatMultilineText(companyContact)}</p>` : ''}
                </article>
                <article class="quote-info-card">
                    <h3>Fechas</h3>
                    <p class="quote-info-meta">Emisión</p>
                    <p class="quote-info-text">${escapeHtml(quoteIssueDate)}</p>
                    <p class="quote-info-meta">Válida hasta</p>
                    <p class="quote-info-text">${escapeHtml(quoteValidityDate)}</p>
                </article>
            </section>
            <section>
                <table class="quote-table">
                    <thead>
                        <tr>
                            <th scope="col">Descripción</th>
                            <th scope="col" style="width: 100px;" class="text-center">Cant.</th>
                            <th scope="col" style="width: 130px;" class="text-end">P. Unit.</th>
                            <th scope="col" style="width: 140px;" class="text-end">Total</th>
                        </tr>
                    </thead>
                    <tbody>${itemsHtml}</tbody>
                </table>
            </section>
            <section class="quote-totals">
                <div class="quote-totals-row"><span>Subtotal</span><span>$${quoteSubtotal.toFixed(2)}</span></div>
                <div class="quote-totals-row"><span>IVA ${(IVA_RATE * 100).toFixed(0)}%</span><span>$${quoteIva.toFixed(2)}</span></div>
                <div class="quote-total-final"><span>Total</span><span>$${quoteTotal.toFixed(2)}</span></div>
            </section>
            ${notesHtml}
            <footer class="quote-footer">
                ${footerLinks.length ? `<div class="quote-footer-links">${footerLinks.join('<span class="mx-1 text-muted">|</span>')}</div>` : ''}
                <p class="mb-1">Gracias por su preferencia.</p>
                <p class="mb-0">Esta cotización es un documento informativo generado por sistema.</p>
            </footer>
        </div>
    `;
}

// Función para reimprimir una cotización del historial
function reprintQuote(quoteId) {
    const quoteData = quotesHistory.find(q => q.id === quoteId);
    if (!quoteData) {
        showAlert("No se encontró la cotización en el historial para reimprimir.");
        return;
    }

    if (quoteData.googlePdfId) {
        window.open(`https://drive.google.com/file/d/${quoteData.googlePdfId}/view`, '_blank');
        return;
    }

    const printArea = document.getElementById('print-area');
    if (printArea) {
        const settingsForPrint = quoteData.companySettingsForPrint || companySettings;
        printArea.innerHTML = buildPrintableHtml({ ...quoteData, companySettings: settingsForPrint });
        setTimeout(() => { window.print(); }, 500);
    } else {
        console.error("Elemento 'print-area' no encontrado para reimprimir.");
    }
}


// --- INICIALIZACIÓN DE LA APP ---
document.addEventListener('DOMContentLoaded', () => {
     console.log("DOM listo. Asignando elementos y iniciando carga de APIs...");
     assignDOMElements(); // Asigna las variables de elementos del DOM
     initModal(); // Inicializa los elementos del modal
     // Pre-cargar las librerías de PDF en segundo plano para evitar bloqueos al generar la primera cotización
     ensurePdfLibrariesAvailable().catch(err => {
         console.warn('No se pudieron preparar las librerías de PDF durante la inicialización:', err);
     });
     initializeApp().catch(err => { // Iniciar carga de GAPI/GIS
         console.error("Fallo la inicialización general de la app:", err);
         if(authStatus) authStatus.innerText = "Error crítico al inicializar. Recarga la página.";
         // Podrías mostrar un mensaje más visible al usuario aquí
         showAlert("Error crítico al inicializar las APIs de Google. Por favor, recarga la página.", "Error Crítico");
      }); 
});

