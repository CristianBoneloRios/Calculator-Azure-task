/**
 * Profile Biometric Handler
 * Soporta Morpho E3/E4 y lector de huella del PC (WebAuthn).
 */

let biometricManager       = null;
let currentBiometricStatus = null;

async function initProfileBiometric() {
    biometricManager = new BiometricManager({ apiBaseUrl: '../api/biometric.php' });
    const caps = await biometricManager.detectCapabilities();
    updateBiometricDeviceStatusUI(caps);
    await refreshBiometricStatus();
    setupBiometricListeners();
}

async function refreshBiometricStatus() {
    try {
        const data = await biometricManager.checkBiometricStatus();
        if (data.ok) { currentBiometricStatus = data; renderBiometricStatusUI(data); }
    } catch (e) { console.error('Error estado biometrico:', e); }
}

function renderBiometricStatusUI(data) {
    const tag        = document.getElementById('profileBiometricTag');
    const enrollBtn  = document.getElementById('profileBiometricEnrollBtn');
    const disableBtn = document.getElementById('profileBiometricDisableBtn');
    const infoPanel  = document.getElementById('profileBiometricInfo');
    const hasMorpho   = data.morpho?.enrolled   && data.morpho?.enabled;
    const hasWebAuthn = data.webauthn?.enrolled;
    const hasAny      = hasMorpho || hasWebAuthn;
    if (hasAny) {
        tag.innerHTML = '<i class="fas fa-fingerprint"></i> Activa';
        tag.style.cssText = 'background:rgba(39,174,96,.2);color:#27ae60';
        enrollBtn.innerHTML = '<i class="fas fa-sync-alt"></i> Reemplazar huella';
        disableBtn.hidden = false; infoPanel.hidden = false;
        const deviceName = hasMorpho ? (data.morpho.device_type||'Morpho E3/E4') : 'Lector del PC (WebAuthn)';
        const verCount = hasMorpho ? (data.morpho.verification_count??0) : (data.webauthn.credential_count??1);
        const lastVerif = hasMorpho ? data.morpho.last_verified : data.webauthn.last_used;
        document.getElementById('profileBiometricDeviceType').textContent       = deviceName;
        document.getElementById('profileBiometricVerificationCount').textContent = verCount;
        document.getElementById('profileBiometricLastVerified').textContent      =
            lastVerif ? new Date(lastVerif).toLocaleString('es-ES') : 'Nunca';
    } else {
        tag.innerHTML = '<i class="fas fa-fingerprint"></i> No configurada';
        tag.style.cssText = 'background:rgba(231,76,60,.2);color:#e74c3c';
        enrollBtn.innerHTML = '<i class="fas fa-fingerprint"></i> Registrar huella digital';
        disableBtn.hidden = true; infoPanel.hidden = true;
    }
}

function updateBiometricDeviceStatusUI(caps) {
    const indicator  = document.getElementById('profileBiometricDeviceIndicator');
    const statusText = document.getElementById('profileBiometricDeviceStatusText');
    const enrollBtn  = document.getElementById('profileBiometricEnrollBtn');
    if (caps.activeMode === 'morpho') {
        indicator.style.color = '#27ae60';
        statusText.textContent = String.fromCharCode(10003) + ' Captor Morpho E3/E4 conectado';
        enrollBtn.disabled = false;
    } else if (caps.activeMode === 'webauthn') {
        indicator.style.color = '#3b82f6';
        statusText.textContent = String.fromCharCode(10003) + ' Lector de huella del PC detectado';
        enrollBtn.disabled = false;
    } else {
        indicator.style.color = '#e74c3c';
        statusText.textContent = String.fromCharCode(10007) + ' No se detecto ningun lector de huella';
        enrollBtn.disabled = true;
    }
}

function setupBiometricListeners() {
    document.getElementById('profileBiometricEnrollBtn')?.addEventListener('click',  startEnroll);
    document.getElementById('profileBiometricDisableBtn')?.addEventListener('click', removeBiometric);
    document.getElementById('profileBiometricCancelBtn')?.addEventListener('click',  cancelCapture);
    document.getElementById('profileBiometricRetryBtn')?.addEventListener('click',   startEnroll);
}

async function startEnroll() {
    const enrollBtn = document.getElementById('profileBiometricEnrollBtn');
    enrollBtn.disabled = true;
    const mode = biometricManager.getActiveMode();
    const prompt = mode === 'webauthn'
        ? 'Usa el lector de huella de tu PC cuando aparezca el dialogo del sistema...'
        : 'Acerca tu dedo al captor Morpho...';
    showCaptureStage(prompt);
    try {
        const result = await biometricManager.registerBiometric();
        showCaptureStage(null);
        if (result.ok) {
            await refreshBiometricStatus();
            showProfileNotification(result.message || 'Huella registrada exitosamente', 'success');
        } else { throw new Error(result.message); }
    } catch (err) {
        showCaptureStage(null);
        showProfileNotification(err.message || 'Error al registrar', 'error');
    } finally { enrollBtn.disabled = false; }
}

function showCaptureStage(message) {
    const stage = document.getElementById('profileBiometricCaptureStage');
    const confirm = document.getElementById('profileBiometricConfirm');
    const prompt  = document.getElementById('profileBiometricCapturePrompt');
    const bar     = document.getElementById('profileBiometricProgressBar');
    if (message === null) {
        stage.hidden = true; confirm.hidden = true;
        if (bar) bar.style.width = '0%'; return;
    }
    stage.hidden = false; confirm.hidden = true;
    if (prompt) prompt.textContent = message;
    if (bar) { bar.style.width = '0%'; setTimeout(() => { bar.style.transition='width 3s ease'; bar.style.width='85%'; }, 50); }
}

function cancelCapture() {
    showCaptureStage(null);
    biometricManager.clearCapturedData();
    document.getElementById('profileBiometricEnrollBtn').disabled = false;
}

async function removeBiometric() {
    if (!confirm('Eliminar la huella registrada? Ya no podras autenticarte con ella.')) return;
    try {
        const r = await biometricManager.disableBiometric('all');
        if (r.ok) { await refreshBiometricStatus(); showProfileNotification('Autenticacion biometrica eliminada', 'success'); }
    } catch (e) { showProfileNotification(e.message||'Error', 'error'); }
}

function showProfileNotification(msg, type) {
    const ex = document.getElementById('_bioNotif'); if (ex) ex.remove();
    const n = document.createElement('div'); n.id = '_bioNotif';
    const bg = type==='success'?'#27ae60':type==='error'?'#e74c3c':'#3b82f6';
    n.style.cssText = `position:fixed;top:20px;right:20px;padding:12px 20px;background:${bg};color:#fff;border-radius:8px;z-index:9999;font-size:14px;display:flex;align-items:center;gap:10px;max-width:360px`;
    const icon = type==='success'?'check-circle':type==='error'?'exclamation-circle':'info-circle';
    n.innerHTML = `<i class="fas fa-${icon}"></i><span>${msg}</span>`;
    document.body.appendChild(n); setTimeout(()=>n.remove(),4000);
}

if (document.readyState==='loading') { document.addEventListener('DOMContentLoaded',initProfileBiometric); } else { initProfileBiometric(); }
