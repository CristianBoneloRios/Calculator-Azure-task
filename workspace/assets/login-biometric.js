/**
 * Login Biometric Handler
 * Soporta Morpho E3/E4 y lector de huella del PC (WebAuthn).
 */

let loginBiometricManager = null;

async function initLoginBiometric() {
    loginBiometricManager = new BiometricManager({ apiBaseUrl: '/api/biometric.php' });
    const caps = await loginBiometricManager.detectCapabilities();
    renderLoginBiometricUI(caps);
}

function renderLoginBiometricUI(caps) {
    const statusDiv   = document.getElementById('loginBiometricStatus');
    const formDiv     = document.getElementById('loginBiometricForm');
    const disabledDiv = document.getElementById('loginBiometricDisabled');
    const indicator   = document.getElementById('loginBiometricDeviceIndicator');
    const statusText  = document.getElementById('loginBiometricStatusText');
    const btn         = document.getElementById('loginBiometricBtn');

    if (caps.activeMode === 'morpho') {
        if (statusDiv)   statusDiv.style.display   = 'none';
        if (disabledDiv) disabledDiv.style.display = 'none';
        if (formDiv)     formDiv.hidden             = false;
        if (indicator)   indicator.style.color      = '#27ae60';
        if (statusText)  statusText.textContent     = String.fromCharCode(10003)+' Captor Morpho E3/E4 conectado';
        if (btn)         btn.innerHTML              = '<i class="fas fa-fingerprint"></i> Autenticarse con captor Morpho';
    } else if (caps.activeMode === 'webauthn') {
        if (statusDiv)   statusDiv.style.display   = 'none';
        if (disabledDiv) disabledDiv.style.display = 'none';
        if (formDiv)     formDiv.hidden             = false;
        if (indicator)   indicator.style.color      = '#3b82f6';
        if (statusText)  statusText.textContent     = String.fromCharCode(10003)+' Lector de huella del PC detectado';
        if (btn)         btn.innerHTML              = '<i class="fas fa-fingerprint"></i> Autenticarse con huella del PC';
    } else {
        if (statusDiv)   statusDiv.style.display   = 'block';
        if (disabledDiv) disabledDiv.style.display = 'block';
        if (formDiv)     formDiv.hidden             = true;
        if (indicator)   indicator.style.color      = '#e74c3c';
        if (statusText)  statusText.textContent     = String.fromCharCode(10007)+' No se detecto ningun lector de huella';
    }

    // Configurar listeners
    document.getElementById('loginBiometricBtn')?.addEventListener('click', startLoginAuth);
    document.getElementById('loginBiometricCancelBtn')?.addEventListener('click', cancelLoginAuth);

    // Sincronizar email con el formulario principal
    const mainEmail = document.getElementById('loginEmail');
    const bioEmail  = document.getElementById('loginBiometricEmail');
    if (mainEmail && bioEmail) {
        mainEmail.addEventListener('input', () => { bioEmail.value = mainEmail.value; });
    }
}

async function startLoginAuth() {
    const emailInput = document.getElementById('loginBiometricEmail');
    const email      = (emailInput?.value || '').trim();

    if (!email || !email.includes('@')) {
        showLoginNotif('Por favor ingresa tu correo electronico', 'error'); return;
    }

    const btn       = document.getElementById('loginBiometricBtn');
    const cancelBtn = document.getElementById('loginBiometricCancelBtn');
    const capture   = document.getElementById('loginBiometricCapture');
    const prompt    = document.getElementById('loginBiometricPrompt');
    const bar       = document.getElementById('loginBiometricProgressBar');

    btn.hidden       = true;
    cancelBtn.hidden = false;
    if (capture) capture.hidden = false;

    const mode    = loginBiometricManager.getActiveMode();
    const message = mode === 'webauthn'
        ? 'Usa el lector de huella de tu PC cuando aparezca el dialogo...'
        : 'Acerca tu dedo al captor Morpho...';
    if (prompt) prompt.textContent = message;
    if (bar)    { bar.style.width = '0%'; setTimeout(()=>{ bar.style.transition='width 3s ease'; bar.style.width='80%'; },50); }

    try {
        const result = await loginBiometricManager.authenticateBiometric(email);
        if (result.ok) {
            if (prompt) prompt.textContent = String.fromCharCode(10003)+' Identidad verificada';
            if (bar) bar.style.width = '100%';
            setTimeout(() => {
                if (result.requires_2fa) { window.location.href = 'verify-2fa.php'; }
                else { window.location.href = 'index.php'; }
            }, 800);
        } else { throw new Error(result.message); }
    } catch (err) {
        if (prompt) { prompt.textContent = String.fromCharCode(10007)+' '+err.message; prompt.style.color='#e74c3c'; }
        if (bar)    { bar.style.background='#e74c3c'; bar.style.width='0%'; }
        showLoginNotif(err.message||'Error en autenticacion', 'error');
        setTimeout(cancelLoginAuth, 2500);
    }
}

function cancelLoginAuth() {
    const btn       = document.getElementById('loginBiometricBtn');
    const cancelBtn = document.getElementById('loginBiometricCancelBtn');
    const capture   = document.getElementById('loginBiometricCapture');
    const prompt    = document.getElementById('loginBiometricPrompt');
    const bar       = document.getElementById('loginBiometricProgressBar');

    btn.hidden       = false;
    cancelBtn.hidden = true;
    if (capture) capture.hidden = true;
    if (prompt)  { prompt.textContent = 'Acerca tu dedo al lector...'; prompt.style.color = ''; }
    if (bar)     { bar.style.background='linear-gradient(90deg,#3b82f6,#06b6d4)'; bar.style.width='0%'; }
    loginBiometricManager.clearCapturedData();
}

function showLoginNotif(msg, type) {
    const ex = document.getElementById('_loginBioNotif'); if (ex) ex.remove();
    const n = document.createElement('div'); n.id = '_loginBioNotif';
    const bg = type==='success'?'#27ae60':type==='error'?'#e74c3c':'#3b82f6';
    n.style.cssText = `position:fixed;top:20px;right:20px;padding:12px 20px;background:${bg};color:#fff;border-radius:8px;z-index:9999;font-size:14px;display:flex;align-items:center;gap:10px;max-width:360px`;
    const icon = type==='success'?'check-circle':type==='error'?'exclamation-circle':'info-circle';
    n.innerHTML = `<i class="fas fa-${icon}"></i><span>${msg}</span>`;
    document.body.appendChild(n); setTimeout(()=>n.remove(),4000);
}

if (document.readyState==='loading') { document.addEventListener('DOMContentLoaded',initLoginBiometric); } else { initLoginBiometric(); }
