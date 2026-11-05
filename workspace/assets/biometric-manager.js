/**
 * BiometricManager — Gestión unificada de autenticación biométrica
 *
 * Prioridad:
 *   1. Captor Morpho E3/E4 externo (si está conectado y su SDK está disponible)
 *   2. Lector de huella del PC vía WebAuthn (Windows Hello, Touch ID, etc.)
 *
 * La API pública es idéntica independientemente del dispositivo activo,
 * por lo que profile-biometric.js y login-biometric.js no necesitan conocer
 * cuál se está usando.
 */
class BiometricManager {

    constructor(options = {}) {
        this.apiBaseUrl = options.apiBaseUrl || '/api/biometric.php';

        // Callbacks opcionales
        this.onStatusChange     = options.onStatusChange     || null;
        this.onCaptureComplete  = options.onCaptureComplete  || null;
        this.onCaptureError     = options.onCaptureError     || null;

        // Estado interno
        this.activeMode         = null;   // 'morpho' | 'webauthn' | null
        this.morphoConnected    = false;
        this.webAuthnAvailable  = false;
        this.platformAuthAvail  = false;

        // Datos Morpho capturados
        this.capturedTemplate   = null;
        this.capturedHash       = null;
        this.deviceSerial       = null;
        this.morphoHandle       = null;

        // Último credential_id WebAuthn usado en auth
        this._webAuthnCredId    = null;
        this._webAuthnUserId    = null;

        this.captureTimeout = options.captureTimeout || 30000;
        this.qualityThreshold = options.qualityThreshold || 50;
    }

    // ─── Detección de capacidades ─────────────────────────────────────────────

    /**
     * Detecta qué dispositivos biométricos están disponibles.
     * Llama esto una vez al inicializar; establece this.activeMode.
     *
     * @returns {{ morpho: boolean, webauthn: boolean, platform: boolean, activeMode: string|null }}
     */
    async detectCapabilities() {
        // 1. Intentar Morpho E3/E4
        this.morphoConnected = await this._detectMorpho();

        // 2. WebAuthn disponible en el navegador?
        this.webAuthnAvailable = typeof window.PublicKeyCredential !== 'undefined';

        // 3. ¿Existe un autenticador de plataforma (lector del PC)?
        if (this.webAuthnAvailable) {
            try {
                this.platformAuthAvail = await PublicKeyCredential
                    .isUserVerifyingPlatformAuthenticatorAvailable();
            } catch {
                this.platformAuthAvail = false;
            }
        }

        // Elegir modo activo
        if (this.morphoConnected) {
            this.activeMode = 'morpho';
        } else if (this.webAuthnAvailable && this.platformAuthAvail) {
            this.activeMode = 'webauthn';
        } else {
            this.activeMode = null;
        }

        this._emitStatus(this.activeMode ? 'connected' : 'disconnected');

        return {
            morpho:     this.morphoConnected,
            webauthn:   this.webAuthnAvailable,
            platform:   this.platformAuthAvail,
            activeMode: this.activeMode,
        };
    }

    /** true si algún dispositivo biométrico está disponible */
    isConnected() {
        return this.activeMode !== null;
    }

    getActiveMode() {
        return this.activeMode;
    }

    getDeviceLabel() {
        if (this.activeMode === 'morpho')   return 'Captor Morpho E3/E4';
        if (this.activeMode === 'webauthn') return 'Lector de huella del PC';
        return 'Sin dispositivo';
    }

    // ─── API unificada ────────────────────────────────────────────────────────

    /**
     * Registra la huella digital del usuario.
     * - Morpho:   captura template físico → guarda en BD via API.
     * - WebAuthn: abre dialogo del SO → verifica con servidor.
     *
     * @returns {Promise<{ ok: boolean, message: string }>}
     */
    async registerBiometric() {
        if (this.activeMode === 'morpho') {
            return await this._morphoRegister();
        }
        if (this.activeMode === 'webauthn') {
            return await this._webAuthnRegister();
        }
        throw new Error('No hay dispositivo biométrico disponible.');
    }

    /**
     * Autentica al usuario con su huella.
     * @param {string} email
     * @returns {Promise<{ ok: boolean, user?: object, message: string }>}
     */
    async authenticateBiometric(email) {
        if (this.activeMode === 'morpho') {
            return await this._morphoAuthenticate(email);
        }
        if (this.activeMode === 'webauthn') {
            return await this._webAuthnAuthenticate(email);
        }
        throw new Error('No hay dispositivo biométrico disponible.');
    }

    /** Estado del usuario en la BD (Morpho + WebAuthn). */
    async checkBiometricStatus() {
        const r = await fetch(this.apiBaseUrl + '?action=check-biometric-status');
        return r.json();
    }

    /** Deshabilita la autenticación biométrica. */
    async disableBiometric(type = 'all') {
        const r = await fetch(this.apiBaseUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ action: 'disable-biometric', type }),
        });
        return r.json();
    }

    clearCapturedData() {
        this.capturedTemplate  = null;
        this.capturedHash      = null;
        this._webAuthnCredId   = null;
        this._webAuthnUserId   = null;
    }

    // ─── Morpho internals ─────────────────────────────────────────────────────

    async _detectMorpho() {
        if (typeof window.MorphoDevice !== 'undefined') {
            this.morphoHandle = window.MorphoDevice;
            try {
                const connected = await this.morphoHandle.isConnected?.();
                if (connected) {
                    this.deviceSerial = await this.morphoHandle.getSerial?.() || null;
                    return true;
                }
            } catch { /* nada */ }
        }
        return false;
    }

    async _morphoRegister() {
        try {
            this._emitStatus('capturing');
            const result = await this._morphoCapture('right_index');
            this.capturedTemplate = result.template;
            this.capturedHash     = result.hash;
            this._emitStatus('captured');
            if (this.onCaptureComplete) this.onCaptureComplete(result);

            // Guardar en BD
            const saveResp = await fetch(this.apiBaseUrl, {
                method:  'POST',
                headers: { 'Content-Type': 'application/json' },
                body:    JSON.stringify({
                    action:           'save-fingerprint',
                    template_data:    this.capturedTemplate,
                    fingerprint_hash: this.capturedHash,
                    device_serial:    this.deviceSerial,
                }),
            });
            const data = await saveResp.json();
            this._emitStatus(data.ok ? 'saved' : 'error');
            return data;
        } catch (err) {
            this._emitStatus('capture_error');
            if (this.onCaptureError) this.onCaptureError(err);
            throw err;
        }
    }

    async _morphoAuthenticate(email) {
        this._emitStatus('capturing');
        try {
            const result = await this._morphoCapture('right_index');
            this.capturedHash = result.hash;
            this._emitStatus('captured');

            const r = await fetch(this.apiBaseUrl, {
                method:  'POST',
                headers: { 'Content-Type': 'application/json' },
                body:    JSON.stringify({
                    action:           'verify-fingerprint',
                    email:            email,
                    fingerprint_hash: this.capturedHash,
                }),
            });
            const data = await r.json();
            this._emitStatus(data.ok ? 'verified' : 'verification_failed');
            return data;
        } catch (err) {
            this._emitStatus('capture_error');
            if (this.onCaptureError) this.onCaptureError(err);
            throw err;
        }
    }

    async _morphoCapture(fingerPosition) {
        if (this.morphoHandle) {
            const raw = await this.morphoHandle.capture({
                fingerPosition,
                qualityThreshold: this.qualityThreshold,
                format: 'iso_template',
            });
            const hash = await this._sha256(raw.template);
            return { template: raw.template, hash, quality: raw.quality || 100 };
        }
        // Modo simulación (desarrollo sin captor físico)
        console.warn('[BiometricManager] Morpho SDK no disponible: modo simulación activado.');
        const template = btoa(JSON.stringify({ ts: Date.now(), r: Math.random() }));
        const hash     = await this._sha256(template);
        return { template, hash, quality: 90, isSimulated: true };
    }

    // ─── WebAuthn internals ───────────────────────────────────────────────────

    async _webAuthnRegister() {
        // 1. Obtener opciones del servidor
        const optResp = await fetch(this.apiBaseUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ action: 'webauthn-register-options' }),
        });
        const { ok, options } = await optResp.json();
        if (!ok) throw new Error('No se pudieron obtener las opciones de registro.');

        // 2. Convertir base64url a ArrayBuffer
        const createOpts = {
            publicKey: {
                ...options,
                challenge: this._b64urlToBuffer(options.challenge),
                user: {
                    ...options.user,
                    id: this._b64urlToBuffer(options.user.id),
                },
                excludeCredentials: (options.excludeCredentials || []).map(c => ({
                    ...c,
                    id: this._b64urlToBuffer(c.id),
                })),
            },
        };

        this._emitStatus('capturing');

        // 3. Solicitar al navegador/SO
        let credential;
        try {
            credential = await navigator.credentials.create(createOpts);
        } catch (err) {
            this._emitStatus('capture_error');
            if (err.name === 'NotAllowedError') {
                throw new Error('Operación cancelada por el usuario.');
            }
            throw new Error('El lector de huella no respondió: ' + err.message);
        }

        this._emitStatus('captured');

        // 4. Serializar y enviar al servidor
        const payload = {
            action:   'webauthn-register-complete',
            response: {
                id:   credential.id,
                type: credential.type,
                response: {
                    clientDataJSON:    this._bufferToB64url(credential.response.clientDataJSON),
                    attestationObject: this._bufferToB64url(credential.response.attestationObject),
                },
            },
        };

        const verifyResp = await fetch(this.apiBaseUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify(payload),
        });
        const result = await verifyResp.json();
        this._emitStatus(result.ok ? 'saved' : 'error');
        return result;
    }

    async _webAuthnAuthenticate(email) {
        // 1. Obtener opciones (reto + credential IDs del usuario)
        const optResp = await fetch(this.apiBaseUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ action: 'webauthn-auth-options', email }),
        });
        const { ok, options, user_id } = await optResp.json();
        if (!ok) throw new Error('No se encontraron credenciales biometricas para este correo.');

        this._webAuthnUserId = user_id;

        const getOpts = {
            publicKey: {
                ...options,
                challenge: this._b64urlToBuffer(options.challenge),
                allowCredentials: (options.allowCredentials || []).map(c => ({
                    ...c,
                    id: this._b64urlToBuffer(c.id),
                })),
            },
        };

        this._emitStatus('capturing');

        // 2. Solicitar al SO
        let assertion;
        try {
            assertion = await navigator.credentials.get(getOpts);
        } catch (err) {
            this._emitStatus('capture_error');
            if (err.name === 'NotAllowedError') {
                throw new Error('Operación cancelada por el usuario.');
            }
            throw new Error('El lector de huella no respondió: ' + err.message);
        }

        this._emitStatus('captured');
        this._webAuthnCredId = assertion.id;

        // 3. Enviar aserción al servidor
        const payload = {
            action:        'webauthn-auth-complete',
            user_id:       user_id,
            credential_id: assertion.id,
            response: {
                clientDataJSON:    this._bufferToB64url(assertion.response.clientDataJSON),
                authenticatorData: this._bufferToB64url(assertion.response.authenticatorData),
                signature:         this._bufferToB64url(assertion.response.signature),
                userHandle:        assertion.response.userHandle
                    ? this._bufferToB64url(assertion.response.userHandle)
                    : null,
            },
        };

        const verifyResp = await fetch(this.apiBaseUrl, {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify(payload),
        });
        const result = await verifyResp.json();
        this._emitStatus(result.ok ? 'verified' : 'verification_failed');
        return result;
    }

    // ─── Utilidades ───────────────────────────────────────────────────────────

    async _sha256(str) {
        try {
            const buf  = new TextEncoder().encode(str);
            const hash = await crypto.subtle.digest('SHA-256', buf);
            return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
        } catch {
            // fallback sin Crypto API
            let h = 0;
            for (let i = 0; i < str.length; i++) {
                h = Math.imul(31, h) + str.charCodeAt(i) | 0;
            }
            return 'fb_' + Math.abs(h).toString(16);
        }
    }

    _b64urlToBuffer(b64url) {
        const b64 = b64url.replace(/-/g, '+').replace(/_/g, '/');
        const bin = atob(b64);
        const buf = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
        return buf.buffer;
    }

    _bufferToB64url(bufferOrArray) {
        const bytes = bufferOrArray instanceof ArrayBuffer
            ? new Uint8Array(bufferOrArray)
            : new Uint8Array(bufferOrArray);
        let bin = '';
        bytes.forEach(b => { bin += String.fromCharCode(b); });
        return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
    }

    _emitStatus(status) {
        if (this.onStatusChange) this.onStatusChange(status);
    }
}

// Alias para compatibilidad con código anterior
const MorphoBiometricManager = BiometricManager;
