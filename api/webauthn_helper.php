<?php

declare(strict_types=1);

/**
 * WebAuthn helper — CBOR decoder + registration/assertion verification
 * Supports ES256 (ECDSA P-256) and RS256 (RSA-PKCS1-SHA256)
 * Zero external dependencies; uses only PHP core + OpenSSL extension.
 */

// ─── Base64URL ────────────────────────────────────────────────────────────────

function webauthn_b64url_encode(string $bytes): string
{
    return rtrim(strtr(base64_encode($bytes), '+/', '-_'), '=');
}

function webauthn_b64url_decode(string $str): string
{
    $padded = str_replace(['-', '_'], ['+', '/'], $str);
    $mod    = strlen($padded) % 4;
    if ($mod !== 0) {
        $padded .= str_repeat('=', 4 - $mod);
    }
    $decoded = base64_decode($padded, true);
    if ($decoded === false) {
        throw new RuntimeException('webauthn: invalid base64url string');
    }
    return $decoded;
}

// ─── Minimal CBOR decoder ────────────────────────────────────────────────────

function cbor_decode(string $data): mixed
{
    $pos = 0;
    $value = _cbor_item($data, $pos);
    return $value;
}

function _cbor_item(string $data, int &$pos): mixed
{
    $len = strlen($data);
    if ($pos >= $len) {
        throw new RuntimeException('CBOR: unexpected end of input at position ' . $pos);
    }

    $b  = ord($data[$pos++]);
    $mt = ($b >> 5) & 0x07;
    $ai = $b & 0x1f;

    $arg = _cbor_arg($data, $pos, $ai);

    switch ($mt) {
        case 0:
            return $arg;
        case 1:
            return -1 - $arg;
        case 2:
            $v = substr($data, $pos, $arg);
            $pos += $arg;
            return $v;
        case 3:
            $v = substr($data, $pos, $arg);
            $pos += $arg;
            return $v;
        case 4:
            $arr = [];
            for ($i = 0; $i < $arg; $i++) {
                $arr[] = _cbor_item($data, $pos);
            }
            return $arr;
        case 5:
            $map = [];
            for ($i = 0; $i < $arg; $i++) {
                $k        = _cbor_item($data, $pos);
                $map[$k]  = _cbor_item($data, $pos);
            }
            return $map;
        case 6:
            return _cbor_item($data, $pos);
        case 7:
            if ($ai === 20) { return false; }
            if ($ai === 21) { return true;  }
            if ($ai === 22 || $ai === 23) { return null; }
            // 16-bit / 32-bit / 64-bit floats not needed for WebAuthn
            throw new RuntimeException("CBOR: unsupported float/simple ai=$ai");
        default:
            throw new RuntimeException("CBOR: unknown major type $mt");
    }
}

function _cbor_arg(string $data, int &$pos, int $ai): int
{
    if ($ai < 24) {
        return $ai;
    }
    if ($ai === 24) {
        return ord($data[$pos++]);
    }
    if ($ai === 25) {
        $v = unpack('n', substr($data, $pos, 2))[1];
        $pos += 2;
        return $v;
    }
    if ($ai === 26) {
        $v = unpack('N', substr($data, $pos, 4))[1];
        $pos += 4;
        return $v;
    }
    if ($ai === 27) {
        [$hi, $lo] = array_values(unpack('N2', substr($data, $pos, 8)));
        $pos += 8;
        return (int) ($hi * 4294967296 + $lo);
    }
    throw new RuntimeException("CBOR: unsupported additional info $ai");
}

// ─── DER / PEM helpers ───────────────────────────────────────────────────────

function _der_tlv(int $tag, string $value): string
{
    $l = strlen($value);
    if ($l < 128) {
        return chr($tag) . chr($l) . $value;
    }
    if ($l < 256) {
        return chr($tag) . "\x81" . chr($l) . $value;
    }
    if ($l < 65536) {
        return chr($tag) . "\x82" . chr($l >> 8) . chr($l & 0xff) . $value;
    }
    throw new RuntimeException('DER: value too long (' . $l . ' bytes)');
}

function _der_int(string $bytes): string
{
    $bytes = ltrim($bytes, "\x00");
    if ($bytes === '') {
        $bytes = "\x00";
    }
    if (ord($bytes[0]) & 0x80) {
        $bytes = "\x00" . $bytes;
    }
    return _der_tlv(0x02, $bytes);
}

/**
 * Convert a COSE key (decoded from authenticatorData) to a PEM public key.
 * Supported: kty=2/alg=-7 (EC P-256 ES256) and kty=3/alg=-257 (RSA RS256).
 */
function cose_key_to_pem(array $cose): string
{
    $kty = $cose[1] ?? null;
    $alg = $cose[3] ?? null;

    // ── EC P-256 (ES256) ──────────────────────────────────────────────────────
    if ($kty === 2) {
        $x = $cose[-2] ?? null;
        $y = $cose[-3] ?? null;
        if ($x === null || $y === null) {
            throw new RuntimeException('COSE EC key missing x or y coordinate');
        }
        // Pad to exactly 32 bytes
        $x = str_pad($x, 32, "\x00", STR_PAD_LEFT);
        $y = str_pad($y, 32, "\x00", STR_PAD_LEFT);

        // SubjectPublicKeyInfo DER for EC P-256
        $spki = _der_tlv(0x30,
            _der_tlv(0x30,
                "\x06\x07\x2a\x86\x48\xce\x3d\x02\x01"        // OID ecPublicKey
              . "\x06\x08\x2a\x86\x48\xce\x3d\x03\x01\x07"    // OID prime256v1
            )
          . _der_tlv(0x03, "\x00\x04" . $x . $y)               // BIT STRING uncompressed
        );

        return "-----BEGIN PUBLIC KEY-----\n"
             . chunk_split(base64_encode($spki), 64, "\n")
             . "-----END PUBLIC KEY-----\n";
    }

    // ── RSA (RS256) ────────────────────────────────────────────────────────────
    if ($kty === 3) {
        $n = $cose[-1] ?? null;
        $e = $cose[-2] ?? null;
        if ($n === null || $e === null) {
            throw new RuntimeException('COSE RSA key missing modulus or exponent');
        }
        $rsaSeq = _der_tlv(0x30, _der_int($n) . _der_int($e));
        $spki   = _der_tlv(0x30,
            _der_tlv(0x30,
                "\x06\x09\x2a\x86\x48\x86\xf7\x0d\x01\x01\x01"  // OID rsaEncryption
              . "\x05\x00"
            )
          . _der_tlv(0x03, "\x00" . $rsaSeq)
        );

        return "-----BEGIN PUBLIC KEY-----\n"
             . chunk_split(base64_encode($spki), 64, "\n")
             . "-----END PUBLIC KEY-----\n";
    }

    throw new RuntimeException("COSE: unsupported key type kty=$kty alg=$alg");
}

// ─── Parse authenticatorData ─────────────────────────────────────────────────

/**
 * @return array{rpIdHash:string, flags:int, signCount:int, credentialId:string|null, publicKeyPem:string|null}
 */
function webauthn_parse_auth_data(string $authData): array
{
    if (strlen($authData) < 37) {
        throw new RuntimeException('authenticatorData: too short (' . strlen($authData) . ' bytes)');
    }

    $rpIdHash  = substr($authData, 0, 32);
    $flags     = ord($authData[32]);
    $signCount = unpack('N', substr($authData, 33, 4))[1];

    $result = [
        'rpIdHash'     => $rpIdHash,
        'flags'        => $flags,
        'signCount'    => $signCount,
        'credentialId' => null,
        'publicKeyPem' => null,
    ];

    // AT flag (bit 6) = attested credential data present
    if ($flags & 0x40) {
        if (strlen($authData) < 55) {
            throw new RuntimeException('authenticatorData: AT flag set but too short');
        }
        // AAGUID is at bytes 37-52 (16 bytes), skip it
        $credIdLen = unpack('n', substr($authData, 53, 2))[1];
        if ($credIdLen === 0) {
            throw new RuntimeException('authenticatorData: credentialIdLength is 0');
        }
        if (55 + $credIdLen > strlen($authData)) {
            throw new RuntimeException('authenticatorData: credentialId exceeds buffer');
        }
        $result['credentialId'] = substr($authData, 55, $credIdLen);

        $coseStart = 55 + $credIdLen;
        $coseBytes = substr($authData, $coseStart);
        if ($coseBytes !== '') {
            $coseKey = cbor_decode($coseBytes);
            if (is_array($coseKey)) {
                $result['publicKeyPem'] = cose_key_to_pem($coseKey);
            }
        }
    }

    return $result;
}

// ─── Relying-party helpers ────────────────────────────────────────────────────

function webauthn_rp_id(): string
{
    $host = $_SERVER['HTTP_HOST'] ?? 'localhost';
    return strtolower(explode(':', $host)[0]);
}

function webauthn_expected_origin(): string
{
    $https  = (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off')
              || ($_SERVER['SERVER_PORT'] ?? '80') === '443';
    $scheme = $https ? 'https' : 'http';
    return $scheme . '://' . ($_SERVER['HTTP_HOST'] ?? 'localhost');
}

// ─── Registration ─────────────────────────────────────────────────────────────

/**
 * Build a PublicKeyCredentialCreationOptions array to send to the browser.
 * Stores the raw challenge bytes in $_SESSION['_webauthn_reg_challenge'].
 */
function webauthn_registration_options(array $user): array
{
    $challenge  = random_bytes(32);
    $_SESSION['_webauthn_reg_challenge'] = $challenge;

    return [
        'rp'      => [
            'name' => 'Azure Task Suite',
            'id'   => webauthn_rp_id(),
        ],
        'user'    => [
            'id'          => webauthn_b64url_encode((string) $user['id']),
            'name'        => $user['email'],
            'displayName' => $user['full_name'] ?? $user['email'],
        ],
        'challenge'        => webauthn_b64url_encode($challenge),
        'pubKeyCredParams' => [
            ['alg' => -7,   'type' => 'public-key'],   // ES256
            ['alg' => -257, 'type' => 'public-key'],   // RS256
        ],
        'timeout'          => 60000,
        'authenticatorSelection' => [
            'authenticatorAttachment' => 'platform',   // built-in reader only
            'requireResidentKey'      => false,
            'userVerification'        => 'preferred',
        ],
        'attestation' => 'none',
    ];
}

/**
 * Verify the registration response, return credential data to store.
 *
 * @return array{credentialId:string, publicKeyPem:string, signCount:int}
 */
function webauthn_verify_registration(array $responseJson): array
{
    if (!isset($_SESSION['_webauthn_reg_challenge'])) {
        throw new RuntimeException('webauthn: no pending registration challenge in session');
    }
    $challenge = (string) $_SESSION['_webauthn_reg_challenge'];
    unset($_SESSION['_webauthn_reg_challenge']);

    // 1. Decode clientDataJSON
    $clientDataRaw = webauthn_b64url_decode($responseJson['response']['clientDataJSON']);
    $clientData    = json_decode($clientDataRaw, true);
    if ($clientData === null) {
        throw new RuntimeException('webauthn: failed to parse clientDataJSON');
    }

    if ($clientData['type'] !== 'webauthn.create') {
        throw new RuntimeException('webauthn: unexpected clientData type: ' . $clientData['type']);
    }

    $receivedChallenge = webauthn_b64url_decode($clientData['challenge']);
    if (!hash_equals($challenge, $receivedChallenge)) {
        throw new RuntimeException('webauthn: challenge mismatch during registration');
    }

    $expectedOrigin = webauthn_expected_origin();
    if ($clientData['origin'] !== $expectedOrigin) {
        throw new RuntimeException(
            "webauthn: origin mismatch — got '{$clientData['origin']}', expected '$expectedOrigin'"
        );
    }

    // 2. Decode attestationObject (CBOR)
    $attestation = cbor_decode(webauthn_b64url_decode($responseJson['response']['attestationObject']));
    if (!is_array($attestation) || !isset($attestation['authData'])) {
        throw new RuntimeException('webauthn: attestationObject missing authData');
    }

    // 3. Parse authenticatorData
    $parsed = webauthn_parse_auth_data((string) $attestation['authData']);

    // 4. Verify rpIdHash
    $expectedRpIdHash = hash('sha256', webauthn_rp_id(), true);
    if (!hash_equals($expectedRpIdHash, $parsed['rpIdHash'])) {
        throw new RuntimeException('webauthn: rpIdHash mismatch during registration');
    }

    // 5. UP flag (bit 0) must be set
    if (!($parsed['flags'] & 0x01)) {
        throw new RuntimeException('webauthn: user presence (UP) flag not set');
    }

    if ($parsed['credentialId'] === null || $parsed['publicKeyPem'] === null) {
        throw new RuntimeException('webauthn: no credential data in authenticatorData');
    }

    return [
        'credentialId' => webauthn_b64url_encode($parsed['credentialId']),
        'publicKeyPem' => $parsed['publicKeyPem'],
        'signCount'    => $parsed['signCount'],
    ];
}

// ─── Authentication ───────────────────────────────────────────────────────────

/**
 * Build a PublicKeyCredentialRequestOptions array.
 * Stores the raw challenge in $_SESSION['_webauthn_auth_challenge'].
 */
function webauthn_auth_options(array $credentials): array
{
    $challenge = random_bytes(32);
    $_SESSION['_webauthn_auth_challenge'] = $challenge;

    $allowCredentials = array_map(fn($c) => [
        'id'         => $c['credential_id'],
        'type'       => 'public-key',
        'transports' => ['internal'],
    ], $credentials);

    return [
        'challenge'        => webauthn_b64url_encode($challenge),
        'timeout'          => 60000,
        'rpId'             => webauthn_rp_id(),
        'allowCredentials' => $allowCredentials,
        'userVerification' => 'preferred',
    ];
}

/**
 * Verify an authentication assertion.
 *
 * @param  array $responseJson  The client's credential.response JSON
 * @param  array $storedCred    ['credential_id', 'public_key_pem', 'sign_count']
 * @return int   The new signCount to save
 */
function webauthn_verify_assertion(array $responseJson, array $storedCred): int
{
    if (!isset($_SESSION['_webauthn_auth_challenge'])) {
        throw new RuntimeException('webauthn: no pending authentication challenge in session');
    }
    $challenge = (string) $_SESSION['_webauthn_auth_challenge'];
    unset($_SESSION['_webauthn_auth_challenge']);

    // 1. Decode clientDataJSON
    $clientDataRaw = webauthn_b64url_decode($responseJson['response']['clientDataJSON']);
    $clientData    = json_decode($clientDataRaw, true);
    if ($clientData === null) {
        throw new RuntimeException('webauthn: failed to parse clientDataJSON');
    }

    if ($clientData['type'] !== 'webauthn.get') {
        throw new RuntimeException('webauthn: unexpected clientData type: ' . $clientData['type']);
    }

    $receivedChallenge = webauthn_b64url_decode($clientData['challenge']);
    if (!hash_equals($challenge, $receivedChallenge)) {
        throw new RuntimeException('webauthn: challenge mismatch during authentication');
    }

    $expectedOrigin = webauthn_expected_origin();
    if ($clientData['origin'] !== $expectedOrigin) {
        throw new RuntimeException(
            "webauthn: origin mismatch — got '{$clientData['origin']}', expected '$expectedOrigin'"
        );
    }

    // 2. Parse authenticatorData
    $authDataRaw = webauthn_b64url_decode($responseJson['response']['authenticatorData']);
    $parsed      = webauthn_parse_auth_data($authDataRaw);

    // 3. Verify rpIdHash
    $expectedRpIdHash = hash('sha256', webauthn_rp_id(), true);
    if (!hash_equals($expectedRpIdHash, $parsed['rpIdHash'])) {
        throw new RuntimeException('webauthn: rpIdHash mismatch during authentication');
    }

    // 4. UP flag must be set
    if (!($parsed['flags'] & 0x01)) {
        throw new RuntimeException('webauthn: user presence flag not set');
    }

    // 5. Verify signature
    $clientDataHash = hash('sha256', $clientDataRaw, true);
    $verifyData     = $authDataRaw . $clientDataHash;
    $signature      = webauthn_b64url_decode($responseJson['response']['signature']);

    $pubKey = openssl_pkey_get_public($storedCred['public_key_pem']);
    if ($pubKey === false) {
        throw new RuntimeException('webauthn: cannot load stored public key — ' . openssl_error_string());
    }

    $ok = openssl_verify($verifyData, $signature, $pubKey, OPENSSL_ALGO_SHA256);
    if ($ok !== 1) {
        throw new RuntimeException('webauthn: signature verification failed — ' . openssl_error_string());
    }

    // 6. signCount replay protection
    $storedCount = (int) $storedCred['sign_count'];
    $newCount    = (int) $parsed['signCount'];
    if ($storedCount > 0 && $newCount <= $storedCount) {
        throw new RuntimeException('webauthn: signCount replay attack detected');
    }

    return $newCount;
}
