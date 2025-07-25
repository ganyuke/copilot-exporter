import { unsafeWindow } from 'vite-plugin-monkey/dist/client'

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

const ENCRYPTION_KEY = "msal.cache.encryption";

// Algorithms
// const PKCS1_V15_KEYGEN_ALG = "RSASSA-PKCS1-v1_5";
const AES_GCM = "AES-GCM";
const HKDF = "HKDF";
// SHA-256 hashing algorithm
const S256_HASH_ALG = "SHA-256";
// MOD length for PoP tokens
// const MODULUS_LENGTH = 2048;
// Public Exponent
// const PUBLIC_EXPONENT: Uint8Array = new Uint8Array([0x01, 0x00, 0x01]);
// UUID hex digits
// const UUID_CHARS = "0123456789abcdef";
// Array to store UINT32 random value
// const UINT32_ARR = new Uint32Array(1);

// Key Format
const RAW = "raw";
// Key Usages
const ENCRYPT = "encrypt";
const DECRYPT = "decrypt";
const DERIVE_KEY = "deriveKey";

// Suberror
// const SUBTLE_SUBERROR = "crypto_subtle_undefined";
/*
const keygenAlgorithmOptions: RsaHashedKeyGenParams = {
    name: PKCS1_V15_KEYGEN_ALG,
    hash: S256_HASH_ALG,
    modulusLength: MODULUS_LENGTH,
    publicExponent: PUBLIC_EXPONENT,
};*/

/**
 * Decodes base64 into Uint8Array
 * @param base64String
 */
function base64DecToArr(base64String: string): Uint8Array {
    let encodedString = base64String.replace(/-/g, "+").replace(/_/g, "/");
    switch (encodedString.length % 4) {
        case 0:
            break;
        case 2:
            encodedString += "==";
            break;
        case 3:
            encodedString += "=";
            break;
        default:
            throw Error("error extracting base64");
    }
    const binString = atob(encodedString);
    return Uint8Array.from(binString, (m) => m.codePointAt(0) || 0);
}

/**
 * Given a base key and a nonce generates a derived key to be used in encryption and decryption.
 * Note: every time we encrypt a new key is derived
 * @param baseKey
 * @param nonce
 * @returns
 */
async function deriveKey(
    baseKey: CryptoKey,
    nonce: Uint8Array<ArrayBufferLike>,//ArrayBuffer,
    context: string
): Promise<CryptoKey> {
    return window.crypto.subtle.deriveKey(
        {
            name: HKDF,
            salt: nonce,
            hash: S256_HASH_ALG,
            info: new TextEncoder().encode(context),
        },
        baseKey,
        { name: AES_GCM, length: 256 },
        false,
        [ENCRYPT, DECRYPT]
    );
}

/**
 * Decrypt data with the given key and nonce
 * @param key
 * @param nonce
 * @param encryptedData
 * @returns
 */
async function decrypt(
    baseKey: CryptoKey,
    nonce: string,
    context: string,
    encryptedData: string
): Promise<string> {
    const encodedData = base64DecToArr(encryptedData);
    const derivedKey = await deriveKey(baseKey, base64DecToArr(nonce), context);
    const decryptedData = await window.crypto.subtle.decrypt(
        {
            name: AES_GCM,
            iv: new Uint8Array(12), // New key is derived for every encrypt so we don't need a new nonce
        },
        derivedKey,
        encodedData
    );

    return new TextDecoder().decode(decryptedData);
}

/**
 * Returns the raw key to be passed into the key derivation function
 * @param baseKey
 * @returns
 */
function generateHKDF(baseKey: /*ArrayBuffer*/Uint8Array<ArrayBufferLike>): Promise<CryptoKey> {
    return window.crypto.subtle.importKey(RAW, baseKey, HKDF, false, [
        DERIVE_KEY,
    ]);
}

const getCookie = (key: string) => document.cookie.match(`(^|;)\\s*${key}\\s*=\\s*([^;]+)`)?.pop() || ''

async function getEncryptionCookie(): Promise<EncryptionCookie> {
    const cookieString = decodeURIComponent(getCookie(ENCRYPTION_KEY));
    let parsedCookie = { key: "", id: "" };
    if (cookieString) {
        try {
            parsedCookie = JSON.parse(cookieString);
        } catch (e) {
            throw Error("failed to parse encryption cookie")
        }
    }
    if (parsedCookie.key && parsedCookie.id) {
        const baseKey = base64DecToArr(parsedCookie.key);
        return {
            id: parsedCookie.id,
            key: await generateHKDF(baseKey),
        };
    } else {
        throw Error("no encryption cookie found")
    }
}

type EncryptionCookie = {
    id: string;
    key: CryptoKey;
};

type EncryptedData = {
    id: string;
    nonce: string;
    data: string;
    lastUpdatedAt: string;
};

type ActiveAccountFilters = {
    homeAccountId: string;
    localAccountId: string;
    tenantId: string;
}

export const getAccessToken = () => getEncryptionCookie().then(result => {
    // get M365 Copilot's client ID from the `window` variable 
    const clientId = (unsafeWindow as any)?.msal?.clientIds?.[0] as string | undefined;
    if (!clientId) {
        throw Error("No client ID found for Copilot application");
    }
    // there should be a localstorage key containing the logged in account IDs
    // profile id (https://graph.microsoft.com/v1.0/me)
    // org id (https://graph.microsoft.com/v1.0/organization)
    const accountIdsKey = `msal.${clientId}.active-account-filters`;
    const accountIdsItem = localStorage.getItem(accountIdsKey);
    if (!accountIdsItem) {
        throw Error("No account ids found for Copilot application");
    }
    const accountIds = JSON.parse(accountIdsItem) as ActiveAccountFilters;
    const { homeAccountId, tenantId } = accountIds;

    // the M365 Copilot uses the access token stored in LocalStorage with these scopes
    const SCOPES = [
        "https://substrate.office.com/activityfeed-internal.readwrite",
        "https://substrate.office.com/context.read",
        "https://substrate.office.com/files.read",
        "https://substrate.office.com/files.readwrite",
        "https://substrate.office.com/files.readwrite.shared",
        "https://substrate.office.com/group.readwrite.all",
        "https://substrate.office.com/notes.readwrite",
        "https://substrate.office.com/officefeed-internal.readwrite",
        "https://substrate.office.com/peoplepredictions-internal.read",
        "https://substrate.office.com/roamingusersettings.readwrite",
        "https://substrate.office.com/signals-internal.readwrite",
        "https://substrate.office.com/substratesearch-internal.readwrite",
        "https://substrate.office.com/usersettings.read",
        "https://substrate.office.com/usersettings.readwrite",
        "https://substrate.office.com/.default"
    ]
    const ACCESS_TOKEN_LS = `${homeAccountId}-login.windows.net-accesstoken-${clientId}-${tenantId}-${SCOPES.join(" ")}--`
    const lskv = localStorage.getItem(ACCESS_TOKEN_LS);
    if (!lskv) {
        throw Error("missing access token localstorage")
    }
    const payload = JSON.parse(lskv) as EncryptedData;
    return decrypt(
        result.key,
        payload.nonce,
        clientId, // context is usually client ID according to MSAL v4 source code: https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/afeaeddc777577b1b16f0084f5e5f9e4c15ee5e9/lib/msal-browser/src/cache/LocalStorage.ts#L302
        payload.data
    )
});
