// `msal.3.token.keys.${clientId}`

import { APP_TAG } from "./main";

// all tokens are encrypted, need to decrypt with cookie `msal.cache.encryption` to get the actual token
type MsalClientTokenEntry = {
    idToken: string[]; // encrypted secret stored in this localStorage key of format MsalIdTokenEntry
    accessToken: string[]; // accessTokens here are encrypted, need to decrypt to get the actual token
    refreshToken: string[];
}

// idToken in LocalStorage is stored as below, found through `msal.3.token.keys.${clientId}`
// `msal.3|${homeAccountId}|${environment}|idtoken|${clientId}|${tenantId}|${SCOPES.join(" ")}||`
/*type MsalIdTokenEntry = {
    credentialType: "IdToken";
    homeAccountId: string;
    environment: string;
    clientId: string;
    secret: string;
    realm: string;
    lastUpdatedAt: string;
}*/

// there is a LocalStoage item with these fields
type ActiveAccountFilters = {
    homeAccountId: string;
    localAccountId: string;
    tenantId: string;
}

interface MsalIds extends ActiveAccountFilters {
    clientId: string;
}

// this is just whatever was stored in the LocalStorage item
interface MsalAccessTokenEntry {
    homeAccountId: string;         // Usually "<uid>.<utid>"
    credentialType: "AccessToken"; // Always this for access tokens
    secret: string;                // The actual JWT for bearer token
    cachedAt: string;              // Unix timestamp in seconds (as string)
    expiresOn: string;             // Unix timestamp in seconds (as string)
    extendedExpiresOn?: string;   // Optional fallback expiry
    environment: string;          // e.g. "login.windows.net"
    clientId: string;             // Azure AD application ID (MSAL clientId)
    realm: string;                // Azure AD tenant ID
    target: string;               // Space-separated list of scopes
    tokenType: "Bearer";          // Usually always "Bearer"
}

// ---------------------------------- //
// The following line is copied from: //
// https://github.com/pionxzh/chatgpt-exporter/blob/0baa7b12a5f9bb93bdbe70ccb52d0c231361b411/src/api.ts#L596C1-L596C105
// Licensed under the MIT License     //
// ---------------------------------- //
const getCookie = (key: string) => document.cookie.match(`(^|;)\\s*${key}\\s*=\\s*([^;]+)`)?.pop() || ''

/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 * Obtained via:
 * - https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/afeaeddc777577b1b16f0084f5e5f9e4c15ee5e9/lib/msal-browser/src/cache/LocalStorage.ts
 * - https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/afeaeddc777577b1b16f0084f5e5f9e4c15ee5e9/lib/msal-browser/src/encode/Base64Decode.ts#L28
 * - https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/afeaeddc777577b1b16f0084f5e5f9e4c15ee5e9/lib/msal-browser/src/crypto/BrowserCrypto.ts#L351
 * Code includes minor modifications for this project.
 */

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
            throw Error(`${APP_TAG} Error extracting base64`);
    }
    const binString = atob(encodedString);
    return Uint8Array.from(binString, (m) => m.codePointAt(0) || 0);
}

/**
 * Helper function to convert Uint8Array<ArrayBufferLike> to ArrayBuffer
 * so that TSC stops yelling at me
 * @param bufferLike arraybufferlike object
 */
function toArrayBuffer(bufferLike: Uint8Array<ArrayBufferLike>): ArrayBuffer {
    return Uint8Array.from(bufferLike).buffer;
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
    nonce: Uint8Array<ArrayBufferLike>,/* was originally ArrayBuffer but TS yells at me */
    context: string
): Promise<CryptoKey> {
    return window.crypto.subtle.deriveKey(
        {
            name: HKDF,
            salt: toArrayBuffer(nonce),
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
        toArrayBuffer(encodedData)
    );

    return new TextDecoder().decode(decryptedData);
}

/**
 * Returns the raw key to be passed into the key derivation function
 * @param baseKey
 * @returns
 */
function generateHKDF(baseKey: /* was originally ArrayBuffer but TS yells at me */Uint8Array<ArrayBufferLike>): Promise<CryptoKey> {
    return window.crypto.subtle.importKey(RAW, toArrayBuffer(baseKey), HKDF, false, [
        DERIVE_KEY,
    ]);
}

async function getEncryptionCookie(): Promise<EncryptionCookie> {
    const cookieString = decodeURIComponent(getCookie(ENCRYPTION_KEY));
    let parsedCookie = { key: "", id: "" };
    if (cookieString) {
        try {
            parsedCookie = JSON.parse(cookieString);
        } catch (e) {
            throw Error(`${APP_TAG} Failed to parse encryption cookie`)
        }
    }
    if (parsedCookie.key && parsedCookie.id) {
        const baseKey = base64DecToArr(parsedCookie.key);
        return {
            id: parsedCookie.id,
            key: await generateHKDF(baseKey),
        };
    } else {
        throw Error(`${APP_TAG} No encryption cookie found`)
    }
}

// ---------------- //
// End of MSAL MIT-licensed code //
//----------------- //

export const getMsalIds = (): MsalIds => {
    // get M365 Copilot's client ID from the `window` variable 
    // const clientId = (unsafeWindow as any)?.msal?.clientIds?.[0] as string | undefined;
    const clientId = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1" // harcoded M365 Copilot Chat UUID

    // there should be a localstorage key containing the logged in account IDs
    // somewhere in the page so that MSAL can find it
    // profile id (https://graph.microsoft.com/v1.0/me)
    // org id (https://graph.microsoft.com/v1.0/organization)
    // officeweb has a different clientId than Copilot Chat
    const msalIds = localStorage.getItem("msal.3.account.keys");
    if (!msalIds) {
        throw Error(`${APP_TAG} No account keys found for Copilot application`);
    }
    const accountKeys = JSON.parse(msalIds) as String[];

    if (accountKeys.length === 0) {
        throw Error(`${APP_TAG} No account keys found for Copilot application`);
    }

    // I only have one account key for the Copilot application so I don't know what multiple accounts
    // look like and I don't want to write code for something I don't know.
    // my org's M365 setup uses this key: `msal.3|${homeAccountId}|${environment}|${tenantId}`
    const accountKey = accountKeys[0];
    const [homeAccountId, _1, tenantId] = accountKey.split('|');
    const [localAccountId, _2] = homeAccountId.split('.');

    return {
        localAccountId: localAccountId,
        tenantId: tenantId,
        homeAccountId: homeAccountId,
        clientId,
    }
}

export const getAccessToken = async (msalIds: MsalIds): Promise<string> => {
    const encryptionCookie = await getEncryptionCookie();

    // my university's M365 setup stores idToken, accessToken, and refreshToken in the key: `msal.3.token.keys.${clientId}`
    // `msal.version` reports 5.9.0  though... so I don't really know why it's 3 instead of 5 but I don't use MSAL myself
    // the tokens might be at different places in other MSAL versions but this is the only M365 account I have so ¯\(ツ)/¯,
    const tokenKeys = localStorage.getItem(`msal.3.token.keys.${msalIds.clientId}`);
    if (!tokenKeys) {
        throw Error(`${APP_TAG} No token keys found for Copilot application`);
    }
    const tokenKeysData = JSON.parse(tokenKeys) as MsalClientTokenEntry;

    // there should be a token among the access tokens under Copilot's client ID that has the Sydney scope
    // that contains the bearer token that Copilot Chat APIs use (guess it was a codename :P). My current
    // M365 Copilot Chat doesn't hit the Sydney API anymore (seems to directly mutate the page POSTing JSON
    // to https://m365.cloud.microsoft/chat/${conversationId}), but the Sydney API endpoints somehow still work.
    const sydneyKey = tokenKeysData.accessToken.find(token => token.includes("https://substrate.office.com/sydney/.default"));
    if (!sydneyKey) {
        throw Error(`${APP_TAG} No Sydney access token found for Copilot application`);
    }
    const sydneyTokenEntry = localStorage.getItem(sydneyKey);
    if (!sydneyTokenEntry) {
        throw Error(`${APP_TAG} No Sydney token found for Copilot application`);
    }

    // `msal.3.token.keys.${clientId}` tokens are all encrypted, need to decrypt with the encryption cookie at `msal.cache.encryption`
    const payload = JSON.parse(sydneyTokenEntry) as EncryptedData;
    const decryptedData = await decrypt(
        encryptionCookie.key,
        payload.nonce,
        msalIds.clientId, // context is usually client ID according to MSAL v4 source code: https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/afeaeddc777577b1b16f0084f5e5f9e4c15ee5e9/lib/msal-browser/src/cache/LocalStorage.ts#L302
        payload.data
    );
    const parsedDecryptedData = JSON.parse(decryptedData) as MsalAccessTokenEntry;
    return parsedDecryptedData.secret;
};