/**
 * officeInterop.js
 * ================
 * Provides Office.js and Word API interop functions callable from Blazor
 * via IJSRuntime.InvokeAsync.
 *
 * All functions return a result object { success, data/error } so Blazor
 * can handle failures without unhandled JS exceptions.
 */

window.officeInterop = {

    /**
     * Acquires an Office SSO access token for the add-in backend API.
     *
     * The token audience is the Application ID URI configured in the Entra
     * app registration (api://HOSTNAME/CLIENT_ID). The backend validates
     * this token and authorises the user.
     *
     * Error codes:
     *   13001 — User not signed in to Office
     *   13002 — User has not consented to the required scopes
     *   13003 — User type not supported (consumer MSA)
     *   13005 — Add-in not registered as trusted (sideload / config issue)
     *   13006 — Office client error; advise sign-out/sign-in
     *   13007 — Another SSO operation is in progress
     *   13008 — Token request timed out
     *   13012 — Add-in not sideloaded or WebApplicationInfo misconfigured
     *
     * @returns {Promise<{success: boolean, token?: string, errorCode?: number, errorMessage?: string}>}
     */
    getOfficeToken: async function () {
        try {
            // Guard: when running in a plain browser (local dev), Office APIs are not available.
            if (typeof Office === 'undefined' || !Office.auth || typeof Office.auth.getAccessToken !== 'function') {
                console.warn('[officeInterop] getOfficeToken called outside of Office host');
                return {
                    success: false,
                    errorCode: 0,
                    errorMessage: 'Not running inside an Office host. Office SSO is unavailable.'
                };
            }
            const options = {
                allowSignInPrompt: true,
                allowConsentPrompt: true,
                forMSGraphAccess: false
            };

            // Office.auth.getAccessToken returns the raw JWT string
            const token = await Office.auth.getAccessToken(options);
            return { success: true, token: token };
        } catch (err) {
            console.error('[officeInterop] getAccessToken error:', err);
            return {
                success: false,
                errorCode: err.code || 0,
                errorMessage: err.message || 'Unknown error acquiring Office SSO token'
            };
        }
    },

    /**
     * Acquires an access token via MSAL interactive login using the Office Dialog API.
     * Called as a fallback when Office SSO (getOfficeToken) fails — e.g. in Word Online
     * where the Office-on-the-web client is not pre-authorised in the app registration.
     *
     * Opens /auth.html in an Office dialog window; that page runs the MSAL redirect flow
     * and sends the token back via Office.context.ui.messageParent().
     *
     * @returns {Promise<{success: boolean, token?: string, errorCode?: number, errorMessage?: string}>}
     */
    getMsalTokenViaDialog: function () {
        return new Promise(function (resolve) {
            if (typeof Office === 'undefined' || !Office.context || !Office.context.ui ||
                typeof Office.context.ui.displayDialogAsync !== 'function') {
                resolve({ success: false, errorCode: 0, errorMessage: 'Office Dialog API is unavailable.' });
                return;
            }

            var dialogUrl = window.location.origin + '/auth.html';
            Office.context.ui.displayDialogAsync(
                dialogUrl,
                { height: 60, width: 35, promptBeforeOpen: false },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error('[officeInterop] displayDialogAsync failed:', asyncResult.error);
                        resolve({
                            success: false,
                            errorCode: asyncResult.error.code,
                            errorMessage: asyncResult.error.message
                        });
                        return;
                    }

                    var dialog = asyncResult.value;

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                        dialog.close();
                        try {
                            var msg = JSON.parse(arg.message);
                            if (msg.status === 'success' && msg.token) {
                                resolve({ success: true, token: msg.token });
                            } else {
                                resolve({
                                    success: false,
                                    errorCode: 0,
                                    errorMessage: msg.message || 'MSAL sign-in failed.'
                                });
                            }
                        } catch (e) {
                            resolve({ success: false, errorCode: 0, errorMessage: 'Failed to parse auth dialog response.' });
                        }
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
                        // arg.error === 12006 means user closed the dialog
                        resolve({
                            success: false,
                            errorCode: arg.error || 12006,
                            errorMessage: 'Sign-in was cancelled.'
                        });
                    });
                }
            );
        });
    },

    /**
     * Makes a GET request and returns the response as a JSON string.
     * Uses the browser's native fetch() directly, bypassing Blazor's BrowserHttpHandler.
     *
     * @param {string} url  - Relative or absolute URL to fetch
     * @param {string|null} token - Bearer token; omitted when null/undefined
     * @returns {Promise<{ok: boolean, status: number, body: string|null, error: string|null}>}
     */
    fetchJson: async function (url, token) {
        try {
            // Log token header+payload (not signature) to console for diagnostics.
            // Safe: signature is not exposed, and this is only for debugging.
            if (token) {
                try {
                    var parts = token.split('.');
                    if (parts.length === 3) {
                        var payload = JSON.parse(atob(parts[1].replace(/-/g, '+').replace(/_/g, '/')));
                        console.log('[fetchJson] token aud:', payload.aud, 'upn:', payload.upn,
                            'preferred_username:', payload.preferred_username,
                            'oid:', payload.oid, 'iss:', payload.iss);
                    }
                } catch (e) { /* ignore decode errors */ }
            }
            var headers = { 'Accept': 'application/json' };
            if (token) headers['Authorization'] = 'Bearer ' + token;
            var response = await fetch(url, { method: 'GET', headers: headers });
            var body = await response.text();
            // Capture diagnostic header if present
            var authErr = response.headers.get('X-Auth-Error');
            if (authErr) console.warn('[fetchJson] X-Auth-Error:', authErr);
            return { ok: response.ok, status: response.status, body: body, error: null };
        } catch (e) {
            return { ok: false, status: 0, body: null, error: e.message + ' [' + e.name + ']' };
        }
    },

    /**
     * Inserts plain text at the current Word document selection.
     *
     * @param {string} text - The text to insert (replaces current selection)
     * @returns {Promise<{success: boolean, errorMessage?: string}>}
     */
    insertText: async function (text) {
        try {
            await Word.run(async function (context) {
                const range = context.document.getSelection();
                range.insertText(text, Word.InsertLocation.replace);
                await context.sync();
            });
            return { success: true };
        } catch (err) {
            console.error('[officeInterop] insertText error:', err);
            return { success: false, errorMessage: err.message };
        }
    },

    /**
     * Inserts a Base64-encoded image at the current Word document selection
     * as an inline picture using the Word JavaScript API.
     *
     * @param {string} base64 - Raw Base64 encoded image data (no data: URI prefix)
     * @returns {Promise<{success: boolean, errorMessage?: string}>}
     */
    insertImage: async function (base64) {
        try {
            await Word.run(async function (context) {
                const range = context.document.getSelection();
                // insertInlinePictureFromBase64 expects raw Base64, not a data URL
                range.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);
                await context.sync();
            });
            return { success: true };
        } catch (err) {
            console.error('[officeInterop] insertImage error:', err);
            return { success: false, errorMessage: err.message };
        }
    },

    /**
     * Returns true if the page is loaded inside an Office host.
     * Useful to detect when running outside Word (e.g., during local dev in browser).
     *
     * @returns {boolean}
     */
    isInOffice: function () {
        return typeof Office !== 'undefined' &&
               Office.context != null &&
               Office.context.host != null;
    },

    /**
     * Returns information about the current Office host.
     *
     * @returns {{ host: string, platform: string, version: string }}
     */
    getHostInfo: function () {
        if (!this.isInOffice()) {
            return { host: 'None', platform: 'Browser', version: 'N/A' };
        }
        return {
            host: String(Office.context.host),
            platform: String(Office.context.platform),
            version: (Office.context.diagnostics && Office.context.diagnostics.version)
                ? String(Office.context.diagnostics.version)
                : 'N/A'
        };
    }
};
