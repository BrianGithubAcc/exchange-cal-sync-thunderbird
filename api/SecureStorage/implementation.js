/* global ExtensionCommon, Components, Services */

"use strict";

var SecureStorage = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const ORIGIN = "chrome://exchange-calendar-sync";
    const REALM_PREFIX = "Exchange Calendar Sync";
    const CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753";
    const SCOPE = "https://outlook.office.com/EWS.AccessAsUser.All offline_access";
    const AUTHORITY = "https://login.microsoftonline.com/common/oauth2/v2.0";

    function realmFor(key) {
      if (typeof key !== "string" || !key || key.length > 300) {
        throw new Error("Invalid secure-storage key");
      }
      return REALM_PREFIX + ":" + key;
    }

    async function findLogin(key) {
      const realm = realmFor(key);
      const logins = await Services.logins.searchLoginsAsync({
        origin: ORIGIN,
        httpRealm: realm,
      });
      return logins.length > 0 ? logins[0] : null;
    }

    function newLogin(key, value) {
      const login = Components.classes[
        "@mozilla.org/login-manager/loginInfo;1"
      ].createInstance(Components.interfaces.nsILoginInfo);
      login.init(ORIGIN, null, realmFor(key), key, value, "", "");
      return login;
    }

    function safeOAuthValue(value, name, maxLength) {
      if (typeof value !== "string" || !value || value.length > maxLength ||
          /[\r\n\x00]/.test(value)) {
        throw new Error("Invalid OAuth " + name);
      }
      return value;
    }

    function formEncode(fields) {
      return Object.keys(fields).map(function (key) {
        return encodeURIComponent(key) + "=" + encodeURIComponent(fields[key]);
      }).join("&");
    }

    function postOAuth(endpoint, fields) {
      // This runs in the privileged Experiment parent, not in the
      // moz-extension background page. Use a system-principal HTTP channel so
      // the request has no moz-extension Origin and is not subject to the
      // browser-page CORS path. Return the body as text so the Experiment
      // schema never has to clone an arbitrary JSON object.
      return new Promise(function (resolve) {
        let channel;
        let timer;
        let settled = false;

        function finish(status, body) {
          if (settled) return;
          settled = true;
          if (timer) timer.cancel();
          resolve({ status: status, body: body || "" });
        }

        function failure(error, status, description) {
          finish(status || 0, JSON.stringify({
            error: error,
            error_description: description || (error === "timeout"
              ? "OAuth request timed out"
              : "OAuth network request failed"),
          }));
        }

        const listener = {
          chunks: [],

          QueryInterface: ChromeUtils.generateQI([
            Components.interfaces.nsIStreamListener,
            Components.interfaces.nsIRequestObserver,
          ]),

          onStartRequest() {},

          onDataAvailable(request, inputStream, offset, count) {
            const binary = Components.classes["@mozilla.org/binaryinputstream;1"]
              .createInstance(Components.interfaces.nsIBinaryInputStream);
            binary.setInputStream(inputStream);
            const bytes = binary.readByteArray(count);
            this.chunks.push(String.fromCharCode.apply(null, bytes));
          },

          onStopRequest(request, requestStatus) {
            if (requestStatus !== Components.results.NS_OK) {
              failure("network_error", 0, "OAuth network request failed (0x" +
                requestStatus.toString(16) + ")");
              return;
            }

            let status = 0;
            try {
              status = request.QueryInterface(Components.interfaces.nsIHttpChannel)
                .responseStatus;
            } catch (e) { /* use status 0 if no HTTP response exists */ }

            const binaryBody = this.chunks.join("");
            let body = binaryBody;
            try {
              body = decodeURIComponent(escape(binaryBody));
            } catch (e) { /* response was already usable as a byte string */ }
            finish(status, body);
          },
        };

        try {
          const uri = Services.io.newURI(AUTHORITY + endpoint);
          const systemPrincipal = Services.scriptSecurityManager.getSystemPrincipal();
          channel = Services.io.newChannelFromURI(
            uri,
            null,
            systemPrincipal,
            null,
            Components.interfaces.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
            Components.interfaces.nsIContentPolicy.TYPE_OTHER
          );

          const httpChannel = channel.QueryInterface(Components.interfaces.nsIHttpChannel);
          httpChannel.setRequestHeader("Accept", "application/json", false);

          const requestBody = formEncode(fields);
          const stream = Components.classes["@mozilla.org/io/string-input-stream;1"]
            .createInstance(Components.interfaces.nsIStringInputStream);
          stream.setUTF8Data(requestBody);

          const uploadChannel = channel.QueryInterface(Components.interfaces.nsIUploadChannel2);
          uploadChannel.explicitSetUploadStream(
            stream,
            "application/x-www-form-urlencoded",
            requestBody.length,
            "POST",
            false
          );

          timer = Components.classes["@mozilla.org/timer;1"]
            .createInstance(Components.interfaces.nsITimer);
          timer.initWithCallback({
            notify() {
              try { channel.cancel(Components.results.NS_ERROR_NET_TIMEOUT); } catch (e) {}
              failure("timeout");
            },
          }, 30000, Components.interfaces.nsITimer.TYPE_ONE_SHOT);

          // Necko requires a load group for asynchronous channels in the
          // parent process. Without one, asyncOpen can fail with status 0.
          const loadGroup = Components.classes["@mozilla.org/network/load-group;1"]
            .createInstance(Components.interfaces.nsILoadGroup);
          channel.loadGroup = loadGroup;
          channel.asyncOpen(listener);
        } catch (e) {
          failure("client_error", 0, String(e && e.message || e));
        }
      });
    }

    return {
      SecureStorage: {
        async get(key) {
          const login = await findLogin(key);
          return login
            ? { found: true, value: login.password }
            : { found: false };
        },

        async set(key, value) {
          const oldLogin = await findLogin(key);
          if (oldLogin) {
            await Services.logins.removeLoginAsync(oldLogin);
          }
          await Services.logins.addLoginAsync(newLogin(key, value));
          return true;
        },

        async remove(key) {
          const login = await findLogin(key);
          if (login) {
            await Services.logins.removeLoginAsync(login);
          }
          return true;
        },

        async startDeviceCode() {
          return await postOAuth("/devicecode", {
            client_id: CLIENT_ID,
            scope: SCOPE,
          });
        },

        async pollDeviceCode(deviceCode) {
          return await postOAuth("/token", {
            client_id: CLIENT_ID,
            grant_type: "urn:ietf:params:oauth:grant-type:device_code",
            device_code: safeOAuthValue(deviceCode, "device code", 4096),
          });
        },

        async refreshAccessToken(refreshToken) {
          return await postOAuth("/token", {
            client_id: CLIENT_ID,
            grant_type: "refresh_token",
            refresh_token: safeOAuthValue(refreshToken, "refresh token", 20000),
            scope: SCOPE,
          });
        },
      },
    };
  }
};
