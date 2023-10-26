/* global window, clearInterval, setInterval, fetch, console, crypto */

function bytesToBase64Url(bytes) {
  return window
    .btoa(String.fromCharCode(...bytes))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/[=]+$/g, "");
}

async function getOauthParams() {
  const state = bytesToBase64Url(crypto.getRandomValues(new Uint8Array(16)));
  const codeVerifier = bytesToBase64Url(crypto.getRandomValues(new Uint8Array(32)));
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(codeVerifier));
  const codeChallenge = bytesToBase64Url(new Uint8Array(hash));
  return { state, codeChallenge, codeVerifier };
}

async function getMicrosoftAccessToken(clientId) {
  const redirectUri = "https://localhost:3010/taskpane.html";
  const { codeChallenge, codeVerifier, state } = await getOauthParams();
  // Base Authentication URL for OneDrive and SharePoint.
  const url = new URL("https://login.microsoftonline.com/common/oauth2/v2.0/authorize");
  url.search = new URLSearchParams({
    client_id: clientId,
    scope: "openid user.read offline_access",
    redirect_uri: redirectUri,
    response_type: "code",
    response_mode: "query",
    code_challenge: codeChallenge,
    code_challenge_method: "S256",
    state,
  }).toString();

  const popupWindow = window.open(url, "authWindow", "popup");

  await new Promise((resolve, reject) => {
    // Monitor the popup window to detect if it was closed. Yes, this is janky,
    // but is the most cross-browser compatible approach.
    // See: https://stackoverflow.com/q/3291712
    const intervalId = setInterval(() => {
      if (popupWindow.closed) {
        cleanup();
        reject(new Error("authorization cancelled"));
      }
    }, 1000);

    window.addEventListener("message", onMessage);
    function onMessage(event) {
      cleanup();
      resolve(event.message);
    }
    function cleanup() {
      window.removeEventListener("message", onMessage);
      clearInterval(intervalId);
      window.sessionStorage.removeItem("isTokenRequest");
    }
  });

  // Extract OAuth callback parameters from the popup window URL.
  const oauthParams = Object.fromEntries(new URLSearchParams(popupWindow.location.search));
  console.log(oauthParams, "oauthParams");
  if (oauthParams.state !== state) {
    throw new Error("OAuth state mismatch");
  }
  if (!oauthParams.code || oauthParams.error) {
    throw new Error(`OAuth error: ${oauthParams.error_description}`);
  }
  const response = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    body: new URLSearchParams({
      grant_type: "authorization_code",
      client_id: clientId,
      code: oauthParams.code,
      code_verifier: codeVerifier,
      redirect_uri: redirectUri,
    }),
  });

  if (!response.ok) {
    throw new Error(`OAuth token error: ${await response.text()}`);
  }
  const tokenResult = await response.json();
  // Prevent an almost-expired token from being used again by subtracting 5
  // mins from its approximate expiration time.
  return {
    accessToken: tokenResult.access_token,
    refreshToken: tokenResult.refresh_token,
    expiresAt: Date.now() + tokenResult.expires_in * 1_000 - 5 * 60_000,
  };
}

async function refreshAccessToken(clientId, refreshToken) {
  const response = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    body: new URLSearchParams({
      grant_type: "refresh_token",
      client_id: clientId,
      refresh_token: refreshToken,
      redirect_uri: "https://localhost:3010/taskpane.html",
    }),
  });

  if (response.ok) {
    const tokenResult = await response.json();
    console.log(tokenResult);
    return {
      accessToken: tokenResult.access_token,
      refreshToken: tokenResult.refresh_token,
      expiresAt: Date.now() + tokenResult.expires_in * 1_000 - 5 * 60_000,
    };
  } else {
    console.log(await response.json());
  }
}

export { getMicrosoftAccessToken, refreshAccessToken };
