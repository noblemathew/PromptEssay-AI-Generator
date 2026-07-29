<script>
const FLOW_URL = "PASTE_YOUR_ADVISOR_FLOW_URL_HERE";

const msalConfig = {
  auth: {
    clientId: "PASTE_CLIENT_ID_HERE",
    authority: "PASTE_TENANT_ID_HERE",
    redirectUri: window.location.origin
  }
};

const loginRequest = { scopes: [".default"] };

let msalApp;

async function initializeMsal() {
  if (!msalApp) {
    msalApp = new msal.PublicClientApplication(msalConfig);
    await msalApp.initialize();
  }
  return msalApp;
}

async function signInUser() {
  const app = await initializeMsal();
  const loginResponse = await app.loginPopup(loginRequest);
  return loginResponse.account;
}

async function getAccessToken() {
  const app = await initializeMsal();
  let accounts = app.getAllAccounts();

  if (accounts.length === 0) {
    await signInUser();
    accounts = app.getAllAccounts();
  }

  const tokenRequest = {
    scopes: loginRequest.scopes,
    account: accounts[0]
  };

  try {
    const tokenResponse = await app.acquireTokenSilent(tokenRequest);
    return tokenResponse.accessToken;
  } catch (err) {
    console.warn("Silent token failed, falling back to popup:", err);
    const tokenResponse = await app.acquireTokenPopup(tokenRequest);
    return tokenResponse.accessToken;
  }
}

async function triggerCopilot() {
  const promptInput = document.getElementById("promptInput");
  const promptText = promptInput.value.trim();

  if (!promptText) {
    alert("Please enter a prompt.");
    return;
  }

  addChatBubble(promptText, "user");
  addTypingBubble();

  try {
    const accessToken = await getAccessToken();
    const app = await initializeMsal();
    const account = app.getAllAccounts()[0];

    const response = await fetch(FLOW_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${accessToken}`
      },
      body: JSON.stringify({
        userPrompt: promptText,
        userName: account?.name || "Unknown User"
      })
    });

    removeTypingBubble();

    if (!response.ok) {
      throw new Error(`${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    const answer = result.answer || result.response || result.message || JSON.stringify(result);
    lastResponseText = answer;
    addChatBubble(answer, "agent");

  } catch (error) {
    removeTypingBubble();
    const message = `Error: ${error.message}`;
    lastResponseText = message;
    addChatBubble(message, "agent");
  }
}

function clearChat() {
  document.getElementById("promptInput").value = "";
  lastResponseText = "";
  const responsePanel = document.getElementById("responsePanel");
  responsePanel.innerHTML = `
    <div id="emptyResponse" class="empty-response">
      No response received yet.
    </div>`;
}

function copyResponse() {
  if (!lastResponseText) {
    alert("No response available to copy.");
    return;
  }
  navigator.clipboard.writeText(lastResponseText)
    .then(() => {
      alert("Response copied to clipboard.");
    })
    .catch(() => {
      alert("Unable to copy response. Please copy manually.");
    });
}
</script>
</body>
</html>
