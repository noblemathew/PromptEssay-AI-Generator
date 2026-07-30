(async function initShell() {
  const me = await requireSession();
  if (!me) return;
  document.getElementById("sidebarAvatar").textContent = initials(me.display_name || me.email);
  document.getElementById("sidebarName").textContent = me.display_name || me.email;

  const params = new URLSearchParams(window.location.search);
  const projectId = params.get("project");
  document.getElementById("backToProjectLink").href =
    projectId ? `../../project.html?id=${projectId}` : "../../dashboard.html";

  document.getElementById("logoutLink").addEventListener("click", (e) => {
    e.preventDefault();
    logout("../../index.html");
  });
})();

let lastResponseText = "";

function insertSampleQuestion(question) {
  const promptInput = document.getElementById("promptInput");
  promptInput.value = question;
  promptInput.focus();
}

function getCurrentTime() {
  const now = new Date();
  return now.toLocaleString([], {
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function removeEmptyState() {
  const emptyResponse = document.getElementById("emptyResponse");
  if (emptyResponse) {
    emptyResponse.remove();
  }
}

function addChatBubble(message, type) {
  removeEmptyState();

  const responsePanel = document.getElementById("responsePanel");

  const bubble = document.createElement("div");
  bubble.className = type === "user"
    ? "chat-bubble user-bubble"
    : "chat-bubble agent-bubble";

  const messageText = document.createElement("div");
  messageText.innerText = message;

  const timestamp = document.createElement("span");
  timestamp.className = "timestamp";
  timestamp.innerText = getCurrentTime();

  bubble.appendChild(messageText);
  bubble.appendChild(timestamp);
  responsePanel.appendChild(bubble);

  responsePanel.scrollTop = responsePanel.scrollHeight;

  return bubble;
}

function addTypingBubble() {
  removeEmptyState();

  const responsePanel = document.getElementById("responsePanel");

  const bubble = document.createElement("div");
  bubble.className = "chat-bubble agent-bubble";
  bubble.id = "typingBubble";

  bubble.innerHTML = `
    <div class="typing" aria-label="Advisor is thinking">
      <span></span>
      <span></span>
      <span></span>
      <span style="width:auto;height:auto;background:none;color:#52687a;font-size:13px;margin-left:6px;">
        Advisor is thinking...
      </span>
    </div>
  `;

  responsePanel.appendChild(bubble);
  responsePanel.scrollTop = responsePanel.scrollHeight;
}

function removeTypingBubble() {
  const typingBubble = document.getElementById("typingBubble");
  if (typingBubble) {
    typingBubble.remove();
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
    const response = await fetch(PROXY_URL + "/api/advisor-chat", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      credentials: "include",
      body: JSON.stringify({ userPrompt: promptText })
    });

    removeTypingBubble();

    const rawText = await response.text();

    if (!response.ok) {
      throw new Error(`${response.status} ${response.statusText} — ${rawText}`);
    }

    let result;
    try {
      result = JSON.parse(rawText);
    } catch {
      result = { raw: rawText };
    }

    const answer = result.answer || result.response || result.message || result.raw || JSON.stringify(result);
    lastResponseText = answer;
    addChatBubble(answer, "agent");

  } catch (error) {
    console.error("triggerCopilot error:", error);
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
