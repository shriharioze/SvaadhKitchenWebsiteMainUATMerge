document.addEventListener("DOMContentLoaded", () => {
  const widget = document.getElementById("chat-widget");
  widget.innerHTML = `
    <button id="chat-toggle" class="chat-button">💬</button>
    <div class="chat-box" id="chat-box">
      <div class="chat-header">Svaadh Kitchen 🧡</div>
      <div class="chat-messages" id="chat-messages"></div>
      <div class="chat-input">
        <textarea id="user-input" placeholder="Type your message..."></textarea>
        <button id="send-btn">Send</button>
      </div>
    </div>
  `;

  const toggle = document.getElementById("chat-toggle");
  const chatBox = document.getElementById("chat-box");
  const sendBtn = document.getElementById("send-btn");
  const userInput = document.getElementById("user-input");
  const messages = document.getElementById("chat-messages");

  // Auto-expand when loaded
  chatBox.style.display = "flex";

  toggle.addEventListener("click", () => {
    chatBox.style.display = chatBox.style.display === "none" ? "flex" : "none";
  });

  function appendMessage(text, sender) {
    const div = document.createElement("div");
    div.classList.add("message", sender);
    div.innerHTML = text.replace(
      /(https?:\/\/[^\s]+|[\w.-]+@[\w.-]+|\+?\d{7,})/g,
      '<a href="$1" target="_blank" style="color:#f47c3c;text-decoration:none;">$1</a>'
    );
    messages.appendChild(div);
    messages.scrollTop = messages.scrollHeight;
  }

  async function sendMessage() {
    const text = userInput.value.trim();
    if (!text) return;
    appendMessage(text, "user");
    userInput.value = "";

    const res = await fetch("https://svaadhkitchenwebsite.onrender.com/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message: text }),
    });
    const data = await res.json();
    appendMessage(data.reply, "bot");
  }

  sendBtn.addEventListener("click", sendMessage);
  userInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
});
