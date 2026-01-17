document.addEventListener("DOMContentLoaded", () => {
  // Backend URL - change this for different environments
  const BACKEND_URL = window.location.hostname === 'localhost' 
    ? "http://localhost:8001" 
    : "https://svaadhkitchenwebsite.onrender.com";
  
  const widget = document.getElementById("chat-widget");
  widget.innerHTML = `
    <button id="chat-toggle" class="chat-button">💬</button>
    <div class="chat-box" id="chat-box">
      <div class="chat-header">Svaadh Kitchen 🧡</div>
      <div class="chat-messages" id="chat-messages"></div>
      <div class="quick-replies" id="quick-replies"></div>
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
  const quickReplies = document.getElementById("quick-replies");

  // Auto-expand when loaded
  chatBox.style.display = "flex";

  toggle.addEventListener("click", () => {
    chatBox.style.display = chatBox.style.display === "none" ? "flex" : "none";
  });

  function showQuickReplies() {
    const quickReplyButtons = [
      { text: "🍛 Today's Menu", message: "What's today's menu?" },
      { text: "⏰ Order Timings", message: "What are your order timings?" },
      { text: "📍 Delivery Areas", message: "Which areas do you deliver to?" },
      { text: "� Full Menu", message: "Show me the full menu with prices", isMenu: true },
      { text: "� Place Order", message: "I want to place an order", isOrder: true }
    ];

    quickReplies.innerHTML = '';
    quickReplyButtons.forEach(button => {
      const btn = document.createElement("button");
      btn.className = "quick-reply-btn";
      btn.textContent = button.text;
      btn.onclick = () => {
        if (button.isOrder) {
          window.open("https://tally.so/r/w4WKZd", "_blank");
          appendMessage("Opening order form...", "bot");
        } else if (button.isMenu) {
          showFullMenu();
        } else {
          userInput.value = button.message;
          sendMessage();
        }
      };
      quickReplies.appendChild(btn);
    });
  }

  function showFullMenu() {
    const menuHTML = `
      <div class="menu-display">
        <h4>🍛 Make Your Own Meal - Menu & Prices</h4>
        <div class="menu-category">
          <strong>Vegetable Curries:</strong>
          <div class="menu-item">• Dry Sabji Mini (100ml) - ₹20</div>
          <div class="menu-item">• Curry Sabji Mini (100ml) - ₹20</div>
          <div class="menu-item">• Dry Sabji (250ml) - ₹45</div>
          <div class="menu-item">• Curry Sabji (250ml) - ₹45</div>
        </div>
        <div class="menu-category">
          <strong>Basics:</strong>
          <div class="menu-item">• Dal (200ml) - ₹20</div>
          <div class="menu-item">• Rice (100gms) - ₹10</div>
          <div class="menu-item">• Salad (40gms) - ₹5</div>
          <div class="menu-item">• Curd (50gms) - ₹10</div>
        </div>
        <div class="menu-category">
          <strong>Breads:</strong>
          <div class="menu-item">• Chapati / Phulka (3 pcs) - ₹20</div>
          <div class="menu-item">• Ghee Phulka (3 pcs) - ₹30</div>
          <div class="menu-item">• Jowar/Bajra Bhakri (1 pc) - ₹20</div>
        </div>
        <div class="menu-category">
          <strong>Special Items:</strong>
          <div class="menu-item">• Special Thali (Complete Meal) - ₹120</div>
          <div class="menu-item">• Weekend Special - ₹150</div>
        </div>
        <div class="menu-note">
          💡 <em>Mix and match items to create your perfect meal!</em><br>
          📞 <strong>Order via WhatsApp or click "Place Order" button</strong>
        </div>
      </div>
    `;
    
    // Create a temporary div to render HTML properly
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = menuHTML;
    appendMessage(tempDiv.innerHTML, "bot");
  }

  // Load chat history and show initial greeting
  loadChatHistory();
  
  // Show greeting only if chat is empty
  const chatHistory = JSON.parse(localStorage.getItem('svaadhChatHistory') || '[]');
  if (chatHistory.length === 0) {
    setTimeout(() => {
      appendMessage("Hello! 👋 Welcome to Svaadh Kitchen! How can I help you today?", "bot");
      showQuickReplies();
    }, 500);
  } else {
    showQuickReplies();
  }

  function saveMessage(text, sender) {
    const chatHistory = JSON.parse(localStorage.getItem('svaadhChatHistory') || '[]');
    chatHistory.push({ text, sender, timestamp: new Date().toISOString() });
    // Keep only last 50 messages
    if (chatHistory.length > 50) {
      chatHistory.shift();
    }
    localStorage.setItem('svaadhChatHistory', JSON.stringify(chatHistory));
  }

  function loadChatHistory() {
    const chatHistory = JSON.parse(localStorage.getItem('svaadhChatHistory') || '[]');
    chatHistory.forEach(msg => {
      appendMessage(msg.text, msg.sender);
    });
  }

  function appendMessage(text, sender) {
    const div = document.createElement("div");
    div.classList.add("message", sender);
    div.innerHTML = text.replace(
      /(https?:\/\/[^\s]+|[\w.-]+@[\w.-]+|\+?\d{7,})/g,
      '<a href="$1" target="_blank" style="color:#f47c3c;text-decoration:none;">$1</a>'
    );
    messages.appendChild(div);
    messages.scrollTop = messages.scrollHeight;
    
    // Save message to localStorage (skip for loading messages)
    if (!sender.includes('loading')) {
      saveMessage(text, sender);
    }
  }

  function showLoading() {
    const div = document.createElement("div");
    div.classList.add("message", "bot", "loading");
    div.innerHTML = '<div class="typing-indicator"><span>Svaadh Kitchen is typing</span><div class="typing-dots"><span></span><span></span><span></span></div></div>';
    messages.appendChild(div);
    messages.scrollTop = messages.scrollHeight;
    return div;
  }

  async function sendMessage() {
  const text = userInput.value.trim();
  if (!text) return;

  // Append user's message to chat
  appendMessage(text, "user");
  userInput.value = "";

  // Show loading indicator
  const loadingDiv = showLoading();

  // 🔹 Google Analytics Event: User Message Sent
  if (typeof gtag === "function") {
    gtag("event", "chat_message", {
      event_category: "Chatbot",
      event_label: text, // captures user's actual message
      value: 1,
    });
  }

  try {
    const res = await fetch(`${BACKEND_URL}/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message: text }),
    });

    const data = await res.json();

    // Remove loading indicator
    loadingDiv.remove();

    // Append chatbot's reply to chat
    appendMessage(data.reply, "bot");

    // 🔹 Google Analytics Event: Bot Reply Sent
    if (typeof gtag === "function") {
      gtag("event", "chatbot_reply", {
        event_category: "Chatbot",
        event_label: data.reply, // captures bot's response
        value: 1,
      });
    }
  } catch (error) {
    console.error("Chat error:", error);
    // Remove loading indicator
    loadingDiv.remove();
    appendMessage("Sorry, I'm having trouble connecting right now. Please try again in a 15-30 seconds or call us directly at +919930748908.", "bot");
  }
}


  sendBtn.addEventListener("click", sendMessage);
  userInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  // FAQ toggle functionality
  function toggleFAQ(button) {
    const answer = button.nextElementSibling;
    const isOpen = answer.style.display === 'block';
    
    if (isOpen) {
      answer.style.display = 'none';
      answer.style.maxHeight = '0';
      answer.style.padding = '0 20px';
      button.querySelector('span').textContent = '+';
    } else {
      answer.style.display = 'block';
      answer.style.maxHeight = '500px';
      answer.style.padding = '20px';
      button.querySelector('span').textContent = '-';
    }
  }

  // Make functions global
  window.toggleFAQ = toggleFAQ;
});
