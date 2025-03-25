const { TeamsActivityHandler, TurnContext } = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.chatHistories = {}; // Store chat history per chat ID

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      
      console.log("______________")
      console.log(context)
      console.log("________________")

      const conversationId = context.activity.conversation.id; // Unique ID for each chat
      if (!this.chatHistories[conversationId]) {
        this.chatHistories[conversationId] = []; // Initialize if not exists
      }

      // Capture message text correctly
      let txt;
      if (context.activity.conversation.isGroup) {
        // Remove bot mention in group chats but still capture the full message
        txt = TurnContext.removeRecipientMention(context.activity) || context.activity.text;
      } else {
        // Normal messages in 1:1 chat
        txt = context.activity.text;
      }

      txt = txt.toLowerCase().replace(/\n|\r/g, "").trim();
      this.chatHistories[conversationId].push(txt); // Store chat-specific history

      // Check if user requests summary
      if (txt.includes("@summary")) {
        const summary = await this.generateSummary(this.chatHistories[conversationId]);
        await context.sendActivity(summary);
      } else {
        const responses = {
          "hello": "Hello! How can I assist you today?",
          "help": "I can answer your questions! Try asking 'What can you do?'.",
          "what can you do?": "I can provide information and assist with common queries. Just ask!"
        };

        const reply = responses[txt] || "I'm not sure how to respond to that. Try asking 'help' to see what I can do!";
        await context.sendActivity(reply);
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            "Hi there! I'm a Teams bot. Ask me a question like 'What can you do?' or type 'help' for guidance."
          );
          break;
        }
      }
      console.log("members count", membersAdded.length);
      await next();
    });
  }

  async generateSummary(chatHistory) {
    const messagesText = chatHistory.join("\n");

    const prompt = `
        You are a professional summarizer. Provide a concise, clear, and objective summary
        of the following conversation and messages from today:
 
        \`\`\`
        ${messagesText}
        \`\`\`
 
        Key requirements for the summary:
        1. Capture the main topics and key points discussed
        2. Identify any important decisions or action items
        3. Be objective and neutral in tone
        4. Limit the summary to 300-500 words
        5. Use clear, professional language
    `;

    const apiKey = "";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

    console.log("Chat History:", chatHistory);

    const requestBody = {
      contents: [{ parts: [{ text: prompt }] }]
    };

    try {
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(requestBody)
      });

      const data = await response.json();
      console.log("Summary API Response:", data);
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "No summary available.";
    } catch (error) {
      console.error("Error generating summary:", error);
      return "Error generating summary.";
    }
  }
}

module.exports.TeamsBot = TeamsBot;
