const { TeamsActivityHandler, TurnContext, MessageFactory } = require("botbuilder");
const axios = require("axios");
const { logMessage, logBotResponse } = require('./logger');

require('dotenv').config();
class TeamsBot extends TeamsActivityHandler {
  constructor(azureDevOpsOrgUrl, personalAccessToken) {
    super();
    this.chatHistories = {};
    this.ticketContext = {};

    this.azureDevOpsOrgUrl = process.env.AZURE_ORG_URL;
    this.personalAccessToken = process.env.AZURE_PAT;
    this.authHeader = {
      'Authorization': `Basic ${Buffer.from(`:${personalAccessToken}`).toString('base64')}`,
      'Accept': 'application/json'
    };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      const conversationId = context.activity.conversation.id;
      if (!this.chatHistories[conversationId]) {
        this.chatHistories[conversationId] = [];
      }

      const sender = context.activity.from.name;
      const messageText = context.activity.text;
      logMessage(sender, messageText);
      // Capture message text correctly
      let txt;
      if (context.activity.conversation.isGroup) {
        txt = TurnContext.removeRecipientMention(context.activity) || context.activity.text;
      } else {
        txt = context.activity.text;
      }

      // Store original text for chat history
      const originalText = txt.replace(/\n|\r/g, "").trim();
      this.chatHistories[conversationId].push(originalText);

      // Convert to lowercase for command processing
      txt = txt.toLowerCase().replace(/\n|\r/g, "").trim();

      // Check for Azure DevOps ticket request
      const workItemId = this.extractWorkItemId(originalText);

      if (workItemId) {
        try {
          const workItemDetails = await this.getWorkItemDetails(workItemId);
          this.ticketContext[conversationId] = workItemDetails;
          logBotResponse(workItemDetails);
          await context.sendActivity(MessageFactory.text(workItemDetails));
        } catch (error) {
          const errorMessage = `Error processing work item: ${error.message}`;
          logBotResponse(errorMessage);
          await context.sendActivity(MessageFactory.text(errorMessage));
        }
      }
      // Check if user requests summary
      else if (txt.includes("@summary")) {
        const summary = await this.generateSummary(this.chatHistories[conversationId]);
        logBotResponse(summary);
        await context.sendActivity(summary);
      }
      else if (txt.endsWith("?") && this.ticketContext[conversationId]) {
        const answer = await this.answerQuestionAboutTicket(originalText, this.ticketContext[conversationId]);
        logBotResponse(answer);
        await context.sendActivity(answer);
      }

      // Regular bot responses
      else {
        const responses = {
          "hello": "Hello! How can I assist you today?",
          "help": "I can help with:\n- Answer questions\n- Read Azure DevOps tickets (just share a ticket URL or ID)\n- Provide conversation summaries (use @summary)",
          "what can you do?": "I can provide information, read Azure DevOps tickets, and summarize conversations. Just ask!"
        };

        const reply = responses[txt] || "I'm not sure how to respond to that. Try asking 'help' to see what I can do!";
        logBotResponse(reply);
        await context.sendActivity(reply);
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const welcomeMessage = "Hi there! I'm a Teams bot. I can answer questions and read Azure DevOps tickets. Just share a ticket URL or ID, or type 'help' for guidance.";
          await context.sendActivity(welcomeMessage);
          logBotResponse(welcomeMessage);
          break;
        }
      }
      console.log("members count", membersAdded.length);
      await next();
    });
  }

  extractWorkItemId(message) {
    try {
      console.log("Analyzing message for work item ID:", message);

      const workItemPatterns = [
        /_workitems\/edit\/(\d+)/i,
        /_workitems\/edit\/(\d+)(?:\/|\?|$)/i,
        /\/workitem[s]?\/(\d+)/i,
        /ticket\s+#?(\d+)/i,
        /#(\d+)\b/
      ];

      for (const pattern of workItemPatterns) {
        const match = message.match(pattern);
        if (match && match[1]) {
          console.log("Found work item ID:", match[1]);
          return match[1];
        }
      }

      // Check if message is just a numeric ID
      if (/^\d+$/.test(message.trim())) {
        console.log("Found numeric ID:", message.trim());
        return message.trim();
      }

      console.log("No work item ID found in message");
      return null;
    } catch (error) {
      console.error('Error extracting work item ID:', error);
      return null;
    }
  }

  async getWorkItemDetails(workItemId) {
    try {
      const baseUrl = "https://dev.azure.com/ShorthillsPM";
      const apiUrl = `${baseUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0`;

      console.log(`Fetching work item #${workItemId} from: ${apiUrl}`);

      // Match exactly how curl formats the auth header
      const authToken = Buffer.from(`:${this.personalAccessToken}`).toString('base64');
      const authHeader = {
        'Authorization': `Basic ${authToken}`,
        'Accept': 'application/json'
      };

      // Make the request
      const response = await axios.get(apiUrl, {
        headers: authHeader
      });

      console.log(`Response status: ${response.status}`);

      const history = await this.getWorkItemHistory(workItemId);
      return this.formatWorkItem(response.data, history);
    } catch (error) {
      console.error("Error fetching work item:", error);

      // Provide detailed error information
      if (error.response) {
        return `Error ${error.response.status}: Cannot retrieve work item #${workItemId}. ${error.response.data?.message || ''}`;
      } else if (error.request) {
        return "Network error: No response received from Azure DevOps API.";
      } else {
        return `Error: ${error.message}`;
      }
    }
  }

  async getWorkItemHistory(workItemId) {
    try {
      const baseUrl = "https://dev.azure.com/ShorthillsPM";
      const apiUrl = `${baseUrl}/_apis/wit/workitems/${workItemId}/updates?api-version=6.0`;

      console.log(`Fetching history for work item #${workItemId}`);

      const authToken = Buffer.from(`:${this.personalAccessToken}`).toString('base64');
      const authHeader = {
        'Authorization': `Basic ${authToken}`,
        'Accept': 'application/json'
      };

      const response = await axios.get(apiUrl, {
        headers: authHeader
      });

      console.log(`History response status: ${response.status}`);

      return response.data.value || [];
    } catch (error) {
      console.error("Error fetching work item history:", error);
      return [];
    }
  }

  formatWorkItem(workItem, history = []) {
    const fields = workItem.fields || {};
    let formattedInfo = '📄 **WORK ITEM DETAILS**\n\n';

    // Basic information
    formattedInfo += `**ID**: ${workItem.id}\n`;
    formattedInfo += `**Title**: ${fields['System.Title'] || 'N/A'}\n`;
    formattedInfo += `**State**: ${fields['System.State'] || 'N/A'}\n`;
    formattedInfo += `**Type**: ${fields['System.WorkItemType'] || 'N/A'}\n`;

    // People
    formattedInfo += `**Created By**: ${fields['System.CreatedBy']?.displayName || 'N/A'}\n`;
    formattedInfo += `**Assigned To**: ${fields['System.AssignedTo']?.displayName || 'N/A'}\n`;

    // Dates
    const createdDate = fields['System.CreatedDate'] ? new Date(fields['System.CreatedDate']).toLocaleString() : 'N/A';
    formattedInfo += `**Created Date**: ${createdDate}\n`;

    if (fields['System.ChangedDate']) {
      const changedDate = new Date(fields['System.ChangedDate']).toLocaleString();
      formattedInfo += `**Last Updated**: ${changedDate}\n`;
    }

    // Priority and effort
    if (fields['Microsoft.VSTS.Common.Priority']) {
      formattedInfo += `**Priority**: ${fields['Microsoft.VSTS.Common.Priority']}\n`;
    }

    if (fields['Microsoft.VSTS.Scheduling.StoryPoints']) {
      formattedInfo += `**Story Points**: ${fields['Microsoft.VSTS.Scheduling.StoryPoints']}\n`;
    }

    // Add iteration and area path if available
    if (fields['System.IterationPath']) {
      formattedInfo += `**Iteration Path**: ${fields['System.IterationPath']}\n`;
    }

    if (fields['System.AreaPath']) {
      formattedInfo += `**Area Path**: ${fields['System.AreaPath']}\n`;
    }

    // Description (may contain HTML)
    formattedInfo += '\n**Description**:\n';
    formattedInfo += fields['System.Description'] ?
      this.stripHtml(fields['System.Description']) : 'No description provided.';

    // Acceptance Criteria if available
    if (fields['Microsoft.VSTS.Common.AcceptanceCriteria']) {
      formattedInfo += '\n\n**Acceptance Criteria**:\n';
      formattedInfo += this.stripHtml(fields['Microsoft.VSTS.Common.AcceptanceCriteria']);
    }

    // Links to related work items
    if (workItem.relations && workItem.relations.length > 0) {
      const relatedItems = workItem.relations.filter(rel =>
        rel.rel === 'System.LinkTypes.Related' ||
        rel.rel === 'System.LinkTypes.Child' ||
        rel.rel === 'System.LinkTypes.Parent'
      );

      if (relatedItems.length > 0) {
        formattedInfo += '\n\n**Related Items**:\n';
        relatedItems.forEach(item => {
          const relationType = item.rel.split('.').pop();
          const itemUrl = item.url;
          const itemId = itemUrl.substring(itemUrl.lastIndexOf('/') + 1);
          formattedInfo += `- ${relationType}: #${itemId}\n`;
        });
      }
    }

    let hasDiscussion = false;
    formattedInfo += '\n\n**Discussion**:\n';

    if (history && history.length > 0) {
      const discussionEntries = history.filter(update =>
        update.fields &&
        (update.fields['System.History'] ||
          update.fields['System.CommentCount'])
      );

      if (discussionEntries.length > 0) {
        hasDiscussion = true;
        discussionEntries.forEach(entry => {
          if (entry.fields['System.History']) {
            const commentDate = new Date(entry.revisedDate).toLocaleString();
            formattedInfo += `\n📝 **${entry.revisedBy?.displayName || 'Unknown'}** (${commentDate}):\n`;
            formattedInfo += `${this.stripHtml(entry.fields['System.History'].newValue)}\n`;
          }
        });
      }
    }

    if (!hasDiscussion && fields['System.Description']) {
      const description = fields['System.Description'];

      if (description.includes('<div class="comment">') ||
        description.includes('<div class="discussion">') ||
        description.includes('Posted by:') ||
        description.match(/On \d{1,2}\/\d{1,2}\/\d{2,4}.*wrote:/i)) {

        formattedInfo += `\nThe discussion appears to be embedded in the work item description. Here's the full description text:\n\n`;
        formattedInfo += this.stripHtml(description);
        hasDiscussion = true;
      }
    }

    if (!hasDiscussion) {
      formattedInfo += 'No comments or discussion found for this work item.';
    }

    return formattedInfo;
  }

  stripHtml(html) {
    if (!html) return '';

    return html
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n\n')
      .replace(/<li>/gi, '- ')
      .replace(/<\/li>/gi, '\n')
      .replace(/<\/h[1-6]>/gi, '\n')
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .trim();
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

    const apiKey = process.env.GEMINI_API_KEY;
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
  async answerQuestionAboutTicket(question, ticketContent) {
    const prompt = `
  You are an assistant helping users understand Azure DevOps tickets. Based on the following work item data:

  """ 
  ${ticketContent}
  """

  Answer the following question clearly and accurately:

  Q: ${question}
  A: `;

    const apiKey = process.env.GEMINI_API_KEY;
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

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
      return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "Sorry, I couldn't find an answer.";
    } catch (error) {
      console.error("QnA Error:", error);
      return "Error answering the question.";
    }
  }

}

module.exports.TeamsBot = TeamsBot;
