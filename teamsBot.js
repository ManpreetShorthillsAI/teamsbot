const { TeamsActivityHandler, TurnContext, MessageFactory } = require("botbuilder");
const axios = require("axios");

require('dotenv').config();
class TeamsBot extends TeamsActivityHandler {
  constructor(azureDevOpsOrgUrl, personalAccessToken) {
    super();
    this.chatHistories = {};
    
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
          await context.sendActivity(MessageFactory.text(workItemDetails));
        } catch (error) {
          await context.sendActivity(MessageFactory.text(
            `Error processing work item: ${error.message}`
          ));
        }
      }
      // Check if user requests summary
      else if (txt.includes("@summary")) {
        const summary = await this.generateSummary(this.chatHistories[conversationId]);
        await context.sendActivity(summary);
      } 
      // Regular bot responses
      else {
        const responses = {
          "hello": "Hello! How can I assist you today?",
          "help": "I can help with:\n- Answer questions\n- Read Azure DevOps tickets (just share a ticket URL or ID)\n- Provide conversation summaries (use @summary)",
          "what can you do?": "I can provide information, read Azure DevOps tickets, and summarize conversations. Just ask!"
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
            "Hi there! I'm a Teams bot. I can answer questions and read Azure DevOps tickets. Just share a ticket URL or ID, or type 'help' for guidance."
          );
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
      
      // If we get here, we successfully retrieved the work item
      return this.formatWorkItem(response.data);
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

  formatWorkItem(workItem) {
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
}

module.exports.TeamsBot = TeamsBot;
