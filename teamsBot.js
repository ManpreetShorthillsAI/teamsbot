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
      else if (originalText.startsWith("@comment")) {
        const commentText = originalText.substring("@comment".length).trim();

        if (!this.ticketContext[conversationId]) {
          await context.sendActivity(MessageFactory.text(
            "Please share a ticket first before adding a comment."
          ));
        } else {
          try {
            const workItemId = this.extractWorkItemIdFromContext(this.ticketContext[conversationId]);
            console.log("Extracted work item ID from context:", workItemId);


            if (workItemId) {
              await this.addCommentToWorkItem(workItemId, commentText);
              console.log("Comment added successfully to work item:", workItemId);

              await context.sendActivity(MessageFactory.text(
                `✅ Comment added to work item #${workItemId} successfully!`
              ));
            } else {
              await context.sendActivity(MessageFactory.text(
                "Could not determine which ticket to comment on. Please share the ticket again."
              ));
            }
          } catch (error) {
            await context.sendActivity(MessageFactory.text(
              `Error adding comment: ${error.message}`
            ));
          }
        }
      }
      else if (txt === "check llm context") {
        const response = this.checkLLMContext();
        await context.sendActivity(response);
      }
      else if (txt.includes("@summary")) {
        const summary = await this.generateSummary(this.chatHistories[conversationId]);
        logBotResponse(summary);
        await context.sendActivity(summary);
      } else if (txt.includes("@ticketsummary")) {
        try {
          // Fetch the board summary
          const workItems = await this.getSummaryFromBoardOriginal();
          console.log("Work Items Retrieved:", workItems); // Debugging log

          if (!workItems || workItems.length === 0) {
            throw new Error("No work items found.");
          }

          const boardSummary = this.generateOverallTicketSummary(workItems);

          const sprintSummary = this.generateSprintSummary(workItems);

          const combinedSummary = `${boardSummary}\n\n${sprintSummary}`;

          this.llmContext = `
            You are an assistant with access to the following board and sprint summaries:
      
            ${combinedSummary}
      
            Use this data to answer questions about the tickets.
          `;
          console.log("LLM Context Set:", this.llmContext);

          logBotResponse(combinedSummary);
          await context.sendActivity(combinedSummary);

          await context.sendActivity("Board and sprint summaries have been stored in the LLM context.");
        } catch (error) {
          console.error("Error processing @activeticketsummary command:", error);
          await context.sendActivity("Failed to process the @activeticketsummary command.");
        }
      }
      else if (txt.endsWith("?")) {
        const answer = await this.queryLLM(originalText);
        await context.sendActivity(answer);
      }

      // Regular bot responses
      else {
        const responses = {
          "hello": "Hello! How can I assist you today?",
          "help": "I can help with:\n- Read Azure DevOps tickets (just share a ticket URL or ID)\n- Post comments on Azure Tickets \n- Perform QnA on Tickets \n- Provide conversation summaries (use @summary)",
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

  extractWorkItemIdFromContext(ticketContext) {
    try {
      const idMatch = ticketContext.match(/\*\*ID\*\*: (\d+)/);
      if (idMatch && idMatch[1]) {
        return idMatch[1];
      }
      return null;
    } catch (error) {
      console.error('Error extracting work item ID from context:', error);
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

  async addCommentToWorkItem(workItemId, commentText) {
    try {
      const baseUrl = "https://dev.azure.com/ShorthillsPM";
      const apiUrl = `${baseUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0`;

      console.log(`Adding comment to work item #${workItemId}`);

      const authToken = Buffer.from(`:${this.personalAccessToken}`).toString('base64');
      const authHeader = {
        'Authorization': `Basic ${authToken}`,
        'Content-Type': 'application/json-patch+json'
      };

      const payload = [
        {
          "op": "add",
          "path": "/fields/System.History",
          "value": commentText
        }
      ];

      const response = await axios.patch(apiUrl, payload, {
        headers: authHeader
      });

      console.log(`Comment response status: ${response.status}`);

      return response.data;
    } catch (error) {
      console.error("Error adding comment to work item:", error);

      if (error.response) {
        throw new Error(`Failed to add comment (${error.response.status}): ${error.response.data?.message || 'Unknown error'}`);
      } else if (error.request) {
        throw new Error("Network error: No response received from Azure DevOps API.");
      } else {
        throw new Error(`Error: ${error.message}`);
      }
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
            formattedInfo += `\n📝 **${entry.revisedBy?.displayName || 'Unknown'}**:\n`;
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
  async queryLLM(question) {
    const prompt = `
      ${this.llmContext}
  
      Question: ${question}
      Answer:
    `;

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
      console.error("Error querying LLM:", error);
      return "Error querying the LLM.";
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

  checkLLMContext() {
    if (this.llmContext) {
      console.log("LLM Context is set:", this.llmContext);
      return this.llmContext;
    } else {
      console.log("LLM Context is not set.");
      return "LLM Context is not set.";
    }
  }
  async storeBoardSummaryInLLMContext() {
    try {
      const boardSummary = await this.getSummaryFromBoardoriginal();
      console.log("Board Summary Retrieved:", boardSummary);

      if (!boardSummary || boardSummary === "Unable to fetch ticket summary from the board.") {
        throw new Error("Board summary is empty or could not be fetched.");
      }

      this.llmContext = `
        You are an assistant with access to the following board summary:
  
        ${boardSummary}
  
        Use this data to answer questions about the tickets.
      `; e

      console.log("LLM Context Set:", this.llmContext);
      return "Board summary has been stored in the LLM context.";
    } catch (error) {
      console.error("Error storing board summary in LLM context:", error);
      return "Failed to store board summary in the LLM context.";
    }
  }
  async getSummaryFromBoardOriginal() {
    try {
      const baseUrl = "https://dev.azure.com/ShorthillsPM";
      const project = process.env.AZURE_PROJECT;
      const wiqlUrl = `${baseUrl}/${project}/_apis/wit/wiql?api-version=6.0`;

      const wiqlQuery = {
        query: `
          SELECT [System.Id], [System.State], [System.AssignedTo], [System.Title], [System.IterationPath]
          FROM WorkItems
          WHERE [System.TeamProject] = @project
          ORDER BY [System.ChangedDate] DESC
        `
      };

      const response = await axios.post(wiqlUrl, wiqlQuery, {
        headers: {
          'Authorization': `Basic ${Buffer.from(`:${this.personalAccessToken}`).toString('base64')}`,
          'Content-Type': 'application/json'
        }
      });

      const workItemRefs = response.data.workItems || [];
      const workItemIds = workItemRefs.map(item => item.id);

      if (workItemIds.length === 0) {
        return "There are no tickets on the board.";
      }

      const chunkSize = 200;
      const allWorkItems = [];

      for (let i = 0; i < workItemIds.length; i += chunkSize) {
        const chunk = workItemIds.slice(i, i + chunkSize).join(",");
        const batchUrl = `${baseUrl}/_apis/wit/workitems?ids=${chunk}&fields=System.State,System.AssignedTo,System.Title,System.IterationPath&api-version=6.0`;

        const batchResponse = await axios.get(batchUrl, {
          headers: {
            'Authorization': `Basic ${Buffer.from(`:${this.personalAccessToken}`).toString('base64')}`
          }
        });

        allWorkItems.push(...batchResponse.data.value);
      }

      return allWorkItems;
    } catch (error) {
      console.error("Error fetching board summary:", error);
      return "Unable to fetch ticket summary from the board.";
    }
  }


  generateOverallTicketSummary(allWorkItems) {
    const total = allWorkItems.length;
    const stateCounts = {
      Active: 0,
      Closed: 0,
      Removed: 0,
      New: 0,
      Other: 0
    };

    const userStats = {};

    for (const wi of allWorkItems) {
      const state = wi.fields['System.State'];
      const assignedTo = wi.fields['System.AssignedTo']?.displayName || "Unassigned";

      if (stateCounts[state] !== undefined) {
        stateCounts[state]++;
      } else {
        stateCounts.Other++;
      }

      if (!userStats[assignedTo]) {
        userStats[assignedTo] = {};
      }
      if (!userStats[assignedTo][state]) {
        userStats[assignedTo][state] = 0;
      }
      userStats[assignedTo][state]++;
    }

    let summary = `📊 **Board Summary** - ${total} total tickets\n\n`;
    summary += `- **Active**: ${stateCounts.Active}\n`;
    summary += `- **Closed**: ${stateCounts.Closed}\n`;
    summary += `- **Removed**: ${stateCounts.Removed || 0}\n`;
    summary += `- **New**: ${stateCounts.New || 0}\n`;
    if (stateCounts.Other > 0) {
      summary += `- **Other States**: ${stateCounts.Other}\n`;
    }

    summary += `\n### Tickets per user:\n`;
    for (const [user, states] of Object.entries(userStats)) {
      const userTotal = Object.values(states).reduce((a, b) => a + b, 0);
      summary += `- **${user}**: ${userTotal} tickets\n`;


    }

    return summary;
  }

  generateSprintSummary(workItems) {
    const sprintSummary = {};

    for (const wi of workItems) {
      const sprint = wi.fields['System.IterationPath'] || "Unassigned Sprint";
      const assignedTo = wi.fields['System.AssignedTo']?.displayName || "Unassigned";
      const title = wi.fields['System.Title'] || "Untitled";
      const state = wi.fields['System.State'] || "Unknown";
      const ticketNumber = wi.id;

      if (!sprintSummary[sprint]) {
        sprintSummary[sprint] = {};
      }

      if (!sprintSummary[sprint][assignedTo]) {
        sprintSummary[sprint][assignedTo] = {};
      }

      if (!sprintSummary[sprint][assignedTo][state]) {
        sprintSummary[sprint][assignedTo][state] = [];
      }

      sprintSummary[sprint][assignedTo][state].push(`${ticketNumber}: ${title}`);
    }

    let summary = `📊 **Sprint Summary**\n\n`;
    for (const [sprint, users] of Object.entries(sprintSummary)) {
      summary += `### Sprint: ${sprint}\n`;
      for (const [user, states] of Object.entries(users)) {
        summary += `- **${user}**:\n`;
        for (const [state, tickets] of Object.entries(states)) {
          summary += `  - **${state}** (${tickets.length} tickets):\n`;
          for (const ticket of tickets) {
            summary += `    - ${ticket}\n`;
          }
        }
      }
      summary += `\n`;
    }

    return summary;
  }
}

module.exports.TeamsBot = TeamsBot;
