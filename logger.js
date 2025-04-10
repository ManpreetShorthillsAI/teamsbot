const fs = require('fs');
const path = require('path');

const logDir = path.join(__dirname, 'log');

try {
    if (!fs.existsSync(logDir)) {
        fs.mkdirSync(logDir, { recursive: true });
        console.log("✅ Log directory created at:", logDir);
    }
} catch (err) {
    console.error("❌ Failed to create log directory:", err);
}

const logFile = path.join(logDir, 'chat.log');

function logMessage(sender, message) {
    const timestamp = new Date().toISOString();
    const logLine = `[${timestamp}] ${sender}: ${message}\n`;

    fs.appendFile(logFile, logLine, (err) => {
        if (err) {
            console.error('Failed to write log:', err);
        }
    });
}

function logBotResponse(message) {
    const timestamp = new Date().toISOString();
    const logLine = `[${timestamp}] 🤖 Bot: ${message}\n`;

    fs.appendFile(logFile, logLine, (err) => {
        if (err) {
            console.error('Failed to write bot log:', err);
        }
    });
}

module.exports = { logMessage, logBotResponse };
