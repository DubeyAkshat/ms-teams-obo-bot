// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { DialogBot } = require('./dialogBot');
const { SimpleGraphClient } = require('../simpleGraphClient');
const { MessageFactory } = require('botbuilder-core');

/**
 * TeamsBot class extends DialogBot to handle Teams-specific activities.
 */
class TeamsBot extends DialogBot {
    /**
     * Creates an instance of TeamsBot.
     * @param {ConversationState} conversationState - The state management object for conversation state.
     * @param {UserState} userState - The state management object for user state.
     * @param {Dialog} dialog - The dialog to be run by the bot.
     */
    constructor(conversationState, userState, dialog) {
        super(conversationState, userState, dialog);

        this.onMembersAdded(this.handleMembersAdded.bind(this));
        
        // Initialize background task manager
        this.backgroundTasks = new Map();
        this.startBackgroundTaskRunner();
    }

    /**
     * Handles members being added to the conversation.
     * @param {TurnContext} context - The context object for the turn.
     * @param {Function} next - The next middleware function in the pipeline.
     */
    async handleMembersAdded(context, next) {
        const membersAdded = context.activity.membersAdded;
        for (const member of membersAdded) {
            if (member.id !== context.activity.recipient.id) {
                await context.sendActivity('Welcome to TeamsBot. Type anything to get logged in. Type \'logout\' to sign-out.');
            }
        }
        await next();
    }

    /**
     * Handles the Teams signin verify state.
     * @param {TurnContext} context - The context object for the turn.
     * @param {Object} query - The query object from the invoke activity.
     */
    async handleTeamsSigninVerifyState(context, query) {
        console.log('Running dialog with signin/verifystate from an Invoke Activity.');
        await this.dialog.run(context, this.dialogState);
    }

    /**
     * Handles the Teams signin token exchange.
     * @param {TurnContext} context - The context object for the turn.
     * @param {Object} query - The query object from the invoke activity.
     */
    async handleTeamsSigninTokenExchange(context, query) {
        console.log('Running dialog with signin/tokenExchange from an Invoke Activity.');
        await this.dialog.run(context, this.dialogState);
    }

    /**
     * Override onMessage to handle special commands, but preserve main dialog for general messages
     * @param {TurnContext} context - The context object for the turn.
     */
    async onMessage(context) {
        // Check if context.activity and text exist
        if (!context.activity || !context.activity.text) {
            // Continue with normal dialog flow if no text
            await super.onMessage(context);
            return;
        }

        const text = context.activity.text.toLowerCase().trim();

        // Handle ONLY specific utility commands, let everything else go to main dialog
        switch (text) {
            case 'token status':
            case 'check token':
                await this.handleTokenInfoCommand(context);
                return;
            case 'refresh token':
            case 'validate token':
                await this.handleRefreshTokenCommand(context);
                return;
            case 'schedule task':
            case 'background task':
                await this.handleScheduleCommand(context);
                return;
            default:
                // Continue with normal dialog flow for all other messages including "calendar", "hi", etc.
                await super.onMessage(context);
        }
    }

    /**
     * Gets a valid token for the user, automatically refreshing if needed
     * @param {TurnContext} context - The context object for the turn.
     * @returns {Promise<string|null>} Valid access token or null
     */
    async getValidToken(context) {
        try {
            // This is the key method - Bot Framework handles refresh automatically
            const tokenResponse = await context.adapter.getUserToken(
                context,
                process.env.connectionName
            );

            if (tokenResponse && tokenResponse.token) {
                console.log('✅ Token retrieved successfully (auto-refreshed if needed)');
                return tokenResponse.token;
            } else {
                console.log('❌ No token available - user needs to authenticate');
                return null;
            }
        } catch (error) {
            console.error('Error getting token:', error);
            return null;
        }
    }

    /**
     * Handles calendar command with automatic token refresh
     * @param {TurnContext} context - The context object for the turn.
     */
    async handleCalendarCommand(context) {
        const token = await this.getValidToken(context);
        
        if (!token) {
            await context.sendActivity('Please sign in first to view your calendar.');
            return;
        }

        try {
            const client = new SimpleGraphClient(token);
            const events = await client.getTodaysEvents();
            
            if (events && events.length > 0) {
                let eventsText = `📅 **Your calendar for today:**\n\n`;
                
                events.forEach(event => {
                    const startTime = new Date(event.start.dateTime).toLocaleTimeString('en-US', {
                        hour: '2-digit',
                        minute: '2-digit',
                        hour12: true
                    });
                    
                    eventsText += `• **${event.subject}** at ${startTime}\n`;
                });
                
                await context.sendActivity(MessageFactory.text(eventsText));
            } else {
                await context.sendActivity('📅 No events scheduled for today!');
            }
        } catch (error) {
            console.error('Calendar command error:', error);
            
            if (error.message.includes('token') || error.message.includes('auth')) {
                await context.sendActivity('Your session has expired. Please sign in again.');
            } else {
                await context.sendActivity('Sorry, I couldn\'t retrieve your calendar events.');
            }
        }
    }

    /**
     * Handles token info command
     * @param {TurnContext} context - The context object for the turn.
     */
    async handleTokenInfoCommand(context) {
        const token = await this.getValidToken(context);
        
        if (!token) {
            await context.sendActivity('No active token. Please sign in first.');
            return;
        }

        const client = new SimpleGraphClient(token);
        const tokenInfo = client.getTokenInfo();
        
        let infoText = `🔐 **Current Token Status:**\n\n`;
        infoText += `**Token Length:** ${tokenInfo.tokenLength} characters\n`;
        infoText += `**Token Preview:** ${tokenInfo.tokenPreview}\n`;
        infoText += `**Status:** ✅ Valid (auto-refreshed by Bot Framework)\n\n`;
        infoText += `**Endpoints:**\n`;
        infoText += `• Graph API: ${tokenInfo.endpoints.graph}\n`;
        infoText += `• Authority: ${tokenInfo.endpoints.authority}\n`;
        
        await context.sendActivity(MessageFactory.text(infoText));
    }

    /**
     * Handles manual token refresh command
     * @param {TurnContext} context - The context object for the turn.
     */
    async handleRefreshTokenCommand(context) {
        await context.sendActivity('🔄 Checking token status...');
        
        const token = await this.getValidToken(context);
        
        if (token) {
            await context.sendActivity('✅ Token is valid! Bot Framework automatically handles refresh.');
        } else {
            await context.sendActivity('❌ No valid token. Please sign in again.');
        }
    }

    /**
     * Handles schedule command to demonstrate background operations
     * @param {TurnContext} context - The context object for the turn.
     */
    async handleScheduleCommand(context) {
        const userId = context.activity.from.id;
        const conversationRef = context.activity.getConversationReference();
        
        // Schedule a background task to check calendar in 5 minutes
        this.scheduleBackgroundTask(userId, conversationRef, 5 * 60 * 1000); // 5 minutes
        
        await context.sendActivity('⏰ Scheduled a calendar check for 5 minutes from now. I\'ll send you an update automatically!');
    }

    /**
     * Schedules a background task for a user
     * @param {string} userId - The user ID
     * @param {Object} conversationRef - Conversation reference for proactive messaging
     * @param {number} delayMs - Delay in milliseconds
     */
    scheduleBackgroundTask(userId, conversationRef, delayMs) {
        const taskId = `${userId}_${Date.now()}`;
        const executeAt = Date.now() + delayMs;
        
        this.backgroundTasks.set(taskId, {
            userId,
            conversationRef,
            executeAt,
            taskType: 'calendarCheck'
        });
        
        console.log(`📝 Scheduled background task ${taskId} for ${new Date(executeAt)}`);
    }

    /**
     * Starts the background task runner
     */
    startBackgroundTaskRunner() {
        setInterval(async () => {
            await this.runBackgroundTasks();
        }, 60000); // Check every minute
    }

    /**
     * Runs due background tasks
     */
    async runBackgroundTasks() {
        const now = Date.now();
        const dueTasks = [];
        
        // Find due tasks
        for (const [taskId, task] of this.backgroundTasks.entries()) {
            if (task.executeAt <= now) {
                dueTasks.push({ taskId, ...task });
                this.backgroundTasks.delete(taskId);
            }
        }
        
        // Execute due tasks
        for (const task of dueTasks) {
            try {
                await this.executeBackgroundTask(task);
            } catch (error) {
                console.error(`Background task ${task.taskId} failed:`, error);
            }
        }
    }

    /**
     * Executes a background task
     * @param {Object} task - The task to execute
     */
    async executeBackgroundTask(task) {
        console.log(`🚀 Executing background task ${task.taskId}`);
        
        try {
            // Create a context for proactive messaging
            await this.adapter.continueConversationAsync(
                process.env.MicrosoftAppId,
                task.conversationRef,
                async (proactiveContext) => {
                    // Try to get a valid token for the user
                    const token = await this.getValidToken(proactiveContext);
                    
                    if (!token) {
                        await proactiveContext.sendActivity(
                            '⚠️ Background task couldn\'t run - your session expired. Please sign in again when you\'re back!'
                        );
                        return;
                    }
                    
                    if (task.taskType === 'calendarCheck') {
                        await this.executeCalendarCheck(proactiveContext, token);
                    }
                }
            );
        } catch (error) {
            console.error('Error in background task execution:', error);
        }
    }

    /**
     * Executes a calendar check background task
     * @param {TurnContext} context - The context object
     * @param {string} token - Valid access token
     */
    async executeCalendarCheck(context, token) {
        try {
            const client = new SimpleGraphClient(token);
            const events = await client.getTodaysEvents();
            
            let message = '🔔 **Scheduled Calendar Update:**\n\n';
            
            if (events && events.length > 0) {
                message += `You have ${events.length} event(s) remaining today:\n\n`;
                
                const upcomingEvents = events.filter(event => 
                    new Date(event.start.dateTime) > new Date()
                );
                
                if (upcomingEvents.length > 0) {
                    upcomingEvents.slice(0, 3).forEach(event => {
                        const startTime = new Date(event.start.dateTime).toLocaleTimeString('en-US', {
                            hour: '2-digit',
                            minute: '2-digit',
                            hour12: true
                        });
                        message += `• **${event.subject}** at ${startTime}\n`;
                    });
                } else {
                    message += 'All your events for today are completed! 🎉';
                }
            } else {
                message += 'No events scheduled for today. Enjoy your free time! 😊';
            }
            
            await context.sendActivity(MessageFactory.text(message));
            
        } catch (error) {
            console.error('Calendar check error:', error);
            await context.sendActivity('❌ Couldn\'t check your calendar. The token may have expired.');
        }
    }
}

module.exports.TeamsBot = TeamsBot;