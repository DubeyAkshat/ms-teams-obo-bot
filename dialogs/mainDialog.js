// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ConfirmPrompt, DialogSet, DialogTurnStatus, OAuthPrompt, WaterfallDialog } = require('botbuilder-dialogs');
const { LogoutDialog } = require('./logoutDialog');
const { SimpleGraphClient } = require('../simpleGraphClient');
const { MessageFactory, CardFactory } = require('botbuilder-core');

const CONFIRM_PROMPT = 'ConfirmPrompt';
const MAIN_DIALOG = 'MainDialog';
const MAIN_WATERFALL_DIALOG = 'MainWaterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';

/**
 * MainDialog class extends LogoutDialog to handle the main dialog flow.
 */
class MainDialog extends LogoutDialog {
    /**
     * Creates an instance of MainDialog.
     * @param {string} connectionName - The connection name for the OAuth provider.
     */
    constructor() {
        super(MAIN_DIALOG, process.env.connectionName);

        this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
            connectionName: process.env.connectionName,
            text: 'Please Sign In',
            title: 'Sign In',
            timeout: 300000
        }));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.promptStep.bind(this),
            this.loginStep.bind(this),
            this.ensureOAuth.bind(this),
            this.displayToken.bind(this)
        ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {TurnContext} context - The context object for the turn.
     * @param {StatePropertyAccessor} accessor - The state property accessor for the dialog state.
     */
    async run(context, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);
        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    /**
     * Prompts the user to sign in.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async promptStep(stepContext) {
        return await stepContext.beginDialog(OAUTH_PROMPT);
    }

    /**
     * Handles the login step.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async loginStep(stepContext) {
        const tokenResponse = stepContext.result;
        if (!tokenResponse || !tokenResponse.token) {
            await stepContext.context.sendActivity('Login was not successful, please try again.');
            return await stepContext.endDialog();
        } else {
            const client = new SimpleGraphClient(tokenResponse.token);
            
            try {
                // Get user information
                const me = await client.getMe();
                const title = me ? me.jobTitle : 'Unknown';
                
                await stepContext.context.sendActivity(
                    MessageFactory.text(`Welcome ${me.displayName} (${me.userPrincipalName})!`)
                );

                // Get calendar events
                const events = await client.getTodaysEvents();
                
                if (events && events.length > 0) {
                    let eventsText = `📅 **Your calendar for today:**\n\n`;
                    
                    events.forEach(event => {
                        const startTime = new Date(event.start.dateTime).toLocaleTimeString('en-US', {
                            hour: '2-digit',
                            minute: '2-digit',
                            hour12: true
                        });
                        const endTime = new Date(event.end.dateTime).toLocaleTimeString('en-US', {
                            hour: '2-digit',
                            minute: '2-digit',
                            hour12: true
                        });
                        
                        eventsText += `• **${event.subject}**\n`;
                        eventsText += `  ⏰ ${startTime} - ${endTime}\n`;
                        
                        if (event.location && event.location.displayName) {
                            eventsText += `  📍 ${event.location.displayName}\n`;
                        }
                        
                        if (event.organizer && event.organizer.emailAddress) {
                            eventsText += `  👤 ${event.organizer.emailAddress.name}\n`;
                        }
                        
                        eventsText += '\n';
                    });
                    
                    await stepContext.context.sendActivity(MessageFactory.text(eventsText));
                } else {
                    await stepContext.context.sendActivity(
                        MessageFactory.text('📅 No events scheduled for today. Enjoy your free time!')
                    );
                }

                // Get upcoming events (next 10)
                const upcomingEvents = await client.getCalendarEvents();
                
                if (upcomingEvents && upcomingEvents.length > 0) {
                    let upcomingText = `📆 **Your upcoming events:**\n\n`;
                    
                    upcomingEvents.slice(0, 5).forEach(event => {
                        const eventDate = new Date(event.start.dateTime).toLocaleDateString('en-US', {
                            weekday: 'short',
                            month: 'short',
                            day: 'numeric'
                        });
                        const startTime = new Date(event.start.dateTime).toLocaleTimeString('en-US', {
                            hour: '2-digit',
                            minute: '2-digit',
                            hour12: true
                        });
                        
                        upcomingText += `• **${event.subject}**\n`;
                        upcomingText += `  📅 ${eventDate} at ${startTime}\n\n`;
                    });
                    
                    await stepContext.context.sendActivity(MessageFactory.text(upcomingText));
                }

                // Try to get and display user photo (optional)
                try {
                    const photoBase64 = await client.getPhotoAsync(tokenResponse.token);
                    const card = CardFactory.thumbnailCard("Your Profile Picture", CardFactory.images([photoBase64]));
                    await stepContext.context.sendActivity({ attachments: [card] });
                } catch (photoError) {
                    console.log('Could not fetch user photo (this is normal if user has no photo):', photoError.message);
                    // Don't show error to user for missing photo
                }

            } catch (error) {
                console.error('Error in loginStep:', error);
                await stepContext.context.sendActivity(
                    MessageFactory.text(`Welcome! Authentication successful, but there was an issue accessing your calendar: ${error.message}`)
                );
            }

            return await stepContext.prompt(CONFIRM_PROMPT, 'Would you like to view your access token?');
        }
    }

    /**
     * Ensures the OAuth token is available.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async ensureOAuth(stepContext) {
        await stepContext.context.sendActivity('Thank you.');

        const result = stepContext.result;
        if (result) {
            return await stepContext.beginDialog(OAUTH_PROMPT);
        }
        return await stepContext.endDialog();
    }

    /**
     * Displays comprehensive token information to the user.
     * @param {WaterfallStepContext} stepContext - The waterfall step context.
     */
    async displayToken(stepContext) {
        const tokenResponse = stepContext.result;
        if (tokenResponse && tokenResponse.token) {
            let tokenInfo = `🔐 **Token Information:**\n\n`;
            
            // Access Token (truncated for display)
            const truncatedToken = tokenResponse.token.length > 50 
                ? tokenResponse.token.substring(0, 50) + '...[truncated]'
                : tokenResponse.token;
            tokenInfo += `**Access Token:** \`${truncatedToken}\`\n\n`;
            
            // Full Access Token (in code block for easy copying)
            tokenInfo += `**Full Access Token:**\n\`\`\`\n${tokenResponse.token}\n\`\`\`\n\n`;
            
            // Token expiration if available
            if (tokenResponse.expiration) {
                const expirationDate = new Date(tokenResponse.expiration);
                tokenInfo += `**Expires:** ${expirationDate.toLocaleString()}\n\n`;
            }
            
            // Refresh token if available
            if (tokenResponse.refreshToken) {
                const truncatedRefreshToken = tokenResponse.refreshToken.length > 50 
                    ? tokenResponse.refreshToken.substring(0, 50) + '...[truncated]'
                    : tokenResponse.refreshToken;
                tokenInfo += `**Refresh Token:** \`${truncatedRefreshToken}\`\n\n`;
                tokenInfo += `**Full Refresh Token:**\n\`\`\`\n${tokenResponse.refreshToken}\n\`\`\`\n\n`;
            } else {
                tokenInfo += `**Refresh Token:** Not available (may not be provided by this OAuth flow)\n\n`;
            }
            
            // OAuth endpoints and URLs
            tokenInfo += `**OAuth Configuration:**\n`;
            tokenInfo += `• **Connection Name:** ${process.env.connectionName}\n`;
            tokenInfo += `• **Authority:** https://login.microsoftonline.com/${process.env.MicrosoftAppTenantId || 'common'}\n`;
            tokenInfo += `• **Token Endpoint:** https://login.microsoftonline.com/${process.env.MicrosoftAppTenantId || 'common'}/oauth2/v2.0/token\n`;
            tokenInfo += `• **Authorization Endpoint:** https://login.microsoftonline.com/${process.env.MicrosoftAppTenantId || 'common'}/oauth2/v2.0/authorize\n`;
            tokenInfo += `• **Bot Framework Token Service:** https://token.botframework.com\n\n`;
            
            // Additional token properties
            if (tokenResponse.channelId) {
                tokenInfo += `**Channel ID:** ${tokenResponse.channelId}\n`;
            }
            
            // Token usage examples
            tokenInfo += `**Usage Examples:**\n`;
            tokenInfo += `• **Authorization Header:** \`Bearer ${truncatedToken}\`\n`;
            tokenInfo += `• **Graph API Call:** \`curl -H "Authorization: Bearer YOUR_TOKEN" https://graph.microsoft.com/v1.0/me\`\n\n`;
            
            tokenInfo += `⚠️ **Security Note:** Keep these tokens secure and never share them publicly!`;
            
            await stepContext.context.sendActivity(MessageFactory.text(tokenInfo));
            
            // Log additional details for debugging (server-side only)
            console.log('=== TOKEN RESPONSE DETAILS ===');
            console.log('Token Response Object:', JSON.stringify(tokenResponse, null, 2));
            console.log('=====================================');
        } else {
            await stepContext.context.sendActivity('No token information available.');
        }
        return await stepContext.endDialog();
    }
}

module.exports.MainDialog = MainDialog;