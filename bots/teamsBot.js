// teamsBot.js - Enhanced industrial strength implementation with token fixes
const { DialogBot } = require('./dialogBot');
const { SimpleGraphClient } = require('../simpleGraphClient');
const { MessageFactory, TurnContext } = require('botbuilder-core');

/**
 * Enhanced TeamsBot class with industrial-strength token management
 * Uses in-memory storage for conversation references with MongoDB as optional enhancement
 */
class TeamsBot extends DialogBot {
    constructor(conversationState, userState, dialog, adapter) {
        super(conversationState, userState, dialog);

        this.adapter = adapter;
        // In-memory storage for conversation references (production should use persistent storage)
        this.userContextMap = new Map();
        
        this.onMembersAdded(this.handleMembersAdded.bind(this));
    }

    async handleMembersAdded(context, next) {
        const membersAdded = context.activity.membersAdded;
        for (const member of membersAdded) {
            if (member.id !== context.activity.recipient.id) {
                await context.sendActivity('Welcome! Teams SSO is configured. Just type anything to get started.');
            }
        }
        await next();
    }

    /**
     * Handle Teams signin verify state for SSO
     */
    async handleTeamsSigninVerifyState(context, query) {
        console.log('Teams SSO signin verification');
        await this.dialog.run(context, this.dialogState);
    }

    /**
     * Handle Teams signin token exchange for SSO
     */
    async handleTeamsSigninTokenExchange(context, query) {
        console.log('Teams SSO token exchange');
        await this.dialog.run(context, this.dialogState);
    }

    /**
     * Override handleMessage to store context and handle SSO
     */
    async handleMessage(context, next) {
        console.log('Processing message with Teams SSO context');
        
        // Always store user context for token retrieval
        await this.storeUserContext(context);
        
        // Continue with dialog processing
        await super.handleMessage(context, next);
    }

    /**
     * Store user context in memory (enhance with persistent storage for production)
     */
    async storeUserContext(context) {
        try {
            if (!context?.activity?.from?.id) return;

            const userId = context.activity.from.id;
            const conversationReference = TurnContext.getConversationReference(context.activity);
            
            const userContext = {
                userId: userId,
                conversationReference: conversationReference,
                userName: context.activity.from.name || 'Unknown',
                channelId: context.activity.channelId,
                serviceUrl: context.activity.serviceUrl,
                tenantId: context.activity.channelData?.tenant?.id || null,
                conversationId: context.activity.conversation.id,
                aadObjectId: context.activity.from.aadObjectId || null,
                ssoEnabled: true,
                lastUpdated: new Date(),
                createdAt: this.userContextMap.has(userId) ? 
                    this.userContextMap.get(userId).createdAt : new Date()
            };

            this.userContextMap.set(userId, userContext);
            console.log(`SSO context stored for user: ${userId}`);
            
        } catch (error) {
            console.error('Failed to store user context:', error);
        }
    }

    /**
     * INDUSTRIAL STRENGTH TOKEN RETRIEVAL - FIXED VERSION
     * This version properly handles the UserTokenClient in proactive scenarios
     */
    async getTokenForUser(userId, forceRefresh = false) {
        try {
            console.log(`Getting token for user: ${userId} (forceRefresh: ${forceRefresh})`);
            
            // Get user context from memory
            const userContext = this.userContextMap.get(userId);
            if (!userContext?.conversationReference) {
                console.log(`No conversation reference for user: ${userId}`);
                return {
                    success: false,
                    error: 'No conversation context found',
                    message: 'User needs to interact with the bot first'
                };
            }

            console.log(`Found conversation reference for user: ${userId}`);

            // Use Bot Framework's proactive conversation to get token
            return new Promise((resolve) => {
                this.adapter.continueConversationAsync(
                    process.env.MicrosoftAppId,
                    userContext.conversationReference,
                    async (proactiveContext) => {
                        try {
                            // Method 1: Try to get UserTokenClient from the adapter
                            let userTokenClient = null;
                            
                            // Different ways to access UserTokenClient based on Bot Framework version
                            if (proactiveContext.turnState && proactiveContext.adapter.UserTokenClientKey) {
                                userTokenClient = proactiveContext.turnState.get(proactiveContext.adapter.UserTokenClientKey);
                            }
                            
                            // Fallback: Try to get from adapter directly
                            if (!userTokenClient && proactiveContext.adapter.getUserToken) {
                                // Use the regular getUserToken method if available
                                console.log('Using adapter.getUserToken method');
                                
                                if (forceRefresh) {
                                    console.log(`Force refresh requested for user: ${userId}`);
                                    // Try to sign out first
                                    try {
                                        await proactiveContext.adapter.signOutUser(proactiveContext, process.env.connectionName);
                                    } catch (signOutError) {
                                        console.log('Could not sign out user (continuing anyway):', signOutError.message);
                                    }
                                }

                                const tokenResponse = await proactiveContext.adapter.getUserToken(
                                    proactiveContext,
                                    process.env.connectionName,
                                    undefined
                                );

                                if (tokenResponse?.token) {
                                    console.log(`Token retrieved for user ${userId} (length: ${tokenResponse.token.length})`);
                                    
                                    // Update context with token metadata
                                    userContext.lastTokenRetrieved = new Date();
                                    userContext.tokenStatus = 'active';
                                    this.userContextMap.set(userId, userContext);
                                    
                                    resolve({
                                        success: true,
                                        token: tokenResponse.token,
                                        expiration: tokenResponse.expiration,
                                        connectionName: tokenResponse.connectionName,
                                        channelId: tokenResponse.channelId
                                    });
                                    return;
                                }
                            }
                            
                            // Method 2: Use UserTokenClient if available
                            if (userTokenClient) {
                                console.log('Using UserTokenClient');
                                
                                if (forceRefresh) {
                                    console.log(`Force refresh requested for user: ${userId}`);
                                    await userTokenClient.signOutUser(
                                        userId, 
                                        process.env.connectionName, 
                                        userContext.channelId
                                    );
                                }

                                const tokenResponse = await userTokenClient.getUserToken(
                                    userId,
                                    process.env.connectionName,
                                    userContext.channelId,
                                    undefined
                                );

                                if (tokenResponse?.token) {
                                    console.log(`Token retrieved for user ${userId} (length: ${tokenResponse.token.length})`);
                                    
                                    userContext.lastTokenRetrieved = new Date();
                                    userContext.tokenStatus = 'active';
                                    this.userContextMap.set(userId, userContext);
                                    
                                    resolve({
                                        success: true,
                                        token: tokenResponse.token,
                                        expiration: tokenResponse.expiration,
                                        connectionName: tokenResponse.connectionName,
                                        channelId: tokenResponse.channelId
                                    });
                                    return;
                                }
                            }
                            
                            // Method 3: Alternative approach using OAuth helpers
                            if (!userTokenClient && proactiveContext.adapter.createConnectorClient) {
                                console.log('Trying alternative OAuth approach');
                                
                                try {
                                    // Create a connector client
                                    const connectorClient = proactiveContext.adapter.createConnectorClient(proactiveContext.activity.serviceUrl);
                                    
                                    // Get OAuth client
                                    if (connectorClient.userToken) {
                                        const tokenResponse = await connectorClient.userToken.getToken(
                                            userId,
                                            process.env.connectionName,
                                            userContext.channelId,
                                            undefined
                                        );

                                        if (tokenResponse?.token) {
                                            console.log(`Token retrieved via OAuth client for user ${userId}`);
                                            
                                            userContext.lastTokenRetrieved = new Date();
                                            userContext.tokenStatus = 'active';
                                            this.userContextMap.set(userId, userContext);
                                            
                                            resolve({
                                                success: true,
                                                token: tokenResponse.token,
                                                expiration: tokenResponse.expiration,
                                                connectionName: process.env.connectionName,
                                                channelId: userContext.channelId
                                            });
                                            return;
                                        }
                                    }
                                } catch (oauthError) {
                                    console.log('OAuth approach failed:', oauthError.message);
                                }
                            }

                            // If we get here, no token was available
                            console.log(`No token available for user ${userId}`);
                            
                            userContext.tokenStatus = 'unavailable';
                            userContext.lastTokenAttempt = new Date();
                            this.userContextMap.set(userId, userContext);
                            
                            resolve({
                                success: false,
                                error: 'Token not available',
                                message: 'User needs to authenticate first'
                            });
                            
                        } catch (error) {
                            console.error(`Error getting token for user ${userId}:`, error);
                            console.error('Error details:', {
                                name: error.name,
                                message: error.message,
                                stack: error.stack
                            });
                            
                            resolve({
                                success: false,
                                error: 'Token retrieval failed',
                                message: error.message,
                                details: {
                                    errorName: error.name,
                                    availableMethods: {
                                        userTokenClient: !!proactiveContext.turnState?.get(proactiveContext.adapter.UserTokenClientKey),
                                        getUserToken: typeof proactiveContext.adapter.getUserToken === 'function',
                                        createConnectorClient: typeof proactiveContext.adapter.createConnectorClient === 'function'
                                    }
                                }
                            });
                        }
                    }
                );
            });

        } catch (error) {
            console.error(`Error in getTokenForUser for ${userId}:`, error);
            return {
                success: false,
                error: 'Internal error',
                message: error.message
            };
        }
    }

    /**
     * Get user profile using their token (demonstrates token usage)
     */
    async getUserProfile(userId) {
        try {
            const tokenResult = await this.getTokenForUser(userId);
            
            if (!tokenResult.success) {
                console.log(`No token available for profile request: ${userId}`);
                return {
                    success: false,
                    error: tokenResult.error,
                    message: tokenResult.message
                };
            }

            const client = new SimpleGraphClient(tokenResult.token);
            const profile = await client.getMe();
            
            console.log(`Profile retrieved for user: ${userId}`);
            return {
                success: true,
                profile: profile
            };
            
        } catch (error) {
            console.error(`Error getting user profile for ${userId}:`, error);
            return {
                success: false,
                error: 'Profile retrieval failed',
                message: error.message
            };
        }
    }

    /**
     * Validate token and make a test Graph API call
     */
    async validateUserToken(userId) {
        try {
            const tokenResult = await this.getTokenForUser(userId);
            
            if (!tokenResult.success) {
                return { 
                    valid: false, 
                    reason: tokenResult.error,
                    message: tokenResult.message
                };
            }

            // Test the token with a simple Graph API call
            const client = new SimpleGraphClient(tokenResult.token);
            await client.getMe();
            
            return { 
                valid: true, 
                tokenLength: tokenResult.token.length,
                expiration: tokenResult.expiration,
                message: 'Token is valid and working' 
            };
            
        } catch (error) {
            return { 
                valid: false, 
                reason: 'Token validation failed',
                error: error.message 
            };
        }
    }

    /**
     * Get user context information
     */
    getUserContext(userId) {
        const context = this.userContextMap.get(userId);
        if (!context) return null;

        // Return sanitized context (no sensitive data)
        return {
            userId: context.userId,
            userName: context.userName,
            channelId: context.channelId,
            tenantId: context.tenantId,
            conversationId: context.conversationId,
            lastUpdated: context.lastUpdated,
            createdAt: context.createdAt,
            hasConversationReference: !!context.conversationReference,
            ssoEnabled: context.ssoEnabled,
            tokenStatus: context.tokenStatus || 'unknown',
            lastTokenRetrieved: context.lastTokenRetrieved,
            lastTokenAttempt: context.lastTokenAttempt
        };
    }

    /**
     * Health check for the context storage
     */
    async healthCheck() {
        return {
            status: 'healthy',
            userContextCount: this.userContextMap.size,
            timestamp: new Date().toISOString()
        };
    }

    /**
     * Legacy method - kept for backwards compatibility
     */
    async fetchUserTokenById(userId) {
        console.log('Using legacy fetchUserTokenById - redirecting to getTokenForUser');
        const result = await this.getTokenForUser(userId);
        return result.success ? result.token : null;
    }

    /**
     * Get valid token in current context (for dialog use)
     */
    async getValidToken(context) {
        try {
            // In regular context, we can use adapter.getUserToken
            const tokenResponse = await context.adapter.getUserToken(
                context,
                process.env.connectionName
            );

            if (tokenResponse?.token) {
                console.log('Token retrieved in current context');
                return tokenResponse.token;
            } else {
                console.log('No token available in current context');
                return null;
            }
        } catch (error) {
            console.error('Error getting token in context:', error);
            return null;
        }
    }

    /**
     * Handle utility commands
     */
    async onMessage(context) {
        // Store context first
        await this.storeUserContext(context);
        
        if (context.activity?.text) {
            const text = context.activity.text.toLowerCase().trim();
            
            switch (text) {
                case 'token status':
                    await this.handleTokenStatusCommand(context);
                    return;
                case 'my profile':
                    await this.handleProfileCommand(context);
                    return;
                case 'context info':
                    await this.handleContextInfoCommand(context);
                    return;
                default:
                    // Continue with normal dialog flow
                    await super.onMessage(context);
            }
        } else {
            await super.onMessage(context);
        }
    }

    async handleTokenStatusCommand(context) {
        const userId = context.activity.from.id;
        const validation = await this.validateUserToken(userId);
        
        let statusText = `Token Status for ${context.activity.from.name}:\n\n`;
        
        if (validation.valid) {
            statusText += `Status: Valid and working\n`;
            statusText += `Token Length: ${validation.tokenLength} characters\n`;
            if (validation.expiration) {
                statusText += `Expires: ${new Date(validation.expiration).toLocaleString()}\n`;
            }
            statusText += `Auto-refresh: Enabled via Bot Framework\n`;
            statusText += `SSO: Teams Silent Authentication Active\n`;
        } else {
            statusText += `Status: ${validation.reason}\n`;
            statusText += `Message: ${validation.message || 'Please interact with the bot to authenticate'}\n`;
        }
        
        await context.sendActivity(MessageFactory.text(statusText));
    }

    async handleProfileCommand(context) {
        const userId = context.activity.from.id;
        const profileResult = await this.getUserProfile(userId);
        
        if (profileResult.success) {
            const profile = profileResult.profile;
            let profileText = `Your Profile:\n\n`;
            profileText += `Name: ${profile.displayName || 'Not available'}\n`;
            profileText += `Email: ${profile.userPrincipalName || 'Not available'}\n`;
            profileText += `Job Title: ${profile.jobTitle || 'Not available'}\n`;
            profileText += `Department: ${profile.department || 'Not available'}\n`;
            profileText += `Office Location: ${profile.officeLocation || 'Not available'}\n`;
            
            await context.sendActivity(MessageFactory.text(profileText));
        } else {
            await context.sendActivity(`Could not retrieve your profile: ${profileResult.message}`);
        }
    }

    async handleContextInfoCommand(context) {
        const userId = context.activity.from.id;
        const userContext = this.getUserContext(userId);
        
        if (userContext) {
            let contextText = `Your Context Information:\n\n`;
            contextText += `User ID: ${userContext.userId}\n`;
            contextText += `User Name: ${userContext.userName}\n`;
            contextText += `Channel ID: ${userContext.channelId}\n`;
            contextText += `Tenant ID: ${userContext.tenantId || 'Not available'}\n`;
            contextText += `SSO Enabled: ${userContext.ssoEnabled ? 'Yes' : 'No'}\n`;
            contextText += `Token Status: ${userContext.tokenStatus}\n`;
            contextText += `Last Updated: ${userContext.lastUpdated?.toLocaleString() || 'Never'}\n`;
            
            await context.sendActivity(MessageFactory.text(contextText));
        } else {
            await context.sendActivity('No context information found.');
        }
    }
}

module.exports.TeamsBot = TeamsBot;