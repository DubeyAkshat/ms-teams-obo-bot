// =============================================================================
// HTTP/HTTPS REQUEST INTERCEPTION - MUST BE FIRST
// =============================================================================
console.log('Setting up HTTP/HTTPS request interception...');

const originalHttpRequest = require('http').request;
const originalHttpsRequest = require('https').request;

// Intercept HTTP requests
require('http').request = function(options, callback) {
    console.log('=== HTTP REQUEST INTERCEPTED ===');
    console.log('URL:', options.href || `${options.protocol || 'http:'}//${options.host || options.hostname}:${options.port || 80}${options.path}`);
    console.log('Method:', options.method || 'GET');
    console.log('Headers:', JSON.stringify(options.headers, null, 2));
    console.log('================================');
    
    const req = originalHttpRequest.call(this, options, callback);
    
    // Intercept request body
    const originalWrite = req.write;
    req.write = function(chunk) {
        console.log('=== HTTP REQUEST BODY ===');
        console.log(chunk.toString());
        console.log('========================');
        return originalWrite.call(this, chunk);
    };
    
    return req;
};

// Intercept HTTPS requests  
require('https').request = function(options, callback) {
    console.log('=== HTTPS REQUEST INTERCEPTED ===');
    console.log('URL:', options.href || `${options.protocol || 'https:'}//${options.host || options.hostname}:${options.port || 443}${options.path}`);
    console.log('Method:', options.method || 'GET');
    console.log('Headers:', JSON.stringify(options.headers, null, 2));
    console.log('==================================');
    
    const req = originalHttpsRequest.call(this, options, callback);
    
    // Intercept request body
    const originalWrite = req.write;
    req.write = function(chunk) {
        console.log('=== HTTPS REQUEST BODY ===');
        console.log(chunk.toString());
        console.log('=========================');
        return originalWrite.call(this, chunk);
    };
    
    return req;
};

console.log('HTTP/HTTPS interception setup complete.');

// =============================================================================
// ORIGINAL CODE STARTS HERE
// =============================================================================

// Import required packages
const path = require('path');

// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

const restify = require('restify');

// Import required bot services.
const {
    CloudAdapter,
    ConversationState,
    MemoryStorage,
    UserState,
    ConfigurationBotFrameworkAuthentication,
    TeamsSSOTokenExchangeMiddleware
} = require('botbuilder');

const { TeamsBot } = require('./bots/teamsBot');
const { MainDialog } = require('./dialogs/mainDialog');
const { env } = require('process');

console.log('Connection name from env:', env.connectionName);

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);

// Create adapter with proper error handling
const adapter = new CloudAdapter(botFrameworkAuthentication);
const memoryStorage = new MemoryStorage();

// ENABLE Teams SSO Token Exchange Middleware for silent authentication
const tokenExchangeMiddleware = new TeamsSSOTokenExchangeMiddleware(memoryStorage, env.connectionName);
adapter.use(tokenExchangeMiddleware);

// Enhanced SSO debugging middleware
adapter.use(async (context, next) => {
    console.log('=== DEBUG: SSO Processing ===');
    console.log('Activity type:', context.activity.type);
    console.log('Channel ID:', context.activity.channelId);
    console.log('User ID:', context.activity.from?.id);
    console.log('User Name:', context.activity.from?.name);
    console.log('Activity ID:', context.activity.id);
    console.log('Conversation ID:', context.activity.conversation?.id);
    console.log('Service URL:', context.activity.serviceUrl);
    console.log('Channel Data:', JSON.stringify(context.activity.channelData, null, 2));
    
    try {
        await next();
        console.log('=== DEBUG: SSO Processing Complete ===');
    } catch (error) {
        console.log('=== DEBUG: SSO Processing Error ===');
        console.log('Error type:', error.constructor.name);
        console.log('Error message:', error.message);
        console.log('Stack:', error.stack);
        throw error;
    }
});

// Enhanced error handler
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${ error }`);
    console.error('Error details:', {
        name: error.name,
        message: error.message,
        stack: error.stack
    });
    
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );
    
    // Clear conversation state on error
    await conversationState.delete(context);
};

// Create conversation and user state
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create the main dialog
const dialog = new MainDialog();

// Create the enhanced bot with adapter for proactive messaging
const bot = new TeamsBot(conversationState, userState, dialog, adapter);

// Create HTTP server with enhanced configuration
const server = restify.createServer({
    name: 'Teams SSO Bot',
    version: '1.0.0'
});

server.use(restify.plugins.bodyParser());

// Enhanced CORS middleware
server.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
    
    if (req.method === 'OPTIONS') {
        res.send(200);
        return;
    }
    
    next();
});

server.listen(process.env.port || process.env.PORT || 3978, function() {
    console.log(`\n${ server.name } listening to ${ server.url }`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo talk to your bot, open the emulator select "Open Bot"');
});

// =============================================================================
// ENHANCED API ENDPOINTS - INDUSTRIAL STRENGTH TOKEN RETRIEVAL
// =============================================================================

/**
 * Enhanced health check endpoint
 */
server.get('/health', async (req, res) => {
    try {
        const healthInfo = await bot.healthCheck();
        
        res.send(200, {
            status: 'Bot is running',
            ssoEnabled: true,
            timestamp: new Date().toISOString(),
            botHealth: healthInfo,
            environment: {
                nodeVersion: process.version,
                connectionName: process.env.connectionName,
                appId: process.env.MicrosoftAppId ? 'configured' : 'missing'
            }
        });
    } catch (error) {
        console.error('Health check error:', error);
        res.send(200, {
            status: 'Bot is running with warnings',
            ssoEnabled: true,
            timestamp: new Date().toISOString(),
            error: error.message
        });
    }
});

/**
 * Enhanced token retrieval endpoint
 * GET /api/token/:userId
 */
server.get('/api/token/:userId', async (req, res) => {
    const userId = req.params.userId;
    console.log(`Token request for user: ${userId}`);
    
    if (!userId) {
        return res.send(400, { 
            success: false,
            error: 'User ID is required' 
        });
    }

    try {
        const result = await bot.getTokenForUser(userId);
        
        if (result.success) {
            console.log(`Token retrieved for user: ${userId}`);
            res.send(200, {
                success: true,
                userId: userId,
                token: result.token,
                expiration: result.expiration,
                connectionName: result.connectionName,
                channelId: result.channelId,
                timestamp: new Date().toISOString(),
                message: 'Token retrieved successfully',
                tokenLength: result.token.length,
                tokenPreview: `${result.token.substring(0, 20)}...${result.token.substring(result.token.length - 10)}`
            });
        } else {
            res.send(404, {
                success: false,
                userId: userId,
                error: result.error,
                message: result.message,
                details: result.details,
                timestamp: new Date().toISOString()
            });
        }
    } catch (error) {
        console.error(`Token retrieval error for ${userId}:`, error);
        res.send(500, {
            success: false,
            error: 'Internal server error',
            message: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Enhanced token refresh endpoint
 * POST /api/token/:userId/refresh
 */
server.post('/api/token/:userId/refresh', async (req, res) => {
    const userId = req.params.userId;
    
    try {
        const result = await bot.getTokenForUser(userId, true); // Force refresh
        
        if (result.success) {
            res.send(200, {
                success: true,
                token: result.token,
                expiration: result.expiration,
                message: 'Token refreshed successfully',
                timestamp: new Date().toISOString(),
                tokenLength: result.token.length
            });
        } else {
            res.send(404, {
                success: false,
                error: result.error,
                message: result.message,
                details: result.details,
                timestamp: new Date().toISOString()
            });
        }
    } catch (error) {
        console.error(`Token refresh error for ${userId}:`, error);
        res.send(500, {
            success: false,
            error: 'Token refresh failed',
            message: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Enhanced user profile endpoint
 * GET /api/user/:userId/profile
 */
server.get('/api/user/:userId/profile', async (req, res) => {
    const userId = req.params.userId;
    
    try {
        const result = await bot.getUserProfile(userId);
        
        if (result.success) {
            res.send(200, {
                success: true,
                profile: result.profile,
                timestamp: new Date().toISOString()
            });
        } else {
            res.send(404, {
                success: false,
                error: result.error,
                message: result.message,
                timestamp: new Date().toISOString()
            });
        }
    } catch (error) {
        console.error(`Profile retrieval error for ${userId}:`, error);
        res.send(500, {
            success: false,
            error: 'Profile retrieval failed',
            message: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Enhanced user context endpoint
 * GET /api/user/:userId/context
 */
server.get('/api/user/:userId/context', async (req, res) => {
    const userId = req.params.userId;
    
    if (!userId) {
        return res.send(400, {
            success: false,
            error: 'User ID is required'
        });
    }

    try {
        const userContext = bot.getUserContext(userId);
        
        if (userContext) {
            res.send(200, {
                success: true,
                context: userContext,
                timestamp: new Date().toISOString()
            });
        } else {
            res.send(404, {
                success: false,
                error: 'User context not found',
                message: 'No context stored for this user ID',
                timestamp: new Date().toISOString()
            });
        }
    } catch (error) {
        console.error(`Error fetching context for user ${userId}:`, error);
        res.send(500, {
            success: false,
            error: 'Internal server error',
            message: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Enhanced token validation endpoint
 * GET /api/token/:userId/validate
 */
server.get('/api/token/:userId/validate', async (req, res) => {
    const userId = req.params.userId;
    
    try {
        const validation = await bot.validateUserToken(userId);
        
        res.send(validation.valid ? 200 : 404, {
            success: validation.valid,
            valid: validation.valid,
            reason: validation.reason,
            message: validation.message,
            tokenLength: validation.tokenLength,
            expiration: validation.expiration,
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        console.error(`Token validation error for ${userId}:`, error);
        res.send(500, {
            success: false,
            valid: false,
            error: 'Validation failed',
            message: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

/**
 * Batch token retrieval endpoint (for multiple users)
 * POST /api/tokens/batch
 * Body: { userIds: ["userId1", "userId2", ...] }
 */
server.post('/api/tokens/batch', async (req, res) => {
    const { userIds } = req.body;
    
    if (!Array.isArray(userIds) || userIds.length === 0) {
        return res.send(400, {
            success: false,
            error: 'userIds array is required'
        });
    }

    if (userIds.length > 50) {
        return res.send(400, {
            success: false,
            error: 'Maximum 50 users per batch request'
        });
    }

    try {
        const results = await Promise.allSettled(
            userIds.map(userId => bot.getTokenForUser(userId))
        );

        const tokenResults = results.map((result, index) => ({
            userId: userIds[index],
            success: result.status === 'fulfilled' && result.value.success,
            token: result.status === 'fulfilled' && result.value.success ? result.value.token : null,
            error: result.status === 'fulfilled' ? result.value.error : result.reason?.message,
            message: result.status === 'fulfilled' ? result.value.message : 'Request failed'
        }));

        const successCount = tokenResults.filter(r => r.success).length;

        res.send(200, {
            success: successCount > 0,
            totalRequested: userIds.length,
            successCount: successCount,
            failureCount: userIds.length - successCount,
            results: tokenResults,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        console.error('Batch token retrieval error:', error);
        res.send(500, {
            success: false,
            error: 'Batch operation failed',
            message: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

// =============================================================================
// BOT MESSAGE HANDLING
// =============================================================================

/**
 * Enhanced message handling with better logging
 */
server.post('/api/messages', async (req, res) => {
    console.log('=== Incoming Bot Message ===');
    console.log('Headers:', {
        'content-type': req.headers['content-type'],
        'authorization': req.headers['authorization'] ? 'present' : 'missing',
        'user-agent': req.headers['user-agent']
    });
    console.log('Body type:', req.body?.type);
    console.log('From:', req.body?.from?.name, req.body?.from?.id);
    console.log('Conversation:', req.body?.conversation?.id);
    console.log('Channel:', req.body?.channelId);
    
    try {
        // Route received request to adapter for processing
        console.log('=== Starting adapter.process() ===');
        await adapter.process(req, res, (context) => bot.run(context));
        console.log('=== Message Processed Successfully ===');
    } catch (error) {
        console.error('=== Message Processing Error ===');
        console.error('Error:', error);
    }
});

// =============================================================================
// GRACEFUL SHUTDOWN AND ERROR HANDLING
// =============================================================================

process.on('SIGINT', async () => {
    console.log('\nShutting down gracefully...');
    
    try {
        // Cleanup resources if needed
        console.log('Bot shutdown complete');
    } catch (error) {
        console.error('Error during shutdown:', error);
    }
    
    process.exit(0);
});

// Handle uncaught exceptions
process.on('uncaughtException', (error) => {
    console.error('Uncaught Exception:', error);
    process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection at:', promise, 'reason:', reason);
});

// Export for testing
module.exports = { server, bot };