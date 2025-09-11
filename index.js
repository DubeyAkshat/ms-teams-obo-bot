// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// index.js is used to setup and configure your bot

// Import required pckages
const path = require('path');

// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

const restify = require('restify');

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
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
console.log('All env vars:', process.env);

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new CloudAdapter(botFrameworkAuthentication);
const memoryStorage = new MemoryStorage();
// const tokenExchangeMiddleware = new TeamsSSOTokenExchangeMiddleware(memoryStorage, env.connectionName);
// adapter.use(tokenExchangeMiddleware);

adapter.use(async (context, next) => {
    console.log('=== DEBUG: Before SSO middleware ===');
    console.log('Activity type:', context.activity.type);
    console.log('Channel ID:', context.activity.channelId);
    
    try {
        await next();
        console.log('=== DEBUG: After SSO middleware - SUCCESS ===');
    } catch (error) {
        console.log('=== DEBUG: SSO middleware error ===');
        console.log('Error type:', error.constructor.name);
        console.log('Error message:', error.message);
        console.log('Stack:', error.stack);
        throw error;
    }
});

adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights. See https://aka.ms/bottelemetry for telemetry
    //       configuration instructions.
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

     // Uncomment below commented line for local debugging.
     // await context.sendActivity(`Sorry, it looks like something went wrong. Exception Caught: ${error}`);

    // Clear out state
    await conversationState.delete(context);
};

// Define the state store for your bot.
// See https://aka.ms/about-bot-state to learn more about using MemoryStorage.
// A bot requires a state storage system to persist the dialog and user state between messages.
//const memoryStorage = new MemoryStorage();

// Create conversation and user state with in-memory storage provider.
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create the main dialog.
const dialog = new MainDialog();
// Create the bot that will handle incoming messages.
const bot = new TeamsBot(conversationState, userState, dialog);

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, function() {
    console.log(`\n${ server.name } listening to ${ server.url }`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo talk to your bot, open the emulator select "Open Bot"');
});

// Add test endpoints
server.get('/test', async (req, res) => {
    res.send({ status: 'Bot is running', timestamp: new Date().toISOString() });
});

server.get('/health', async (req, res) => {
    res.send('Bot is running');
});

// Listen for incoming requests.
server.post('/api/messages', async (req, res) => {
    console.log('Received request:', {
        headers: req.headers,
        body: req.body ? 'Body present' : 'No body'
    });
    
    try {
        // Route received a request to adapter for processing
        await adapter.process(req, res, (context) => bot.run(context));
        console.log('Request processed successfully');
    } catch (error) {
        console.error('Error processing request:', error);
    }
});