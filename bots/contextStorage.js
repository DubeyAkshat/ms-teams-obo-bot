// contextStorage.js
const { MongoClient } = require('mongodb');

class ContextStorage {
    constructor() {
        this.client = null;
        this.db = null;
        this.collection = null;
        this.isConnected = false;
    }

    /**
     * Initialize MongoDB connection
     */
    async initialize() {
        try {
            const connectionString = process.env.MONGODB_CONNECTION_STRING || 'mongodb://localhost:27017';
            const dbName = process.env.MONGODB_DB_NAME || 'teamsbot';
            const collectionName = process.env.MONGODB_COLLECTION_NAME || 'user_contexts';

            this.client = new MongoClient(connectionString);
            await this.client.connect();
            
            this.db = this.client.db(dbName);
            this.collection = this.db.collection(collectionName);
            
            // Create index on userId for better performance
            await this.collection.createIndex({ userId: 1 }, { unique: true });
            
            this.isConnected = true;
            console.log('✅ MongoDB connected successfully');
        } catch (error) {
            console.error('❌ MongoDB connection failed:', error);
            throw error;
        }
    }

    /**
     * Store user context (conversation reference and other data)
     * @param {string} userId - The user ID
     * @param {Object} conversationReference - The conversation reference
     * @param {Object} additionalData - Any additional context data
     */
    async storeUserContext(userId, conversationReference, additionalData = {}) {
        if (!this.isConnected) {
            throw new Error('MongoDB not connected');
        }

        try {
            const now = new Date();
            const contextData = {
                userId,
                conversationReference,
                ...additionalData,
                lastUpdated: now
            };

            // Check if document exists first
            const existingDoc = await this.collection.findOne({ userId });
            
            if (existingDoc) {
                // Update existing document (don't touch createdAt)
                await this.collection.updateOne(
                    { userId },
                    { $set: contextData }
                );
            } else {
                // Insert new document with createdAt
                contextData.createdAt = now;
                await this.collection.insertOne(contextData);
            }

            console.log(`✅ Context stored for user: ${userId}`);
            return true;
        } catch (error) {
            console.error(`❌ Failed to store context for user ${userId}:`, error);
            return false;
        }
    }

    /**
     * Retrieve user context by user ID
     * @param {string} userId - The user ID
     * @returns {Object|null} The user context or null if not found
     */
    async getUserContext(userId) {
        if (!this.isConnected) {
            throw new Error('MongoDB not connected');
        }

        try {
            const context = await this.collection.findOne({ userId });
            return context;
        } catch (error) {
            console.error(`❌ Failed to retrieve context for user ${userId}:`, error);
            return null;
        }
    }

    /**
     * Update user context
     * @param {string} userId - The user ID
     * @param {Object} updateData - Data to update
     */
    async updateUserContext(userId, updateData) {
        if (!this.isConnected) {
            throw new Error('MongoDB not connected');
        }

        try {
            const result = await this.collection.updateOne(
                { userId },
                { 
                    $set: { 
                        ...updateData, 
                        lastUpdated: new Date() 
                    }
                }
            );

            return result.matchedCount > 0;
        } catch (error) {
            console.error(`❌ Failed to update context for user ${userId}:`, error);
            return false;
        }
    }

    /**
     * Remove user context
     * @param {string} userId - The user ID
     */
    async removeUserContext(userId) {
        if (!this.isConnected) {
            throw new Error('MongoDB not connected');
        }

        try {
            const result = await this.collection.deleteOne({ userId });
            return result.deletedCount > 0;
        } catch (error) {
            console.error(`❌ Failed to remove context for user ${userId}:`, error);
            return false;
        }
    }

    /**
     * Get all users with stored contexts (for admin purposes)
     * @param {number} limit - Maximum number of records to return
     */
    async getAllUserContexts(limit = 100) {
        if (!this.isConnected) {
            throw new Error('MongoDB not connected');
        }

        try {
            const contexts = await this.collection
                .find({}, { projection: { userId: 1, lastUpdated: 1, createdAt: 1 } })
                .limit(limit)
                .toArray();
            
            return contexts;
        } catch (error) {
            console.error('❌ Failed to retrieve all contexts:', error);
            return [];
        }
    }

    /**
     * Close MongoDB connection
     */
    async close() {
        if (this.client) {
            await this.client.close();
            this.isConnected = false;
            console.log('📴 MongoDB connection closed');
        }
    }

    /**
     * Health check
     */
    async healthCheck() {
        try {
            if (!this.isConnected) return false;
            await this.db.admin().ping();
            return true;
        } catch (error) {
            console.error('❌ MongoDB health check failed:', error);
            return false;
        }
    }
}

module.exports = { ContextStorage };