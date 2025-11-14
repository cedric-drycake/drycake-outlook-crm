// SharePoint REST API Client
// This handles all interactions with SharePoint lists

class SharePointClient {
    constructor(siteUrl) {
        this.siteUrl = https://drycakesbc.sharepoint.com/:u:/s/drycakestreak/Ee_yYI_Zei5CnPI9vLBXkbEB9OpopzmPV2026nByuBgjSA?e=fnKk5Z || window.location.origin; // Your SharePoint site URL
        this.baseUrl = `${this.siteUrl}/_api/web/lists`;
        
        // List names in SharePoint
        this.lists = {
            pipelines: 'CRM_Pipelines',
            boxes: 'CRM_Boxes',
            emails: 'CRM_Emails',
            activities: 'CRM_Activities',
            stages: 'CRM_Stages'
        };
    }
    
    // Get request digest for POST operations
    async getRequestDigest() {
        try {
            const response = await fetch(`${this.siteUrl}/_api/contextinfo`, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });
            const data = await response.json();
            return data.d.GetContextWebInformation.FormDigestValue;
        } catch (error) {
            console.error('Error getting request digest:', error);
            throw error;
        }
    }
    
    // Generic GET request
    async get(endpoint, select = '', filter = '', expand = '', orderby = '') {
        try {
            let url = `${this.baseUrl}${endpoint}`;
            let params = [];
            
            if (select) params.push(`$select=${select}`);
            if (filter) params.push(`$filter=${filter}`);
            if (expand) params.push(`$expand=${expand}`);
            if (orderby) params.push(`$orderby=${orderby}`);
            
            if (params.length > 0) {
                url += '?' + params.join('&');
            }
            
            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const data = await response.json();
            return data.d.results || data.d;
        } catch (error) {
            console.error('GET request error:', error);
            throw error;
        }
    }
    
    // Generic POST request
    async post(endpoint, data) {
        try {
            const digest = await this.getRequestDigest();
            
            const response = await fetch(`${this.baseUrl}${endpoint}`, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': digest
                },
                body: JSON.stringify(data)
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const result = await response.json();
            return result.d;
        } catch (error) {
            console.error('POST request error:', error);
            throw error;
        }
    }
    
    // Generic UPDATE request
    async update(endpoint, itemId, data) {
        try {
            const digest = await this.getRequestDigest();
            
            const response = await fetch(`${this.baseUrl}${endpoint}/items(${itemId})`, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': digest,
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'MERGE'
                },
                body: JSON.stringify(data)
            });
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            return true;
        } catch (error) {
            console.error('UPDATE request error:', error);
            throw error;
        }
    }
    
    // PIPELINES
    async getPipelines() {
        return await this.get(
            `/getByTitle('${this.lists.pipelines}')/items`,
            'ID,Title,Description,Created,Modified',
            '',
            '',
            'Title'
        );
    }
    
    async createPipeline(name, description = '') {
        const data = {
            __metadata: { type: 'SP.Data.CRM_PipelinesListItem' },
            Title: name,
            Description: description
        };
        return await this.post(`/getByTitle('${this.lists.pipelines}')/items`, data);
    }
    
    // STAGES
    async getStagesByPipeline(pipelineId) {
        return await this.get(
            `/getByTitle('${this.lists.stages}')/items`,
            'ID,Title,StageOrder,PipelineId',
            `PipelineId eq ${pipelineId}`,
            '',
            'StageOrder'
        );
    }
    
    async createStage(pipelineId, stageName, order) {
        const data = {
            __metadata: { type: 'SP.Data.CRM_StagesListItem' },
            Title: stageName,
            PipelineId: pipelineId,
            StageOrder: order
        };
        return await this.post(`/getByTitle('${this.lists.stages}')/items`, data);
    }
    
    // BOXES
    async getBoxesByPipeline(pipelineId) {
        return await this.get(
            `/getByTitle('${this.lists.boxes}')/items`,
            'ID,Title,PipelineId,StageId,BoxValue,ContactEmail,ContactName,Notes,Created,Modified',
            `PipelineId eq ${pipelineId}`,
            '',
            'Modified desc'
        );
    }
    
    async getBoxById(boxId) {
        return await this.get(
            `/getByTitle('${this.lists.boxes}')/items(${boxId})`,
            'ID,Title,PipelineId,StageId,BoxValue,ContactEmail,ContactName,Notes,Created,Modified'
        );
    }
    
    async createBox(boxData) {
        const data = {
            __metadata: { type: 'SP.Data.CRM_BoxesListItem' },
            Title: boxData.name,
            PipelineId: boxData.pipelineId,
            StageId: boxData.stageId,
            BoxValue: boxData.value || 0,
            ContactEmail: boxData.contactEmail || '',
            ContactName: boxData.contactName || '',
            Notes: boxData.notes || ''
        };
        return await this.post(`/getByTitle('${this.lists.boxes}')/items`, data);
    }
    
    async updateBox(boxId, boxData) {
        const data = {
            __metadata: { type: 'SP.Data.CRM_BoxesListItem' }
        };
        
        if (boxData.name) data.Title = boxData.name;
        if (boxData.stageId) data.StageId = boxData.stageId;
        if (boxData.value !== undefined) data.BoxValue = boxData.value;
        if (boxData.contactEmail) data.ContactEmail = boxData.contactEmail;
        if (boxData.contactName) data.ContactName = boxData.contactName;
        if (boxData.notes !== undefined) data.Notes = boxData.notes;
        
        return await this.update(`/getByTitle('${this.lists.boxes}')`, boxId, data);
    }
    
    // EMAILS
    async getEmailsByBox(boxId) {
        return await this.get(
            `/getByTitle('${this.lists.emails}')/items`,
            'ID,Title,EmailSubject,EmailFrom,EmailTo,EmailDate,EmailMessageId,BoxId,Created',
            `BoxId eq ${boxId}`,
            '',
            'EmailDate desc'
        );
    }
    
    async getBoxesByEmailMessageId(messageId) {
        // Get all emails with this message ID
        const emails = await this.get(
            `/getByTitle('${this.lists.emails}')/items`,
            'ID,BoxId',
            `EmailMessageId eq '${messageId}'`
        );
        
        if (!emails || emails.length === 0) return [];
        
        // Get unique box IDs
        const boxIds = [...new Set(emails.map(e => e.BoxId))];
        
        // Fetch box details for each ID
        const boxes = [];
        for (const boxId of boxIds) {
            try {
                const box = await this.getBoxById(boxId);
                boxes.push(box);
            } catch (error) {
                console.error(`Error fetching box ${boxId}:`, error);
            }
        }
        
        return boxes;
    }
    
    async linkEmailToBox(emailData, boxId) {
        const data = {
            __metadata: { type: 'SP.Data.CRM_EmailsListItem' },
            Title: emailData.subject.substring(0, 255), // SharePoint title limit
            EmailSubject: emailData.subject,
            EmailFrom: emailData.from,
            EmailTo: emailData.to,
            EmailDate: emailData.date,
            EmailMessageId: emailData.messageId,
            BoxId: boxId
        };
        
        return await this.post(`/getByTitle('${this.lists.emails}')/items`, data);
    }
    
    // ACTIVITIES
    async getActivitiesByBox(boxId) {
        return await this.get(
            `/getByTitle('${this.lists.activities}')/items`,
            'ID,Title,ActivityType,ActivityText,BoxId,Created,Author/Title',
            `BoxId eq ${boxId}`,
            'Author',
            'Created desc'
        );
    }
    
    async getRecentActivities(limit = 20) {
        return await this.get(
            `/getByTitle('${this.lists.activities}')/items`,
            'ID,Title,ActivityType,ActivityText,BoxId,Created,Author/Title',
            '',
            'Author',
            'Created desc'
        ).then(results => results.slice(0, limit));
    }
    
    async createActivity(boxId, activityType, activityText) {
        const data = {
            __metadata: { type: 'SP.Data.CRM_ActivitiesListItem' },
            Title: activityText.substring(0, 255),
            ActivityType: activityType,
            ActivityText: activityText,
            BoxId: boxId
        };
        
        return await this.post(`/getByTitle('${this.lists.activities}')/items`, data);
    }
    
    // SETUP - Create lists if they don't exist
    async setupLists() {
        try {
            // This is a helper function for initial setup
            // In production, you would create these lists manually or via PowerShell
            console.log('Lists should be created in SharePoint with the following structure:');
            console.log(`
            1. ${this.lists.pipelines}
               - Title (Single line text)
               - Description (Multiple lines text)
            
            2. ${this.lists.stages}
               - Title (Single line text)
               - PipelineId (Number)
               - StageOrder (Number)
            
            3. ${this.lists.boxes}
               - Title (Single line text)
               - PipelineId (Number)
               - StageId (Number)
               - BoxValue (Currency)
               - ContactEmail (Single line text)
               - ContactName (Single line text)
               - Notes (Multiple lines text)
            
            4. ${this.lists.emails}
               - Title (Single line text)
               - EmailSubject (Multiple lines text)
               - EmailFrom (Single line text)
               - EmailTo (Multiple lines text)
               - EmailDate (Date and Time)
               - EmailMessageId (Single line text)
               - BoxId (Number)
            
            5. ${this.lists.activities}
               - Title (Single line text)
               - ActivityType (Choice: Email, Note, Stage Change, Created, Updated)
               - ActivityText (Multiple lines text)
               - BoxId (Number)
            `);
        } catch (error) {
            console.error('Setup error:', error);
        }
    }
}

// Export for use in app.js
window.SharePointClient = SharePointClient;
