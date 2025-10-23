// Google Sheets Integration for Sales KPI Dashboard
class GoogleSheetsIntegration {
    constructor() {
        this.apiKey = null;
        this.sheetId = null;
        this.baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
        this.isInitialized = false;
    }

    // Initialize with API key and sheet ID
    async initialize(apiKey, sheetId) {
        this.apiKey = apiKey;
        this.sheetId = sheetId;
        this.isInitialized = true;
        
        try {
            await this.testConnection();
            return true;
        } catch (error) {
            console.error('Failed to initialize Google Sheets integration:', error);
            return false;
        }
    }

    // Test connection to Google Sheets
    async testConnection() {
        const url = `${this.baseUrl}/${this.sheetId}?key=${this.apiKey}`;
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`Failed to connect to Google Sheets: ${response.statusText}`);
        }
        
        return await response.json();
    }

    // Get sheet metadata
    async getSheetInfo() {
        if (!this.isInitialized) {
            throw new Error('Google Sheets integration not initialized');
        }

        const url = `${this.baseUrl}/${this.sheetId}?key=${this.apiKey}`;
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`Failed to get sheet info: ${response.statusText}`);
        }
        
        return await response.json();
    }

    // Read data from a specific sheet
    async readSheetData(sheetName, range = 'A:Z') {
        if (!this.isInitialized) {
            throw new Error('Google Sheets integration not initialized');
        }

        const url = `${this.baseUrl}/${this.sheetId}/values/${sheetName}!${range}?key=${this.apiKey}`;
        const response = await fetch(url);
        
        if (!response.ok) {
            throw new Error(`Failed to read sheet data: ${response.statusText}`);
        }
        
        const data = await response.json();
        return this.parseSheetData(data.values || []);
    }

    // Parse raw sheet data into structured format
    parseSheetData(rawData) {
        if (rawData.length === 0) return [];

        const headers = rawData[0];
        const rows = rawData.slice(1);
        
        return rows.map(row => {
            const obj = {};
            headers.forEach((header, index) => {
                obj[header.toLowerCase().replace(/\s+/g, '_')] = row[index] || '';
            });
            return obj;
        });
    }

    // Get deals data for KPI calculations
    async getDealsData() {
        try {
            const data = await this.readSheetData('Deals', 'A:Z');
            return data.map(row => ({
                id: row.deal_id || row.id,
                closeDate: row.close_date ? new Date(row.close_date) : null,
                value: parseFloat(row.deal_value || row.value || 0),
                teamMember: row.team_member || row.assigned_to,
                stage: row.stage || row.status,
                type: 'closed'
            })).filter(deal => deal.closeDate);
        } catch (error) {
            console.error('Error fetching deals data:', error);
            return [];
        }
    }

    // Get meetings data
    async getMeetingsData() {
        try {
            const data = await this.readSheetData('Meetings', 'A:Z');
            return data.map(row => ({
                id: row.meeting_id || row.id,
                meetingDate: row.meeting_date ? new Date(row.meeting_date) : null,
                status: row.status || 'scheduled',
                teamMember: row.team_member || row.assigned_to,
                type: 'meeting'
            })).filter(meeting => meeting.meetingDate);
        } catch (error) {
            console.error('Error fetching meetings data:', error);
            return [];
        }
    }

    // Get tasks data
    async getTasksData() {
        try {
            const data = await this.readSheetData('Tasks', 'A:Z');
            return data.map(row => ({
                id: row.task_id || row.id,
                dealId: row.deal_id,
                hasTask: row.has_task === 'TRUE' || row.has_task === 'true' || row.has_task === true,
                taskDate: row.task_date ? new Date(row.task_date) : null,
                teamMember: row.team_member || row.assigned_to
            }));
        } catch (error) {
            console.error('Error fetching tasks data:', error);
            return [];
        }
    }

    // Get limbo deals data
    async getLimboData() {
        try {
            const data = await this.readSheetData('Limbo', 'A:Z');
            return data.map(row => ({
                id: row.limbo_id || row.id,
                dealId: row.deal_id,
                limboDate: row.limbo_date ? new Date(row.limbo_date) : null,
                movedDate: row.moved_date ? new Date(row.moved_date) : null,
                teamMember: row.team_member || row.assigned_to
            }));
        } catch (error) {
            console.error('Error fetching limbo data:', error);
            return [];
        }
    }

    // Get outbound meetings data
    async getOutboundData() {
        try {
            const data = await this.readSheetData('Outbound', 'A:Z');
            return data.map(row => ({
                id: row.outbound_id || row.id,
                outboundDate: row.outbound_date ? new Date(row.outbound_date) : null,
                teamMember: row.team_member || row.assigned_to,
                type: 'outbound'
            })).filter(outbound => outbound.outboundDate);
        } catch (error) {
            console.error('Error fetching outbound data:', error);
            return [];
        }
    }

    // Get new subscriptions data
    async getSubscriptionsData() {
        try {
            const data = await this.readSheetData('Subscriptions', 'A:Z');
            return data.map(row => ({
                id: row.sub_id || row.id,
                firstMeetingDate: row.first_meeting_date ? new Date(row.first_meeting_date) : null,
                calledWithin24h: row.called_within_24h === 'TRUE' || row.called_within_24h === 'true' || row.called_within_24h === true,
                teamMember: row.team_member || row.assigned_to
            }));
        } catch (error) {
            console.error('Error fetching subscriptions data:', error);
            return [];
        }
    }

    // Get all data for dashboard
    async getAllData() {
        try {
            const [deals, meetings, tasks, limbo, outbound, subscriptions] = await Promise.all([
                this.getDealsData(),
                this.getMeetingsData(),
                this.getTasksData(),
                this.getLimboData(),
                this.getOutboundData(),
                this.getSubscriptionsData()
            ]);

            return {
                deals,
                meetings,
                tasks,
                limbo,
                outbound,
                subscriptions,
                lastUpdated: new Date()
            };
        } catch (error) {
            console.error('Error fetching all data:', error);
            return {
                deals: [],
                meetings: [],
                tasks: [],
                limbo: [],
                outbound: [],
                subscriptions: [],
                lastUpdated: new Date()
            };
        }
    }

    // Write data back to sheet (for future use)
    async writeData(sheetName, data) {
        if (!this.isInitialized) {
            throw new Error('Google Sheets integration not initialized');
        }

        // This would require additional authentication for write access
        console.log('Write functionality requires additional authentication setup');
        return false;
    }
}

// Google Sheets Setup Instructions
const GOOGLE_SHEETS_SETUP = {
    instructions: `
    To connect your Google Sheets data to the dashboard:
    
    1. Create a Google Sheet with the following tabs:
       - "Deals" - Columns: Deal_ID, Close_Date, Deal_Value, Team_Member, Stage
       - "Meetings" - Columns: Meeting_ID, Meeting_Date, Status, Team_Member
       - "Tasks" - Columns: Task_ID, Deal_ID, Has_Task, Task_Date, Team_Member
       - "Limbo" - Columns: Limbo_ID, Deal_ID, Limbo_Date, Moved_Date, Team_Member
       - "Outbound" - Columns: Outbound_ID, Outbound_Date, Team_Member
       - "Subscriptions" - Columns: Sub_ID, First_Meeting_Date, Called_Within_24h, Team_Member
    
    2. Make the sheet publicly accessible:
       - Go to File > Share > Change to "Anyone with the link can view"
       - Copy the sheet ID from the URL
    
    3. Get a Google Sheets API key:
       - Go to Google Cloud Console
       - Create a new project or select existing
       - Enable Google Sheets API
       - Create credentials (API Key)
       - Restrict the key to Google Sheets API
    
    4. Enter the API key and Sheet ID in the dashboard
    `,
    
    sampleData: {
        deals: [
            ['Deal_ID', 'Close_Date', 'Deal_Value', 'Team_Member', 'Stage'],
            ['D001', '2024-01-15', '5000', 'John Doe', 'Closed Won'],
            ['D002', '2024-01-20', '7500', 'Jane Smith', 'Closed Won']
        ],
        meetings: [
            ['Meeting_ID', 'Meeting_Date', 'Status', 'Team_Member'],
            ['M001', '2024-01-10', 'Completed', 'John Doe'],
            ['M002', '2024-01-12', 'Missed', 'Jane Smith']
        ]
    }
};

// Export for use in main dashboard
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { GoogleSheetsIntegration, GOOGLE_SHEETS_SETUP };
}
