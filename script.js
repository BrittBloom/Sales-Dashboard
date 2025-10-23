// Sales KPI Dashboard JavaScript
console.log('Script.js loaded successfully');

class SalesKPIDashboard {
    constructor() {
        console.log('SalesKPIDashboard constructor called');
        this.data = [];
        this.lastUpdated = null;
        this.sheetId = null;
        this.currentFilter = 'all';
        this.accountExecutives = [];
        this.dateRange = {
            startDate: null,
            endDate: null
        };
        
        this.init();
    }

    async init() {
        try {
            console.log('Setting up event listeners...');
            this.setupEventListeners();
            
            console.log('Loading data...');
            await this.loadRealData();
            
            console.log('Initializing date range...');
            this.initializeMonthSelector(); // Initialize month selector on load
            
            console.log('Data loaded:', this.data.length, 'deals');
            console.log('Date range:', this.dateRange);
            
            console.log('Updating dashboard...');
            this.updateDashboard();
            
            console.log('Starting auto refresh...');
            this.startAutoRefresh();
            
            console.log('Dashboard initialization complete');
        } catch (error) {
            console.error('Error in dashboard initialization:', error);
            throw error;
        }
    }

    setupEventListeners() {
        console.log('Setting up event listeners...');
        
        // Refresh button
        const refreshBtn = document.getElementById('refreshBtn');
        console.log('Refresh button found:', refreshBtn);
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => {
                this.refreshData();
            });
        } else {
            console.error('Refresh button not found!');
        }

        // Export button
        const exportBtn = document.getElementById('exportBtn');
        console.log('Export button found:', exportBtn);
        if (exportBtn) {
            exportBtn.addEventListener('click', () => {
                this.exportReport();
            });
        } else {
            console.error('Export button not found!');
        }

        // Config button removed

        // Account Executive filter
        const aeFilter = document.getElementById('aeFilter');
        console.log('AE Filter found:', aeFilter);
        if (aeFilter) {
            aeFilter.addEventListener('change', (event) => {
                this.currentFilter = event.target.value;
                console.log('AE Filter changed to:', this.currentFilter);
                this.updateDashboard();
            });
        } else {
            console.error('AE Filter not found!');
        }

        // Month selector changes
        const monthSelector = document.getElementById('monthSelector');
        
        if (monthSelector) {
            monthSelector.addEventListener('change', () => {
                console.log('Month selector changed to:', monthSelector.value);
                this.updateDateRange();
            });
        } else {
            console.error('Month selector not found!');
        }

        // Google Sheets modal removed
    }

    // Load real data from Google Sheets
    async loadRealData() {
        try {
            // Check if Google Sheets is configured
            const apiKey = localStorage.getItem('googleSheetsApiKey');
            const sheetId = localStorage.getItem('googleSheetsId');
            const useSampleData = localStorage.getItem('useSampleData') === 'true';
            
            console.log('Data loading configuration:');
            console.log('- API Key:', apiKey ? 'Present' : 'Missing');
            console.log('- Sheet ID:', sheetId || 'Using default');
            console.log('- Use Sample Data:', useSampleData);
            
            // Default to your Google Sheet if no configuration
            const defaultSheetId = '1xzobbUqsFFc9WlP027XoUYOzG1ZUfjrZJaMERfuUbyk';
            const finalSheetId = sheetId || defaultSheetId;
            
            if (useSampleData) {
                console.log('Using sample data as requested');
                this.loadSampleDataWithRealStructure();
            } else if (apiKey) {
                console.log('Loading from Google Sheets:', finalSheetId);
                await this.loadFromGoogleSheets(apiKey, finalSheetId);
            } else {
                console.log('No API key found, using sample data');
                this.loadSampleDataWithRealStructure();
            }
            
        } catch (error) {
            console.error('Error loading real data:', error);
            console.error('Error details:', error.message);
            // Fallback to sample data
            this.loadSampleDataWithRealStructure();
        }
    }

    // Load data from Google Sheets
    async loadFromGoogleSheets(apiKey, sheetId) {
        try {
            console.log('Starting Google Sheets data load...');
            const baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
            const url = `${baseUrl}/${sheetId}/values/A:Z?key=${apiKey}`;
            
            console.log('Fetching from URL:', url);
            const response = await fetch(url);
            
            console.log('Response status:', response.status);
            console.log('Response ok:', response.ok);
            
            if (!response.ok) {
                const errorText = await response.text();
                console.error('Response error:', errorText);
                throw new Error(`Failed to fetch data: ${response.statusText} - ${errorText}`);
            }
            
            const data = await response.json();
            console.log('Raw Google Sheets response:', data);
            
            const rows = data.values || [];
            console.log('Total rows from sheet:', rows.length);
            
            if (rows.length < 4) {
                throw new Error('Not enough data in sheet');
            }
            
            // Show first few rows to understand structure
            console.log('First 5 rows from sheet:');
            rows.slice(0, 5).forEach((row, index) => {
                console.log(`Row ${index}:`, row);
            });
            
            // Skip header rows (first 3 rows) and parse data
            const dataRows = rows.slice(3);
            console.log('Data rows (after skipping headers):', dataRows.length);
            
            this.data = [];
            
            dataRows.forEach((row, index) => {
                console.log(`Processing row ${index}:`, row);
                
                // Skip empty rows or rows with insufficient data
                if (row.length >= 6 && row[0] && row[0].trim() !== '') {
                    const deal = this.parseGoogleSheetRow(row);
                    if (deal) {
                        this.data.push(deal);
                        console.log(`Added deal: ${deal.id} - ${deal.name}`);
                    } else {
                        console.log(`Failed to parse row ${index}:`, row);
                    }
                } else {
                    console.log(`Skipping row ${index} - insufficient data:`, row);
                }
            });
            
            console.log('=== GOOGLE SHEETS DATA LOADED ===');
            console.log('Total deals loaded:', this.data.length);
            console.log('Sample deals:', this.data.slice(0, 3));
            
            // Debug: Show deals by stage
            const dealsByStage = {};
            this.data.forEach(deal => {
                const stage = deal.stage || 'Unknown';
                dealsByStage[stage] = (dealsByStage[stage] || 0) + 1;
            });
            console.log('Deals by stage:', dealsByStage);
            
            // Debug: Show closed deals specifically
            const closedDeals = this.data.filter(deal => deal.stage === 'Closed Won');
            console.log('Closed Won deals found:', closedDeals.length);
            
            // Show detailed info about closed deals
            closedDeals.forEach((deal, index) => {
                console.log(`Closed Deal ${index + 1}: ID=${deal.id}, CloseDate=${deal.closeDate}, Owner=${deal.teamMember}`);
            });
            
            // Debug: Show date range of all deals
            const allDates = this.data.map(deal => deal.closeDate).filter(d => d);
            if (allDates.length > 0) {
                const minDate = new Date(Math.min(...allDates.map(d => new Date(d).getTime())));
                const maxDate = new Date(Math.max(...allDates.map(d => new Date(d).getTime())));
                console.log('Date range of all close dates:', minDate.toLocaleDateString(), 'to', maxDate.toLocaleDateString());
            }
            
            console.log('=== END GOOGLE SHEETS DATA ===');
            
            // Extract unique account executives
            this.accountExecutives = [...new Set(this.data.map(deal => deal.teamMember))];
            console.log('Account executives found:', this.accountExecutives);
            
            this.lastUpdated = new Date();
            this.populateAccountExecutiveFilter();
            
        } catch (error) {
            console.error('Error loading from Google Sheets:', error);
            throw error;
        }
    }

    // Parse a single row from Google Sheets
    parseGoogleSheetRow(row) {
        try {
            console.log('Parsing row:', row);
            
            const dealId = row[0] || '';
            const amount = parseFloat(row[1]) || 0;
            const owner = row[2] || '';
            const stage = row[3] || '';
            const name = row[4] || '';
            const createDate = row[5] ? new Date(row[5]) : new Date();
            const closeDate = row[6] ? new Date(row[6]) : null;
            
            console.log(`Deal ${dealId}: ${name} - Stage: ${stage} - Owner: ${owner}`);
            
            // Parse stage entry dates
            const coldAppInstallDate = row[7] ? new Date(row[7]) : null;
            const meetingBookedDate = row[8] ? new Date(row[8]) : null;
            const missedMeetingDate = row[9] ? new Date(row[9]) : null;
            const limboDate = row[10] ? new Date(row[10]) : null;
            const onboardingBookedDate = row[11] ? new Date(row[11]) : null;
            const subscribedDate = row[12] ? new Date(row[12]) : null;
            const activateDate = row[13] ? new Date(row[13]) : null;
            const closedWonDate = row[14] ? new Date(row[14]) : null;
            const closedLostDate = row[15] ? new Date(row[15]) : null;
            
            // Debug: Log the raw row data for the first few rows
            if (dealId && dealId.includes('D')) {
                console.log(`Raw row data for ${dealId}:`, {
                    rowLength: row.length,
                    meetingBookedRaw: row[8],
                    missedMeetingRaw: row[9],
                    limboRaw: row[10]
                });
            }
            
            // Parse additional columns for new calculations
            const meetingBookedExitDate = row[24] ? new Date(row[24]) : null; // Column Y
            const nextStepDate = row[25] ? new Date(row[25]) : null; // Column Z
            const limboExitDate = row[26] ? new Date(row[26]) : null; // Column AA
            
            // Determine current stage date and last contact
            let stageDate = createDate;
            let lastContactDate = createDate;
            let nextActivityDate = null;
            
            // Set stage date based on current stage
            switch (stage) {
                case 'Cold App Install':
                    stageDate = coldAppInstallDate || createDate;
                    break;
                case 'Meeting Booked':
                    stageDate = meetingBookedDate || createDate;
                    break;
                case 'Missed Meeting':
                    stageDate = missedMeetingDate || createDate;
                    break;
                case 'Limbo':
                    stageDate = limboDate || createDate;
                    break;
                case 'Onboarding Booked':
                    stageDate = onboardingBookedDate || createDate;
                    break;
                case 'SUBSCRIBED':
                    stageDate = subscribedDate || createDate;
                    break;
                case 'Activate':
                    stageDate = activateDate || createDate;
                    break;
                case 'Closed Won':
                    stageDate = closedWonDate || createDate;
                    break;
                case 'Closed Lost':
                    stageDate = closedLostDate || createDate;
                    break;
            }
            
            // Set last contact date (most recent activity)
            const dates = [createDate, coldAppInstallDate, meetingBookedDate, missedMeetingDate, 
                          limboDate, onboardingBookedDate, subscribedDate, activateDate, 
                          closedWonDate, closedLostDate].filter(d => d);
            if (dates.length > 0) {
                lastContactDate = new Date(Math.max(...dates.map(d => d.getTime())));
            }
            
            const deal = {
                id: dealId,
                name: name,
                stage: stage,
                stageDate: stageDate,
                createDate: createDate,
                lastContactDate: lastContactDate,
                nextActivityDate: nextActivityDate,
                value: amount,
                teamMember: owner,
                type: 'active',
                closeDate: closeDate,
                meetingDate: meetingBookedDate,
                taskStatus: 'Completed',
                limboDate: limboDate,
                limboExitDate: limboExitDate, // Column AA
                outboundMeetingDate: null,
                firstMeetingDate: meetingBookedDate,
                missedMeetingDate: missedMeetingDate,
                meetingBookedDate: meetingBookedDate, // Add this field explicitly
                meetingBookedExitDate: meetingBookedExitDate, // Column Y
                nextStepDate: nextStepDate // Column Z
            };
            
            // Debug: Log the parsed deal data for the first few deals
            if (dealId && dealId.includes('D') && dealId.length < 10) {
                console.log(`Parsed deal data for ${dealId}:`, {
                    meetingBookedDate: deal.meetingBookedDate,
                    missedMeetingDate: deal.missedMeetingDate,
                    stage: deal.stage
                });
            }
            
            return deal;
        } catch (error) {
            console.error('Error parsing Google Sheets row:', error);
            return null;
        }
    }

    // Create sample data that matches the real CSV structure
    loadSampleDataWithRealStructure() {
        const currentDate = new Date();
        const currentMonth = currentDate.getMonth();
        const currentYear = currentDate.getFullYear();
        
        // Use dates from the past to avoid negative day calculations
        const baseDate = new Date(currentYear, currentMonth - 1, 15);
        const daysAgo = (days) => new Date(baseDate.getTime() - days * 24 * 60 * 60 * 1000);
        
        // Define account executives from real data
        this.accountExecutives = ['Tom', 'Eddy', 'Galina'];
        
        this.data = [
            // Tom's deals
            {
                id: '45335194478',
                name: 'Inbound Lead Paul Mardeys',
                stage: 'Meeting Booked',
                stageDate: daysAgo(3),
                createDate: daysAgo(3), // Create date for filtering
                lastContactDate: daysAgo(1),
                nextActivityDate: daysAgo(-2), // Future date
                value: 1,
                teamMember: 'Tom',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(3),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(3)
            },
            {
                id: '45315568809',
                name: 'Inbound Lead Artion',
                stage: 'Activate',
                stageDate: daysAgo(7),
                createDate: daysAgo(7), // Add createDate for filtering
                lastContactDate: daysAgo(1),
                nextActivityDate: null,
                value: 20,
                teamMember: 'Tom',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(7),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(7)
            },
            {
                id: '44968061999',
                name: 'DGP New Deal',
                stage: 'Closed Won',
                stageDate: daysAgo(10),
                createDate: daysAgo(10), // Create date for filtering
                lastContactDate: daysAgo(10),
                nextActivityDate: null,
                value: 600,
                teamMember: 'Tom',
                type: 'active',
                closeDate: daysAgo(10),
                meetingDate: daysAgo(12),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(12)
            },
            // Eddy's deals
            {
                id: '45367198992',
                name: 'Inbound Lead John',
                stage: 'Qualified',
                stageDate: daysAgo(3),
                createDate: daysAgo(3), // Add createDate for filtering
                lastContactDate: daysAgo(2),
                nextActivityDate: daysAgo(-1), // Future date
                value: 1,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: null
            },
            {
                id: '45137411084',
                name: 'Win Back Melissa Marks',
                stage: 'Closed Won',
                stageDate: daysAgo(7),
                createDate: daysAgo(7), // Add createDate for filtering
                lastContactDate: daysAgo(7),
                nextActivityDate: null,
                value: 1,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: daysAgo(7),
                meetingDate: daysAgo(8),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(8)
            },
            {
                id: '45182979959',
                name: 'Inbound Lead Saurav',
                stage: 'SUBSCRIBED',
                stageDate: daysAgo(5),
                createDate: daysAgo(5), // Add createDate for filtering
                lastContactDate: daysAgo(4),
                nextActivityDate: null,
                value: 800,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(5),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(5)
            },
            // Galina's deals
            {
                id: '45206261043',
                name: 'Work Wear Choice New Deal',
                stage: 'Cold App Install',
                stageDate: daysAgo(5),
                createDate: daysAgo(5), // Add createDate for filtering
                lastContactDate: daysAgo(4),
                nextActivityDate: daysAgo(-1), // Future date
                value: 1,
                teamMember: 'Galina',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: null
            },
            {
                id: '45149797001',
                name: 'Inbound Lead Stacey',
                stage: 'Missed Meeting',
                stageDate: daysAgo(6),
                createDate: daysAgo(6), // Add createDate for filtering
                lastContactDate: daysAgo(3),
                nextActivityDate: daysAgo(-1), // Future date
                value: 1,
                teamMember: 'Galina',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(6),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(6),
                missedMeetingDate: daysAgo(6) // Add missed meeting date
            },
            {
                id: '45088411724',
                name: 'Inbound Lead John',
                stage: 'Limbo',
                stageDate: daysAgo(8),
                createDate: daysAgo(8), // Add createDate for filtering
                lastContactDate: daysAgo(6),
                nextActivityDate: null,
                value: 200,
                teamMember: 'Galina',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(8),
                taskStatus: 'Completed',
                limboDate: daysAgo(6),
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(8)
            },
            // Additional deals for better KPI calculations
            {
                id: '45176356139',
                name: 'Inbound Lead Mohamed',
                stage: 'Meeting Booked',
                stageDate: daysAgo(6),
                createDate: daysAgo(6), // Add createDate for filtering
                lastContactDate: daysAgo(5),
                nextActivityDate: daysAgo(-2), // Future date
                value: 1,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(6),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(6)
            },
            {
                id: '45151577264',
                name: 'Inbound Lead Johnathan',
                stage: 'Meeting Booked',
                stageDate: daysAgo(7),
                createDate: daysAgo(7), // Add createDate for filtering
                lastContactDate: daysAgo(6),
                nextActivityDate: daysAgo(-1), // Future date
                value: 1,
                teamMember: 'Tom',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(7),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(7)
            },
            {
                id: '45122184610',
                name: 'Inbound Lead krunal',
                stage: 'Missed Meeting',
                stageDate: daysAgo(8),
                createDate: daysAgo(8), // Add createDate for filtering
                lastContactDate: daysAgo(5),
                nextActivityDate: daysAgo(-1), // Future date
                value: 1,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(8),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(8),
                missedMeetingDate: daysAgo(8) // Add missed meeting date
            },
            // Additional active deals to populate the table
            {
                id: '45235987997',
                name: 'Sitandpawz.com New Deal',
                stage: 'Cold App Install',
                stageDate: daysAgo(5),
                createDate: daysAgo(5), // Add createDate for filtering
                lastContactDate: daysAgo(4),
                nextActivityDate: daysAgo(-1),
                value: 1,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: null
            },
            {
                id: '45208732572',
                name: 'Inbound Lead Edward',
                stage: 'Meeting Booked',
                stageDate: daysAgo(6),
                createDate: daysAgo(6), // Add createDate for filtering
                lastContactDate: daysAgo(5),
                nextActivityDate: daysAgo(-1),
                value: 1,
                teamMember: 'Tom',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(6),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(6)
            },
            {
                id: '45223049022',
                name: 'Inbound Lead Alex',
                stage: 'Meeting Booked',
                stageDate: daysAgo(6),
                createDate: daysAgo(6), // Add createDate for filtering
                lastContactDate: daysAgo(5),
                nextActivityDate: daysAgo(-1),
                value: 1,
                teamMember: 'Galina',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(6),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(6)
            },
            {
                id: '45091578509',
                name: 'artsabers.com New Deal',
                stage: 'Qualified',
                stageDate: daysAgo(8),
                createDate: daysAgo(8), // Add createDate for filtering
                lastContactDate: daysAgo(7),
                nextActivityDate: daysAgo(-1),
                value: 1,
                teamMember: 'Tom',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: null
            },
            {
                id: '45082859625',
                name: 'Shop CAFE LIGHTING & LIVING New Deal',
                stage: 'Onboarding Booked',
                stageDate: daysAgo(8),
                createDate: daysAgo(8), // Add createDate for filtering
                lastContactDate: daysAgo(7),
                nextActivityDate: daysAgo(-1),
                value: 300,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(8),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(8)
            },
            {
                id: '45002512703',
                name: 'Ancestral Nutrition - New Deal',
                stage: 'Activate',
                stageDate: daysAgo(9),
                createDate: daysAgo(9), // Add createDate for filtering
                lastContactDate: daysAgo(8),
                nextActivityDate: null,
                value: 1,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(9),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(9)
            },
            {
                id: '44995692827',
                name: 'The Christian Art Store New Deal',
                stage: 'Qualified',
                stageDate: daysAgo(9),
                createDate: daysAgo(9), // Add createDate for filtering
                lastContactDate: daysAgo(8),
                nextActivityDate: daysAgo(-1),
                value: 1,
                teamMember: 'Tom',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: null
            },
            {
                id: '45006609015',
                name: 'Sunflare Xplor New Deal',
                stage: 'Cold App Install',
                stageDate: daysAgo(9),
                createDate: daysAgo(9), // Add createDate for filtering
                lastContactDate: daysAgo(8),
                nextActivityDate: daysAgo(-1),
                value: 1,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: null
            },
            {
                id: '44969753762',
                name: 'Inbound Lead Kurt',
                stage: 'Meeting Booked',
                stageDate: daysAgo(10),
                createDate: daysAgo(10), // Add createDate for filtering
                lastContactDate: daysAgo(9),
                nextActivityDate: daysAgo(-1),
                value: 700,
                teamMember: 'Eddy',
                type: 'active',
                closeDate: null,
                meetingDate: daysAgo(10),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(10)
            }
        ];
        
        this.lastUpdated = new Date();
        this.populateAccountExecutiveFilter();
    }

    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        result.push(current.trim());
        return result;
    }

    parseDealData(columns) {
        try {
            const dealId = columns[0];
            const amount = parseFloat(columns[1]) || 0;
            const owner = columns[2];
            const stage = columns[3];
            const name = columns[4];
            const createDate = new Date(columns[5]);
            const closeDate = columns[6] ? new Date(columns[6]) : null;
            
            // Parse stage entry dates
            const coldAppInstallDate = columns[7] ? new Date(columns[7]) : null;
            const meetingBookedDate = columns[8] ? new Date(columns[8]) : null;
            const missedMeetingDate = columns[9] ? new Date(columns[9]) : null;
            const limboDate = columns[10] ? new Date(columns[10]) : null;
            const onboardingBookedDate = columns[11] ? new Date(columns[11]) : null;
            const subscribedDate = columns[12] ? new Date(columns[12]) : null;
            const activateDate = columns[13] ? new Date(columns[13]) : null;
            const closedWonDate = columns[14] ? new Date(columns[14]) : null;
            const closedLostDate = columns[15] ? new Date(columns[15]) : null;
            
            // Determine current stage date and last contact
            let stageDate = createDate;
            let lastContactDate = createDate;
            let nextActivityDate = null;
            
            // Set stage date based on current stage
            switch (stage) {
                case 'Cold App Install':
                    stageDate = coldAppInstallDate || createDate;
                    break;
                case 'Meeting Booked':
                    stageDate = meetingBookedDate || createDate;
                    break;
                case 'Missed Meeting':
                    stageDate = missedMeetingDate || createDate;
                    break;
                case 'Limbo':
                    stageDate = limboDate || createDate;
                    break;
                case 'Onboarding Booked':
                    stageDate = onboardingBookedDate || createDate;
                    break;
                case 'SUBSCRIBED':
                    stageDate = subscribedDate || createDate;
                    break;
                case 'Activate':
                    stageDate = activateDate || createDate;
                    break;
                case 'Closed Won':
                    stageDate = closedWonDate || createDate;
                    break;
                case 'Closed Lost':
                    stageDate = closedLostDate || createDate;
                    break;
            }
            
            // Set last contact date (most recent activity)
            const dates = [createDate, coldAppInstallDate, meetingBookedDate, missedMeetingDate, 
                          limboDate, onboardingBookedDate, subscribedDate, activateDate, 
                          closedWonDate, closedLostDate].filter(d => d);
            if (dates.length > 0) {
                lastContactDate = new Date(Math.max(...dates.map(d => d.getTime())));
            }
            
            return {
                id: dealId,
                name: name,
                stage: stage,
                stageDate: stageDate,
                lastContactDate: lastContactDate,
                nextActivityDate: nextActivityDate,
                value: amount,
                teamMember: owner,
                type: 'active',
                closeDate: closeDate,
                meetingDate: meetingBookedDate,
                taskStatus: 'Completed', // Assume completed for real data
                limboDate: limboDate,
                outboundMeetingDate: null, // Not available in this data
                firstMeetingDate: meetingBookedDate
            };
        } catch (error) {
            console.error('Error parsing deal data:', error);
            return null;
        }
    }

    // Enhanced sample data matching typical sales KPI report structure
    loadSampleData() {
        const currentDate = new Date();
        const currentMonth = currentDate.getMonth();
        const currentYear = currentDate.getFullYear();
        
        // Use dates from the past to avoid negative day calculations
        const baseDate = new Date(currentYear, currentMonth - 1, 15); // Use last month as base
        
        // Helper function to create past dates
        const daysAgo = (days) => new Date(baseDate.getTime() - days * 24 * 60 * 60 * 1000);
        
        // Define account executives from real data
        this.accountExecutives = [
            'Tom', 'Eddy', 'Galina'
        ];
        
        this.data = [
            // John Smith's deals
            {
                id: 'D001',
                name: 'Enterprise Software License',
                stage: 'Closed Won',
                stageDate: daysAgo(5),
                lastContactDate: daysAgo(8),
                nextActivityDate: null,
                value: 50000,
                teamMember: 'John Smith',
                type: 'active',
                closeDate: daysAgo(5),
                meetingDate: daysAgo(10),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: daysAgo(10)
            },
            {
                id: 'D002',
                name: 'Cloud Infrastructure',
                stage: 'Closed Won',
                stageDate: new Date(currentYear, currentMonth, 30),
                lastContactDate: new Date(currentYear, currentMonth, 28),
                nextActivityDate: null,
                value: 75000,
                teamMember: 'John Smith',
                type: 'active',
                closeDate: new Date(currentYear, currentMonth, 30),
                meetingDate: new Date(currentYear, currentMonth, 25),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 25)
            },
            {
                id: 'D003',
                name: 'AI Integration',
                stage: 'Meeting Booked',
                stageDate: new Date(currentYear, currentMonth, 28),
                lastContactDate: new Date(currentYear, currentMonth, 26),
                nextActivityDate: new Date(currentYear, currentMonth + 1, 2),
                value: 40000,
                teamMember: 'John Smith',
                type: 'active',
                closeDate: null,
                meetingDate: new Date(currentYear, currentMonth, 28),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 28)
            },
            {
                id: 'D004',
                name: 'Data Analytics Platform',
                stage: 'Qualified',
                stageDate: new Date(currentYear, currentMonth, 20),
                lastContactDate: new Date(currentYear, currentMonth, 18),
                nextActivityDate: new Date(currentYear, currentMonth, 25),
                value: 25000,
                teamMember: 'John Smith',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 20),
                firstMeetingDate: null
            },
            {
                id: 'D005',
                name: 'Mobile App Development',
                stage: 'Cold App Install',
                stageDate: new Date(currentYear, currentMonth, 15),
                lastContactDate: new Date(currentYear, currentMonth, 13),
                nextActivityDate: new Date(currentYear, currentMonth, 22),
                value: 18000,
                teamMember: 'John Smith',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 15),
                firstMeetingDate: null
            },
            // Sarah Johnson's deals
            {
                id: 'D006',
                name: 'CRM Implementation',
                stage: 'Limbo',
                stageDate: new Date(currentYear, currentMonth, 20),
                lastContactDate: new Date(currentYear, currentMonth, 22),
                nextActivityDate: new Date(currentYear, currentMonth, 28),
                value: 30000,
                teamMember: 'Sarah Johnson',
                type: 'active',
                closeDate: null,
                meetingDate: new Date(currentYear, currentMonth, 20),
                taskStatus: 'Pending',
                limboDate: new Date(currentYear, currentMonth, 25),
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 20)
            },
            {
                id: 'D007',
                name: 'Marketing Automation',
                stage: 'Missed Meeting',
                stageDate: new Date(currentYear, currentMonth, 22),
                lastContactDate: new Date(currentYear, currentMonth, 20),
                nextActivityDate: new Date(currentYear, currentMonth, 29),
                value: 35000,
                teamMember: 'Sarah Johnson',
                type: 'active',
                closeDate: null,
                meetingDate: new Date(currentYear, currentMonth, 22),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 22)
            },
            {
                id: 'D008',
                name: 'Security Solution',
                stage: 'Qualified',
                stageDate: new Date(currentYear, currentMonth, 18),
                lastContactDate: new Date(currentYear, currentMonth, 16),
                nextActivityDate: new Date(currentYear, currentMonth, 24),
                value: 22000,
                teamMember: 'Sarah Johnson',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 18),
                firstMeetingDate: null
            },
            {
                id: 'D009',
                name: 'Workflow Automation',
                stage: 'Cold App Install',
                stageDate: new Date(currentYear, currentMonth, 12),
                lastContactDate: new Date(currentYear, currentMonth, 10),
                nextActivityDate: new Date(currentYear, currentMonth, 18),
                value: 15000,
                teamMember: 'Sarah Johnson',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 12),
                firstMeetingDate: null
            },
            // Mike Wilson's deals
            {
                id: 'D010',
                name: 'Enterprise Integration',
                stage: 'Closed Won',
                stageDate: new Date(currentYear, currentMonth, 25),
                lastContactDate: new Date(currentYear, currentMonth, 23),
                nextActivityDate: null,
                value: 60000,
                teamMember: 'Mike Wilson',
                type: 'active',
                closeDate: new Date(currentYear, currentMonth, 25),
                meetingDate: new Date(currentYear, currentMonth, 20),
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 20)
            },
            {
                id: 'D011',
                name: 'Business Intelligence',
                stage: 'Meeting Booked',
                stageDate: new Date(currentYear, currentMonth, 30),
                lastContactDate: new Date(currentYear, currentMonth, 28),
                nextActivityDate: new Date(currentYear, currentMonth + 1, 5),
                value: 45000,
                teamMember: 'Mike Wilson',
                type: 'active',
                closeDate: null,
                meetingDate: new Date(currentYear, currentMonth, 30),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 30)
            },
            {
                id: 'D012',
                name: 'Customer Support Platform',
                stage: 'Limbo',
                stageDate: new Date(currentYear, currentMonth, 15),
                lastContactDate: new Date(currentYear, currentMonth, 13),
                nextActivityDate: new Date(currentYear, currentMonth, 20),
                value: 28000,
                teamMember: 'Mike Wilson',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: new Date(currentYear, currentMonth, 20),
                outboundMeetingDate: new Date(currentYear, currentMonth, 10),
                firstMeetingDate: null
            },
            {
                id: 'D013',
                name: 'API Management',
                stage: 'Missed Meeting',
                stageDate: new Date(currentYear, currentMonth, 18),
                lastContactDate: new Date(currentYear, currentMonth, 16),
                nextActivityDate: new Date(currentYear, currentMonth, 25),
                value: 32000,
                teamMember: 'Mike Wilson',
                type: 'active',
                closeDate: null,
                meetingDate: new Date(currentYear, currentMonth, 18),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 18)
            },
            {
                id: 'D014',
                name: 'Database Migration',
                stage: 'Qualified',
                stageDate: new Date(currentYear, currentMonth, 14),
                lastContactDate: new Date(currentYear, currentMonth, 12),
                nextActivityDate: new Date(currentYear, currentMonth, 19),
                value: 19000,
                teamMember: 'Mike Wilson',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 14),
                firstMeetingDate: null
            },
            // Additional deals for better KPI calculations
            {
                id: 'D015',
                name: 'Email Marketing Tool',
                stage: 'Cold App Install',
                stageDate: new Date(currentYear, currentMonth, 8),
                lastContactDate: new Date(currentYear, currentMonth, 6),
                nextActivityDate: new Date(currentYear, currentMonth, 15),
                value: 12000,
                teamMember: 'John Smith',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 8),
                firstMeetingDate: null
            },
            {
                id: 'D016',
                name: 'E-commerce Platform',
                stage: 'Meeting Booked',
                stageDate: new Date(currentYear, currentMonth, 26),
                lastContactDate: new Date(currentYear, currentMonth, 24),
                nextActivityDate: new Date(currentYear, currentMonth + 1, 1),
                value: 38000,
                teamMember: 'Sarah Johnson',
                type: 'active',
                closeDate: null,
                meetingDate: new Date(currentYear, currentMonth, 26),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 26)
            },
            {
                id: 'D017',
                name: 'Project Management Tool',
                stage: 'Cold App Install',
                stageDate: new Date(currentYear, currentMonth, 5),
                lastContactDate: new Date(currentYear, currentMonth, 3),
                nextActivityDate: new Date(currentYear, currentMonth, 12),
                value: 16000,
                teamMember: 'Mike Wilson',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 5),
                firstMeetingDate: null
            },
            // Emily Davis deals
            {
                id: 'D018',
                name: 'HR Management System',
                stage: 'Inbound Lead',
                stageDate: new Date(currentYear, currentMonth, 10),
                lastContactDate: new Date(currentYear, currentMonth, 12),
                nextActivityDate: new Date(currentYear, currentMonth, 18),
                value: 42000,
                teamMember: 'Emily Davis',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: null
            },
            {
                id: 'D019',
                name: 'Financial Analytics',
                stage: 'Qualified',
                stageDate: new Date(currentYear, currentMonth, 16),
                lastContactDate: new Date(currentYear, currentMonth, 14),
                nextActivityDate: new Date(currentYear, currentMonth, 21),
                value: 33000,
                teamMember: 'Emily Davis',
                type: 'active',
                closeDate: null,
                meetingDate: null,
                taskStatus: 'Completed',
                limboDate: null,
                outboundMeetingDate: new Date(currentYear, currentMonth, 16),
                firstMeetingDate: null
            },
            // Alex Chen deals
            {
                id: 'D020',
                name: 'IoT Platform',
                stage: 'Meeting Booked',
                stageDate: new Date(currentYear, currentMonth, 24),
                lastContactDate: new Date(currentYear, currentMonth, 22),
                nextActivityDate: new Date(currentYear, currentMonth + 1, 3),
                value: 55000,
                teamMember: 'Alex Chen',
                type: 'active',
                closeDate: null,
                meetingDate: new Date(currentYear, currentMonth, 24),
                taskStatus: 'Pending',
                limboDate: null,
                outboundMeetingDate: null,
                firstMeetingDate: new Date(currentYear, currentMonth, 24)
            }
        ];
        
        this.lastUpdated = new Date();
        this.populateAccountExecutiveFilter();
    }

    initializeMonthSelector() {
        // Set default month to the current month
        const today = new Date();
        const currentMonth = today.getMonth() + 1; // JavaScript months are 0-based
        const currentYear = today.getFullYear();
        
        // Format as YYYY-MM for the option value
        const currentMonthValue = `${currentYear}-${currentMonth.toString().padStart(2, '0')}`;
        
        // Set the month selector to current month
        const monthSelector = document.getElementById('monthSelector');
        if (monthSelector) {
            monthSelector.value = currentMonthValue;
        }
        
        // Set the date range based on the selected month
        this.updateDateRangeFromMonth(currentMonthValue);
        
        console.log('Month selector initialized to:', currentMonthValue);
    }

    updateDateRangeFromMonth(monthValue) {
        // Parse the month value (YYYY-MM format)
        const [year, month] = monthValue.split('-');
        
        // Get the first day of the month
        const startDate = new Date(parseInt(year), parseInt(month) - 1, 1);
        
        // Get the last day of the month
        const endDate = new Date(parseInt(year), parseInt(month), 0);
        
        this.dateRange.startDate = startDate;
        this.dateRange.endDate = endDate;
        
        console.log('Date range updated from month selector:', startDate, 'to', endDate);
    }

    updateDateRange() {
        const monthSelector = document.getElementById('monthSelector');
        if (monthSelector) {
            const selectedMonth = monthSelector.value;
            this.updateDateRangeFromMonth(selectedMonth);
        } else {
            // Fallback to old date inputs if they exist
            const startDate = document.getElementById('startDate');
            const endDate = document.getElementById('endDate');
            if (startDate && endDate) {
                this.dateRange.startDate = new Date(startDate.value);
                this.dateRange.endDate = new Date(endDate.value);
            }
        }
        this.updateDashboard();
    }

    populateAccountExecutiveFilter() {
        const filterSelect = document.getElementById('aeFilter');
        filterSelect.innerHTML = '<option value="all">All Account Executives</option>';
        
        // Only show these 3 specific account executives in the dropdown
        const allowedAEs = ['Eddy', 'Tom', 'Galina'];
        
        allowedAEs.forEach(ae => {
            const option = document.createElement('option');
            option.value = ae;
            option.textContent = ae;
            filterSelect.appendChild(option);
        });
    }

    getFilteredData() {
        let filteredData = this.data;
        
        console.log('getFilteredData - Starting with', filteredData.length, 'deals');
        
        // Filter by account executive
        if (this.currentFilter !== 'all') {
            filteredData = filteredData.filter(item => item.teamMember === this.currentFilter);
            console.log('After AE filter:', filteredData.length, 'deals');
        }
        
        // Filter by date range using create date
        if (this.dateRange.startDate && this.dateRange.endDate) {
            console.log('Date range filter:', this.dateRange.startDate, 'to', this.dateRange.endDate);
            filteredData = filteredData.filter(item => {
                const createDate = new Date(item.createDate || item.stageDate); // Use createDate if available, fallback to stageDate
                const inRange = createDate >= this.dateRange.startDate && createDate <= this.dateRange.endDate;
                return inRange;
            });
            console.log('After date filter:', filteredData.length, 'deals');
        }
        
        return filteredData;
    }

    // Helper method to check if a date is within the selected date range
    isDateInRange(date) {
        if (!this.dateRange.startDate || !this.dateRange.endDate) return true;
        return date >= this.dateRange.startDate && date <= this.dateRange.endDate;
    }

    calculateRiskRating(deal) {
        const now = new Date();
        const stageDate = new Date(deal.stageDate);
        const lastContactDate = new Date(deal.lastContactDate);
        
        const daysInStage = Math.floor((now - stageDate) / (1000 * 60 * 60 * 24));
        const daysSinceContact = Math.floor((now - lastContactDate) / (1000 * 60 * 60 * 24));
        
        let riskScore = 0;
        
        // Stage-based risk
        if (deal.stage === 'Inbound Lead') {
            if (daysInStage > 12) riskScore += 35;
            else if (daysInStage > 9) riskScore += 25;
            else if (daysInStage > 6) riskScore += 10;
        } else if (deal.stage === 'Cold App Install') {
            if (daysInStage > 12) riskScore += 35;
            else if (daysInStage > 9) riskScore += 25;
            else if (daysInStage > 6) riskScore += 10;
        } else if (deal.stage === 'Qualified') {
            if (daysInStage > 14) riskScore += 35;
            else if (daysInStage > 10) riskScore += 25;
            else if (daysInStage > 7) riskScore += 10;
        } else if (deal.stage === 'Meeting Booked') {
            if (daysInStage > 14) riskScore += 35;
            else if (daysInStage > 10) riskScore += 25;
            else if (daysInStage > 7) riskScore += 10;
        } else if (deal.stage === 'Limbo') {
            if (daysInStage > 7) riskScore += 35;
            else if (daysInStage > 5) riskScore += 25;
            else if (daysInStage > 3) riskScore += 10;
        } else if (deal.stage === 'Missed Meeting') {
            if (daysInStage > 5) riskScore += 35;
            else if (daysInStage > 3) riskScore += 25;
            else if (daysInStage > 1) riskScore += 10;
        }
        
        // Contact-based risk
        if (daysSinceContact > 8) riskScore += 30;
        else if (daysSinceContact > 5) riskScore += 20;
        else if (daysSinceContact > 3) riskScore += 10;
        
        // Limbo-specific risk
        if (deal.stage === 'Limbo') {
            if (daysInStage > 10) riskScore += 25;
            else if (daysInStage > 6) riskScore += 20;
            else if (daysInStage > 3) riskScore += 10;
        } else if (deal.stage === 'Missed Meeting') {
            if (daysInStage > 5) riskScore += 25;
            else if (daysInStage > 3) riskScore += 20;
            else if (daysInStage > 1) riskScore += 10;
        }
        
        // Next activity risk
        if (!deal.nextActivityDate) riskScore += 10;
        
        return riskScore;
    }

    getRiskLevel(riskScore) {
        if (riskScore >= 80) return { level: 'Critical', class: 'risk-critical' };
        if (riskScore >= 60) return { level: 'High', class: 'risk-high' };
        if (riskScore >= 40) return { level: 'Medium', class: 'risk-medium' };
        return { level: 'Low', class: 'risk-low' };
    }

    // KPI Calculation Methods
    calculateClosedWonPerMonth() {
        // For closed deals, we need to check ALL deals (not just date-filtered ones)
        // because we want deals closed within the date range, not deals created within the date range
        let dealsToCheck = this.data;
        
        // Only apply account executive filter, not date filter
        if (this.currentFilter !== 'all') {
            dealsToCheck = dealsToCheck.filter(item => item.teamMember === this.currentFilter);
        }
        
        const currentDate = new Date();
        const currentMonth = currentDate.getMonth();
        const currentYear = currentDate.getFullYear();
        
        console.log('=== CALCULATING CLOSED WON PER MONTH ===');
        console.log('- Total deals to check:', dealsToCheck.length);
        console.log('- Current month:', currentMonth, 'Year:', currentYear);
        console.log('- Date range:', this.dateRange.startDate, 'to', this.dateRange.endDate);
        console.log('- Date range formatted:', this.dateRange.startDate?.toLocaleDateString(), 'to', this.dateRange.endDate?.toLocaleDateString());
        
        // Expected HubSpot IDs for September 2025
        const expectedIds = [
            '44828328860', '44491110031', '43725332915', '43639571356', '43597673539',
            '43588822035', '43385056594', '43347621682', '43248772182', '43209388451',
            '43214594907', '42005983339', '41645065852', '41204705923', '40076116236',
            '43948596325', '38868189586'
        ];
        
        console.log('- Expected IDs (17):', expectedIds);
        
        // Check which expected IDs are in our dataset
        const foundExpectedIds = dealsToCheck.filter(deal => expectedIds.includes(deal.id));
        console.log('- Found expected IDs in dataset:', foundExpectedIds.length);
        console.log('- Found IDs:', foundExpectedIds.map(deal => deal.id));
        
        // Check which expected IDs are missing
        const foundIds = foundExpectedIds.map(deal => deal.id);
        const missingIds = expectedIds.filter(id => !foundIds.includes(id));
        console.log('- Missing IDs from dataset:', missingIds);
        
        // Debug: Show all closed won deals
        const allClosedWonDeals = dealsToCheck.filter(deal => deal.stage === 'Closed Won' && deal.closeDate);
        console.log('- All closed won deals found:', allClosedWonDeals.length);
        
        // Show detailed info about each closed won deal
        allClosedWonDeals.forEach((deal, index) => {
            const closeDate = new Date(deal.closeDate);
            const isInRange = this.dateRange.startDate && this.dateRange.endDate && 
                             closeDate >= this.dateRange.startDate && closeDate <= this.dateRange.endDate;
            console.log(`  Deal ${index + 1}: ${deal.id} - Close Date: ${closeDate.toLocaleDateString()} - In Range: ${isInRange}`);
        });
        
        // Count deals closed within the selected date range only
        const closedWonDeals = dealsToCheck.filter(deal => {
            if (deal.stage !== 'Closed Won' || !deal.closeDate) return false;
            
            const closeDate = new Date(deal.closeDate);
            
            // Only count deals closed within the selected date range
            const isInDateRange = this.dateRange.startDate && this.dateRange.endDate && 
                                 closeDate >= this.dateRange.startDate && closeDate <= this.dateRange.endDate;
            
            return isInDateRange;
        });
        
        console.log('- Closed won deals in date range:', closedWonDeals.length);
        console.log('- Actual deal IDs found:', closedWonDeals.map(deal => deal.id));
        console.log('- Deal ID types:', closedWonDeals.map(deal => typeof deal.id));
        
        // Check which expected IDs are in the final result
        const foundInResult = closedWonDeals.filter(deal => expectedIds.includes(deal.id));
        console.log('- Expected IDs in final result:', foundInResult.length);
        console.log('- Expected IDs found:', foundInResult.map(deal => deal.id));
        
        // Check which expected IDs are missing from final result
        const foundInResultIds = foundInResult.map(deal => deal.id);
        const missingFromResult = expectedIds.filter(id => !foundInResultIds.includes(id));
        console.log('- Expected IDs missing from result:', missingFromResult);
        
        // Show details of missing deals
        if (missingFromResult.length > 0) {
            console.log('- Details of missing deals:');
            missingFromResult.forEach(id => {
                const deal = dealsToCheck.find(d => d.id === id);
                if (deal) {
                    console.log(`  ${id}: stage=${deal.stage}, closeDate=${deal.closeDate}, owner=${deal.teamMember}`);
                } else {
                    console.log(`  ${id}: NOT FOUND IN DATASET`);
                }
            });
        }
        
        // Show the actual deals that are being counted
        console.log('- Details of deals being counted:');
        closedWonDeals.forEach((deal, index) => {
            console.log(`  ${index + 1}. ID: ${deal.id} (${typeof deal.id}), Stage: ${deal.stage}, Close Date: ${deal.closeDate}, Owner: ${deal.teamMember}`);
        });
        
        console.log('- Final result:', closedWonDeals.length);
        console.log('=== END CLOSED WON CALCULATION ===');
        
        // Log the actual calculated value for debugging
        console.log(`🔢 CLOSED WON PER MONTH RESULT: ${closedWonDeals.length} deals`);
        
        return closedWonDeals.length;
    }

    calculateMeetingClosedRate() {
        // Calculate meeting closed rate: deals that exited "Meeting Booked" within date range AND closed won within 14 days
        let dealsToCheck = this.data;
        
        // Only apply account executive filter, not date filter
        if (this.currentFilter !== 'all') {
            dealsToCheck = dealsToCheck.filter(item => item.teamMember === this.currentFilter);
        }
        
        console.log('Calculating Meeting Closed Rate:');
        console.log('- Total deals to check:', dealsToCheck.length);
        console.log('- Date range:', this.dateRange.startDate, 'to', this.dateRange.endDate);
        
        // Deals that exited "Meeting Booked" within the selected date range
        const dealsExitedMeetingBooked = dealsToCheck.filter(deal => {
            if (!deal.meetingBookedExitDate) return false;
            const exitDate = new Date(deal.meetingBookedExitDate);
            return this.isDateInRange(exitDate);
        });
        
        console.log('- Deals that exited Meeting Booked in date range:', dealsExitedMeetingBooked.length);
        
        // Deals that closed won within 14 days of exiting Meeting Booked
        const closedDeals = dealsExitedMeetingBooked.filter(deal => {
            if (deal.stage !== 'Closed Won' || !deal.closeDate) return false;
            
            const exitDate = new Date(deal.meetingBookedExitDate);
            const closeDate = new Date(deal.closeDate);
            const daysDiff = Math.floor((closeDate - exitDate) / (1000 * 60 * 60 * 24));
            
            return daysDiff <= 14;
        });
        
        console.log('- Deals that closed won within 14 days of exiting Meeting Booked:', closedDeals.length);
        const rate = dealsExitedMeetingBooked.length > 0 ? Math.round((closedDeals.length / dealsExitedMeetingBooked.length) * 100) : 0;
        console.log('- Meeting closed rate:', rate, '%');
        console.log(`🔢 MEETING CLOSED RATE RESULT: ${rate}% (${closedDeals.length}/${dealsExitedMeetingBooked.length})`);
        
        if (dealsExitedMeetingBooked.length === 0) return 0;
        return rate;
    }

    calculateMissedMeetingRate() {
        // Calculate missed meeting rate: deals that entered "Meeting Booked" AND "Missed Meeting" within date range
        let dealsToCheck = this.data;
        
        // Only apply account executive filter, not date filter
        if (this.currentFilter !== 'all') {
            dealsToCheck = dealsToCheck.filter(item => item.teamMember === this.currentFilter);
        }
        
        console.log('=== CALCULATING MISSED MEETING RATE ===');
        console.log('- Total deals to check:', dealsToCheck.length);
        console.log('- Date range:', this.dateRange.startDate, 'to', this.dateRange.endDate);
        console.log('- Current filter:', this.currentFilter);
        
        // Debug: Show all deals with meetingBookedDate
        const allDealsWithMeetingBooked = dealsToCheck.filter(deal => deal.meetingBookedDate);
        console.log('- All deals with meetingBookedDate:', allDealsWithMeetingBooked.length);
        
        // Debug: Show sample of deals with meetingBookedDate
        console.log('- Sample deals with meetingBookedDate:', allDealsWithMeetingBooked.slice(0, 5).map(deal => ({
            id: deal.id,
            meetingBookedDate: deal.meetingBookedDate,
            missedMeetingDate: deal.missedMeetingDate,
            stage: deal.stage
        })));
        
        // Deals that entered "Meeting Booked" within the selected date range
        const dealsEnteredMeetingBooked = dealsToCheck.filter(deal => {
            if (!deal.meetingBookedDate) return false;
            const meetingBookedDate = new Date(deal.meetingBookedDate);
            const isInRange = this.isDateInRange(meetingBookedDate);
            console.log(`Deal ${deal.id}: meetingBookedDate=${meetingBookedDate}, isInRange=${isInRange}`);
            return isInRange;
        });
        
        console.log('- Deals that entered Meeting Booked in date range:', dealsEnteredMeetingBooked.length);
        
        // Debug: Show details of deals that entered Meeting Booked
        console.log('- Details of deals that entered Meeting Booked:', dealsEnteredMeetingBooked.map(deal => ({
            id: deal.id,
            meetingBookedDate: deal.meetingBookedDate,
            missedMeetingDate: deal.missedMeetingDate,
            stage: deal.stage
        })));
        
        // Of those deals, how many also entered "Missed Meeting" within the date range
        const dealsEnteredMissedMeeting = dealsEnteredMeetingBooked.filter(deal => {
            if (!deal.missedMeetingDate) return false;
            const missedMeetingDate = new Date(deal.missedMeetingDate);
            const isInRange = this.isDateInRange(missedMeetingDate);
            console.log(`Deal ${deal.id}: missedMeetingDate=${missedMeetingDate}, isInRange=${isInRange}`);
            return isInRange;
        });
        
        console.log('- Deals that entered both Meeting Booked and Missed Meeting in date range:', dealsEnteredMissedMeeting.length);
        
        // Debug: Show details of deals that entered Missed Meeting
        console.log('- Details of deals that entered Missed Meeting:', dealsEnteredMissedMeeting.map(deal => ({
            id: deal.id,
            meetingBookedDate: deal.meetingBookedDate,
            missedMeetingDate: deal.missedMeetingDate,
            stage: deal.stage
        })));
        
        const rate = dealsEnteredMeetingBooked.length > 0 ? Math.round((dealsEnteredMissedMeeting.length / dealsEnteredMeetingBooked.length) * 100) : 0;
        console.log('- Final calculation:', dealsEnteredMissedMeeting.length, '/', dealsEnteredMeetingBooked.length, '=', rate, '%');
        console.log(`🔢 MISSED MEETING RATE RESULT: ${rate}% (${dealsEnteredMissedMeeting.length}/${dealsEnteredMeetingBooked.length})`);
        console.log('=== END MISSED MEETING RATE CALCULATION ===');
        
        return rate;
    }

    calculateDealsWithTask() {
        const filteredData = this.getFilteredData();
        const activeDeals = filteredData.filter(deal => 
            deal.stage !== 'Closed Won' && 
            deal.stage !== 'Closed Lost' &&
            deal.stage !== 'SUBSCRIBED'
        );
        
        // For now, assume all active deals have tasks (this would need to be updated based on your actual task data)
        // This is a placeholder calculation - you'll need to add task data to your Google Sheet
        const dealsWithTasks = activeDeals.filter(deal => 
            deal.taskStatus === 'Completed' || 
            deal.taskStatus === 'Pending' ||
            deal.taskStatus === 'Open'
        );
        
        if (activeDeals.length === 0) return 0;
        return Math.round((dealsWithTasks.length / activeDeals.length) * 100);
    }

    calculateLimboMovedRate() {
        const filteredData = this.getFilteredData();
        
        // Deals that moved into limbo within the date range
        const limboDeals = filteredData.filter(deal => {
            if (!deal.limboDate) return false;
            const limboDate = new Date(deal.limboDate);
            return this.isDateInRange(limboDate);
        });
        
        // Deals that moved into limbo within date range AND moved out within 14 days
        const movedDeals = limboDeals.filter(deal => {
            if (deal.stage === 'Limbo') return false; // Still in limbo
            
            // Find when they moved out of limbo (next stage date after limbo)
            const limboDate = new Date(deal.limboDate);
            let movedOutDate = null;
            
            // Check if they have a stage date after limbo
            if (deal.stageDate && new Date(deal.stageDate) > limboDate) {
                movedOutDate = new Date(deal.stageDate);
            }
            
            if (!movedOutDate) return false;
            
            const daysDiff = Math.floor((movedOutDate - limboDate) / (1000 * 60 * 60 * 24));
            return daysDiff <= 14;
        });
        
        if (limboDeals.length === 0) return 0;
        return Math.round((movedDeals.length / limboDeals.length) * 100);
    }

    calculateOutboundMeetings() {
        // Placeholder for outbound meetings calculation
        // This will be implemented when the outbound meetings column is added to the Google Sheet
        console.log('Outbound Meetings: Placeholder - column to be added');
        console.log(`🔢 OUTBOUND MEETINGS RESULT: 0 (placeholder)`);
        return 0;
    }

    calculateOutboundDealsClosed() {
        // Placeholder for outbound deals closed calculation
        // This will be implemented when the outbound deals column is added to the Google Sheet
        console.log('Outbound Deals Closed: Placeholder - column to be added');
        console.log(`🔢 OUTBOUND DEALS CLOSED RESULT: 0 (placeholder)`);
        return 0;
    }

    calculateLimboOver14Days() {
        // Calculate percentage of open deals that entered limbo and either:
        // 1. Haven't exited limbo yet (no limboExitDate) AND have been in limbo >14 days, OR
        // 2. Exited limbo but took more than 14 days to exit
        const dealsToCheck = this.currentFilter === 'all' ? this.data : this.data.filter(deal => deal.teamMember === this.currentFilter);
        
        const openDeals = dealsToCheck.filter(deal => deal.stage !== 'Closed Won' && deal.stage !== 'Closed Lost');
        
        console.log('=== CALCULATING LIMBO OVER 14 DAYS ===');
        console.log('- Total open deals:', openDeals.length);
        
        // First, find all deals that have entered limbo
        const dealsThatEnteredLimbo = openDeals.filter(deal => deal.limboDate);
        console.log('- Deals that entered limbo:', dealsThatEnteredLimbo.length);
        
        const limboDeals = dealsThatEnteredLimbo.filter(deal => {
            const limboEnterDate = new Date(deal.limboDate);
            const limboExitDate = deal.limboExitDate ? new Date(deal.limboExitDate) : null;
            
            if (!limboExitDate) {
                // Still in limbo - check if it's been more than 14 days since entering
                const daysInLimbo = Math.floor((new Date() - limboEnterDate) / (1000 * 60 * 60 * 24));
                console.log(`Deal ${deal.id}: Still in limbo for ${daysInLimbo} days`);
                return daysInLimbo > 14;
            } else {
                // Exited limbo - check if it took more than 14 days to exit
                const daysInLimbo = Math.floor((limboExitDate - limboEnterDate) / (1000 * 60 * 60 * 24));
                console.log(`Deal ${deal.id}: Exited limbo after ${daysInLimbo} days`);
                return daysInLimbo > 14;
            }
        });
        
        // Calculate percentage of deals that entered limbo and took >14 days to exit
        const percentage = dealsThatEnteredLimbo.length > 0 ? Math.round((limboDeals.length / dealsThatEnteredLimbo.length) * 100) : 0;
        console.log('- Deals that entered limbo and took >14 days to exit (or still in limbo >14 days):', limboDeals.length);
        console.log('- Percentage of deals that entered limbo and got stuck >14 days:', percentage + '%');
        console.log(`🔢 LIMBO OVER 14 DAYS RESULT: ${percentage}% (${limboDeals.length}/${dealsThatEnteredLimbo.length})`);
        console.log('=== END LIMBO CALCULATION ===');
        
        return percentage;
    }

    calculateDealsWithoutNextStep() {
        // Calculate percentage of open deals that do not have a date for next step (Column Z)
        const dealsToCheck = this.currentFilter === 'all' ? this.data : this.data.filter(deal => deal.teamMember === this.currentFilter);
        
        const openDeals = dealsToCheck.filter(deal => deal.stage !== 'Closed Won' && deal.stage !== 'Closed Lost');
        const dealsWithoutNextStep = openDeals.filter(deal => {
            // Check if deal has no next step date (Column Z is empty/unknown)
            return !deal.nextStepDate || deal.nextStepDate === '';
        });
        
        const percentage = openDeals.length > 0 ? Math.round((dealsWithoutNextStep.length / openDeals.length) * 100) : 0;
        console.log('Deals without next step (Column Z empty):', dealsWithoutNextStep.length, 'out of', openDeals.length, 'open deals =', percentage + '%');
        console.log(`🔢 DEALS WITHOUT NEXT STEP RESULT: ${percentage}% (${dealsWithoutNextStep.length}/${openDeals.length})`);
        return percentage;
    }

    calculateAverageDealAge() {
        // Calculate average age of open deals in days (all open deals, no date filter)
        const dealsToCheck = this.currentFilter === 'all' ? this.data : this.data.filter(deal => deal.teamMember === this.currentFilter);
        
        const openDeals = dealsToCheck.filter(deal => deal.stage !== 'Closed Won' && deal.stage !== 'Closed Lost');
        
        if (openDeals.length === 0) return 0;
        
        const totalAge = openDeals.reduce((sum, deal) => {
            const createDate = new Date(deal.createDate || deal.stageDate);
            const daysOld = Math.floor((new Date() - createDate) / (1000 * 60 * 60 * 24));
            return sum + daysOld;
        }, 0);
        
        const averageAge = Math.round(totalAge / openDeals.length);
        console.log('Average deal age:', averageAge, 'days for', openDeals.length, 'open deals');
        console.log(`🔢 AVERAGE DEAL AGE RESULT: ${averageAge} days (${openDeals.length} open deals)`);
        return averageAge;
    }

    updateKPICard(kpiId, value, target, progressId, textId, statusId, isPercentage = false, isLowerBetter = false, comparisonId = null) {
        // Debug: Log the value being passed to updateKPICard
        if (kpiId === 'closedWonValue') {
            console.log(`=== UPDATEKPICARD DEBUG ===`);
            console.log(`- KPI ID: ${kpiId}`);
            console.log(`- Value passed: ${value} (type: ${typeof value})`);
            console.log(`- Target: ${target}`);
            console.log(`- Progress ID: ${progressId}`);
            console.log(`- Text ID: ${textId}`);
            console.log(`- Status ID: ${statusId}`);
            console.log(`- Call stack:`, new Error().stack);
            
            // If this is the call setting it to 16, let's get more details
            if (value === 16) {
                console.log(`🚨 FOUND THE CULPRIT! Setting closedWonValue to 16`);
                console.log(`- Full call stack:`, new Error().stack);
            }
            
            // Also check if this is being called multiple times
            if (!this.updateKPICardCount) {
                this.updateKPICardCount = 0;
            }
            this.updateKPICardCount++;
            console.log(`- updateKPICard call count: ${this.updateKPICardCount}`);
        }
        
        // Display value with % if it's a percentage KPI
        const displayValue = isPercentage ? `${value}%` : value;
        document.getElementById(kpiId).textContent = displayValue;
        
        // Debug: Log what's actually being displayed
        if (kpiId === 'closedWonValue') {
            console.log(`- Display value set to: ${displayValue}`);
            console.log(`- Element text content after update: ${document.getElementById(kpiId).textContent}`);
            
            // Add a mutation observer to catch any changes to the element
            const element = document.getElementById(kpiId);
            if (element) {
                const observer = new MutationObserver((mutations) => {
                    mutations.forEach((mutation) => {
                        if (mutation.type === 'childList' || mutation.type === 'characterData') {
                            console.log(`- MUTATION DETECTED: Value changed from ${mutation.oldValue} to ${element.textContent}`);
                            console.log(`- Mutation type: ${mutation.type}`);
                            console.log(`- Stack trace:`, new Error().stack);
                            
                            // If the value is being set to 16, let's see what's in the DOM
                            if (element.textContent === '16') {
                                console.log(`🚨 DOM MANIPULATION DETECTED! Element set to 16`);
                                console.log(`- Element innerHTML: ${element.innerHTML}`);
                                console.log(`- Element textContent: ${element.textContent}`);
                                console.log(`- Element parentNode: ${element.parentNode}`);
                                console.log(`- Full call stack:`, new Error().stack);
                            }
                        }
                    });
                });
                observer.observe(element, {
                    childList: true,
                    characterData: true,
                    subtree: true,
                    characterDataOldValue: true
                });
            }
        }
        
        const percentage = Math.min((value / target) * 100, 100);
        const progressElement = document.getElementById(progressId);
        progressElement.style.width = `${percentage}%`;
        
        // For "lower is better" metrics, show percentage over target instead of percentage of target
        if (isLowerBetter) {
            if (value <= target) {
                document.getElementById(textId).textContent = "Target Met";
            } else {
                const percentOverTarget = Math.round(((value - target) / target) * 100);
                document.getElementById(textId).textContent = `${percentOverTarget}% over target`;
            }
        } else {
            document.getElementById(textId).textContent = `${Math.round(percentage)}%`;
        }
        
        let status = 'Excellent';
        let statusClass = 'excellent';
        
        // For KPIs where lower is better (like Missed Meeting Rate, % Deals in Limbo, etc.)
        if (isLowerBetter) {
            // For "<%" targets: hitting the low target is Excellent, up to 25% over is Okay, more is Poor
            if (value <= target) {
                status = 'Excellent';
                statusClass = 'excellent';
            } else if (value <= target * 1.25) { // Up to 25% over target
                status = 'Good';
                statusClass = 'okay';
            } else {
                status = 'Poor';
                statusClass = 'poor';
            }
        } else {
            // For KPIs where higher is better
            if (percentage >= 100) {
                status = 'Excellent';
                statusClass = 'excellent';
            } else if (percentage >= 75) { // 75-100% is Good
                status = 'Good';
                statusClass = 'okay';
            } else { // <75% is Poor
                status = 'Poor';
                statusClass = 'poor';
            }
        }
        
        const statusElement = document.getElementById(statusId);
        statusElement.textContent = status;
        statusElement.className = `kpi-status ${statusClass}`;
        
        // Update progress bar color based on status
        progressElement.className = `progress-fill ${statusClass}`;
        
        // Update comparison if provided
        if (comparisonId) {
            this.updateComparison(comparisonId, value, isLowerBetter);
        }
    }

    // Calculate previous year's date range
    getPreviousYearDateRange() {
        const currentStart = this.dateRange.startDate;
        const currentEnd = this.dateRange.endDate;
        
        if (!currentStart || !currentEnd) return null;
        
        const prevStart = new Date(currentStart);
        prevStart.setFullYear(prevStart.getFullYear() - 1);
        
        const prevEnd = new Date(currentEnd);
        prevEnd.setFullYear(prevEnd.getFullYear() - 1);
        
        return { startDate: prevStart, endDate: prevEnd };
    }
    
    // Calculate previous year KPI values
    calculatePreviousYearKPIs() {
        const prevYearRange = this.getPreviousYearDateRange();
        if (!prevYearRange) {
            console.log('No previous year range available');
            return {};
        }
        
        console.log('Calculating previous year KPIs for range:', prevYearRange.startDate, 'to', prevYearRange.endDate);
        
        // Store current date range
        const currentRange = { ...this.dateRange };
        
        // Temporarily set to previous year range
        this.dateRange = prevYearRange;
        
        const prevYearKPIs = {
            closedWon: this.calculateClosedWonPerMonth(),
            meetingClosedRate: this.calculateMeetingClosedRate(),
            missedMeetingRate: this.calculateMissedMeetingRate(),
            outboundMeetings: this.calculateOutboundMeetings(),
            outboundDealsClosed: this.calculateOutboundDealsClosed(),
            limboOver14Days: this.calculateLimboOver14Days(),
            dealsWithoutNextStep: this.calculateDealsWithoutNextStep(),
            averageDealAge: this.calculateAverageDealAge()
        };
        
        console.log('Previous year KPIs calculated:', prevYearKPIs);
        
        // Restore current date range
        this.dateRange = currentRange;
        
        return prevYearKPIs;
    }
    
    // Update comparison display
    updateComparison(comparisonId, currentValue, isLowerBetter = false) {
        const prevYearKPIs = this.calculatePreviousYearKPIs();
        const kpiName = comparisonId.replace('Comparison', '');
        const prevValue = prevYearKPIs[kpiName] || 0;
        
        console.log(`Comparison for ${kpiName}: Current=${currentValue}, Previous=${prevValue}`);
        
        if (prevValue === 0) {
            console.log(`No previous year data for ${kpiName}`);
            document.getElementById(comparisonId).innerHTML = '<span class="comparison-arrow">-</span><span class="comparison-text">No data</span>';
            return;
        }
        
        const change = currentValue - prevValue;
        const changePercent = Math.round((change / prevValue) * 100);
        
        let arrowClass = 'up';
        let arrowSymbol = '↑';
        let isBetter = change > 0;
        
        // For KPIs where lower is better (like Missed Meeting Rate)
        if (isLowerBetter) {
            isBetter = change < 0;
            arrowClass = isBetter ? 'up' : 'down';
            arrowSymbol = isBetter ? '↓' : '↑';
        } else {
            arrowClass = isBetter ? 'up' : 'down';
            arrowSymbol = isBetter ? '↑' : '↓';
        }
        
        const comparisonElement = document.getElementById(comparisonId);
        comparisonElement.innerHTML = `
            <span class="comparison-arrow ${arrowClass}">${arrowSymbol}</span>
            <span class="comparison-text">${Math.abs(changePercent)}%</span>
        `;
    }

    updateActiveDealsTable() {
        // Get all deals and apply both date and AE filters
        let filteredDeals = this.data;
        
        // Filter by account executive
        if (this.currentFilter !== 'all') {
            filteredDeals = filteredDeals.filter(deal => deal.teamMember === this.currentFilter);
            console.log('After AE filter:', filteredDeals.length, 'deals');
        }
        
        // Filter by create date within the selected month
        if (this.dateRange.startDate && this.dateRange.endDate) {
            filteredDeals = filteredDeals.filter(deal => {
                const createDate = new Date(deal.createDate);
                return createDate >= this.dateRange.startDate && createDate <= this.dateRange.endDate;
            });
            console.log('After date filter:', filteredDeals.length, 'deals');
        }
        
        console.log('Final filtered deals for table:', filteredDeals.length);
        
        // Store the data for sorting
        this.tableData = filteredDeals;
        
        this.renderTable();
        this.setupTableSorting();
    }

    renderTable() {
        const tbody = document.querySelector('#activeDealsTable tbody');
        tbody.innerHTML = '';
        
        this.tableData.forEach(deal => {
            const row = document.createElement('tr');
            
            const now = new Date();
            const stageDate = new Date(deal.stageDate);
            
            const daysInStage = Math.floor((now - stageDate) / (1000 * 60 * 60 * 24));
            
            // Calculate risk rating - 0 for closed deals
            let riskDisplay;
            if (deal.stage === 'Closed Won' || deal.stage === 'Closed Lost') {
                riskDisplay = '<span class="risk-closed">Low (0)</span>';
            } else {
                const riskScore = this.calculateRiskRating(deal);
                const riskInfo = this.getRiskLevel(riskScore);
                riskDisplay = `<span class="${riskInfo.class}">${riskInfo.level} (${riskScore})</span>`;
            }
            
            row.innerHTML = `
                <td>${deal.id}</td>
                <td>${deal.name}</td>
                <td><span class="stage-${deal.stage.toLowerCase().replace(/\s+/g, '-')}">${deal.stage}</span></td>
                <td>${deal.teamMember}</td>
                <td>${new Date(deal.createDate).toLocaleDateString()}</td>
                <td>${daysInStage}</td>
                <td>${riskDisplay}</td>
            `;
            
            tbody.appendChild(row);
        });
    }

    setupTableSorting() {
        const sortableHeaders = document.querySelectorAll('.sortable');
        
        sortableHeaders.forEach(header => {
            header.addEventListener('click', () => {
                const sortType = header.dataset.sort;
                const currentSort = header.dataset.currentSort || 'none';
                
                // Determine new sort direction
                let newSort;
                if (currentSort === 'none' || currentSort === 'desc') {
                    newSort = 'asc';
                } else {
                    newSort = 'desc';
                }
                
                // Clear other sort indicators
                sortableHeaders.forEach(h => {
                    h.dataset.currentSort = 'none';
                    h.querySelector('.sort-arrow').className = 'sort-arrow';
                });
                
                // Set current sort
                header.dataset.currentSort = newSort;
                const arrow = header.querySelector('.sort-arrow');
                arrow.className = `sort-arrow active ${newSort}`;
                
                // Sort the data
                this.sortTableData(sortType, newSort);
                this.renderTable();
            });
        });
    }

    sortTableData(sortType, direction) {
        this.tableData.sort((a, b) => {
            let aValue, bValue;
            
            if (sortType === 'stage') {
                aValue = a.stage;
                bValue = b.stage;
            } else if (sortType === 'risk') {
                // Handle 0 values for closed deals
                if (a.stage === 'Closed Won' || a.stage === 'Closed Lost') {
                    aValue = 0;
                } else {
                    aValue = this.calculateRiskRating(a);
                }
                
                if (b.stage === 'Closed Won' || b.stage === 'Closed Lost') {
                    bValue = 0;
                } else {
                    bValue = this.calculateRiskRating(b);
                }
            }
            
            if (direction === 'asc') {
                return aValue > bValue ? 1 : aValue < bValue ? -1 : 0;
            } else {
                return aValue < bValue ? 1 : aValue > bValue ? -1 : 0;
            }
        });
    }

    updateDashboard() {
        try {
            console.log('=== UPDATE DASHBOARD CALLED ===');
            console.log('Updating dashboard with', this.data.length, 'deals');
            console.log('Current filter:', this.currentFilter);
            console.log('Date range:', this.dateRange);
            console.log('Stack trace:', new Error().stack);
            
            // Track if this is a second call
            if (!this.dashboardUpdateCount) {
                this.dashboardUpdateCount = 0;
            }
            this.dashboardUpdateCount++;
            console.log(`- Dashboard update count: ${this.dashboardUpdateCount}`);
            
            // If this is a second call, let's see what's happening
            if (this.dashboardUpdateCount > 1) {
                console.log(`🚨 MULTIPLE DASHBOARD UPDATES DETECTED!`);
                console.log(`- This is update #${this.dashboardUpdateCount}`);
                console.log(`- Call stack:`, new Error().stack);
            }
            
            // Check if dashboard container is visible
            const container = document.querySelector('.dashboard-container');
            console.log('Dashboard container found:', container);
            if (container) {
                console.log('Container display:', window.getComputedStyle(container).display);
                console.log('Container visibility:', window.getComputedStyle(container).visibility);
            }
            
            // Update KPI cards
            console.log('Updating KPI cards...');
            const closedWonElement = document.getElementById('closedWonValue');
            console.log('Closed Won element found:', closedWonElement);
            
            // Set target based on whether showing individual or group data
            const closedWonTarget = this.currentFilter === 'all' ? 30 : 10;
            
            // Update the target label
            document.getElementById('closedWonTarget').textContent = `Target: ${closedWonTarget}`;
            
            // Date-Filtered KPIs (these use date range)
            const closedWonValue = this.calculateClosedWonPerMonth();
            console.log(`=== DASHBOARD UPDATE DEBUG ===`);
            console.log(`- Closed Won Value calculated: ${closedWonValue}`);
            console.log(`- Dashboard update count: ${this.dashboardUpdateCount}`);
            console.log(`- About to call updateKPICard with value: ${closedWonValue}`);
            this.updateKPICard('closedWonValue', closedWonValue, closedWonTarget, 'closedWonProgress', 'closedWonText', 'closedWonStatus', false, false, 'closedWonComparison');
            
            // Debug: Check value immediately after setting
            setTimeout(() => {
                const element = document.getElementById('closedWonValue');
                console.log(`- Value on page after 100ms: ${element ? element.textContent : 'ELEMENT NOT FOUND'}`);
            }, 100);
            this.updateKPICard('meetingClosedRate', this.calculateMeetingClosedRate(), 35, 'meetingClosedProgress', 'meetingClosedText', 'meetingClosedStatus', true, false, 'meetingClosedComparison');
            this.updateKPICard('missedMeetingRate', this.calculateMissedMeetingRate(), 10, 'missedMeetingProgress', 'missedMeetingText', 'missedMeetingStatus', true, true, 'missedMeetingComparison'); // Lower is better
            this.updateKPICard('outboundMeetings', this.calculateOutboundMeetings(), 5, 'outboundMeetingsProgress', 'outboundMeetingsText', 'outboundMeetingsStatus', false, false, 'outboundMeetingsComparison');
            this.updateKPICard('outboundDealsClosed', this.calculateOutboundDealsClosed(), 3, 'outboundDealsClosedProgress', 'outboundDealsClosedText', 'outboundDealsClosedStatus', false, false, 'outboundDealsClosedComparison');
            
            // Non-Date-Filtered KPIs (these use all open deals, no date filter)
            this.updateKPICard('limboOver14Days', this.calculateLimboOver14Days(), 20, 'limboOver14DaysProgress', 'limboOver14DaysText', 'limboOver14DaysStatus', true, true, 'limboOver14DaysComparison'); // Lower is better
            this.updateKPICard('dealsWithoutNextStep', this.calculateDealsWithoutNextStep(), 15, 'dealsWithoutNextStepProgress', 'dealsWithoutNextStepText', 'dealsWithoutNextStepStatus', true, true, 'dealsWithoutNextStepComparison'); // Lower is better
            this.updateKPICard('averageDealAge', this.calculateAverageDealAge(), 30, 'averageDealAgeProgress', 'averageDealAgeText', 'averageDealAgeStatus', false, true, 'averageDealAgeComparison'); // Lower is better
            
            // Update active deals table
            this.updateActiveDealsTable();
            
            // Update last updated time
            if (this.lastUpdated) {
                document.getElementById('lastUpdated').textContent = this.lastUpdated.toLocaleTimeString();
            }
            
            console.log('Dashboard update complete');
        } catch (error) {
            console.error('Error updating dashboard:', error);
            // Show basic error message
            document.body.innerHTML += `
                <div style="position: fixed; top: 10px; right: 10px; background: red; color: white; padding: 10px; border-radius: 5px;">
                    Dashboard Error: ${error.message}
                </div>
            `;
        }
    }

    async refreshData() {
        console.log('Refreshing data...');
        await this.loadRealData();
        this.updateDashboard();
        console.log('Data refresh complete');
    }

    startAutoRefresh() {
        // Auto-refresh every 5 minutes
        setInterval(() => {
            this.refreshData();
        }, 300000);
    }

    async exportReport() {
        try {
            console.log('Starting PDF export...');
            
            // Get the KPI container element
            const kpiContainer = document.querySelector('.kpi-container');
            if (!kpiContainer) {
                console.error('KPI container not found');
                return;
            }

            // Create a temporary container for PDF export
            const tempContainer = document.createElement('div');
            tempContainer.style.position = 'absolute';
            tempContainer.style.left = '-9999px';
            tempContainer.style.top = '0';
            tempContainer.style.width = '1200px';
            tempContainer.style.backgroundColor = '#f5f5f5';
            tempContainer.style.padding = '20px';
            tempContainer.style.fontFamily = 'Arial, sans-serif';
            
            // Clone the KPI container
            const clonedContainer = kpiContainer.cloneNode(true);
            
            // Add title and date range
            const title = document.createElement('h1');
            title.textContent = 'Sales KPI Dashboard Report';
            title.style.textAlign = 'center';
            title.style.marginBottom = '20px';
            title.style.color = '#333';
            title.style.fontSize = '24px';
            
            const dateRange = document.createElement('div');
            const startDate = this.dateRange.startDate ? this.dateRange.startDate.toLocaleDateString() : 'N/A';
            const endDate = this.dateRange.endDate ? this.dateRange.endDate.toLocaleDateString() : 'N/A';
            dateRange.textContent = `Date Range: ${startDate} - ${endDate}`;
            dateRange.style.textAlign = 'center';
            dateRange.style.marginBottom = '20px';
            dateRange.style.color = '#666';
            dateRange.style.fontSize = '16px';
            
            const accountExecutive = document.createElement('div');
            const aeText = this.currentFilter === 'all' ? 'All Account Executives' : this.currentFilter;
            accountExecutive.textContent = `Account Executive: ${aeText}`;
            accountExecutive.style.textAlign = 'center';
            accountExecutive.style.marginBottom = '30px';
            accountExecutive.style.color = '#666';
            accountExecutive.style.fontSize = '16px';
            
            tempContainer.appendChild(title);
            tempContainer.appendChild(dateRange);
            tempContainer.appendChild(accountExecutive);
            tempContainer.appendChild(clonedContainer);
            
            // Add to document temporarily
            document.body.appendChild(tempContainer);
            
            // Wait a moment for styles to apply
            await new Promise(resolve => setTimeout(resolve, 100));
            
            // Convert to canvas
            const canvas = await html2canvas(tempContainer, {
                scale: 2,
                useCORS: true,
                backgroundColor: '#f5f5f5',
                width: 1200,
                height: tempContainer.scrollHeight
            });
            
            // Remove temporary container
            document.body.removeChild(tempContainer);
            
            // Create PDF
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF('p', 'mm', 'a4');
            const imgWidth = 210; // A4 width in mm
            const pageHeight = 295; // A4 height in mm
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            let heightLeft = imgHeight;
            
            let position = 0;
            
            // Add image to PDF
            pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, position, imgWidth, imgHeight);
            heightLeft -= pageHeight;
            
            // Add new page if content is longer than one page
            while (heightLeft >= 0) {
                position = heightLeft - imgHeight;
                pdf.addPage();
                pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, position, imgWidth, imgHeight);
                heightLeft -= pageHeight;
            }
            
            // Generate filename with current date
            const now = new Date();
            const dateStr = now.toISOString().split('T')[0];
            const filename = `sales-kpi-report-${dateStr}.pdf`;
            
            // Download PDF
            pdf.save(filename);
            
            console.log('PDF export completed successfully');
            
        } catch (error) {
            console.error('Error exporting PDF:', error);
            alert('Error generating PDF report. Please try again.');
        }
    }

    // Configuration methods removed
}

// Initialize dashboard when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM Content Loaded event fired');
    console.log('Document ready state:', document.readyState);
    
    try {
        console.log('Initializing Sales KPI Dashboard...');
        console.log('Available elements:');
        console.log('- refreshBtn:', document.getElementById('refreshBtn'));
        console.log('- startDate:', document.getElementById('startDate'));
        console.log('- endDate:', document.getElementById('endDate'));
        
        new SalesKPIDashboard();
        console.log('Dashboard initialized successfully');
    } catch (error) {
        console.error('Error initializing dashboard:', error);
        console.error('Error stack:', error.stack);
        // Show error message to user
        document.body.innerHTML = `
            <div style="padding: 20px; text-align: center; color: red;">
                <h2>Error Loading Dashboard</h2>
                <p>There was an error loading the dashboard. Please check the console for details.</p>
                <p>Error: ${error.message}</p>
                <p>Stack: ${error.stack}</p>
            </div>
        `;
    }
});

// Also try immediate execution as fallback
if (document.readyState === 'loading') {
    console.log('Document still loading, waiting for DOMContentLoaded');
} else {
    console.log('Document already loaded, initializing immediately');
    try {
        new SalesKPIDashboard();
    } catch (error) {
        console.error('Immediate initialization failed:', error);
    }
}