# Sales KPI Dashboard

A comprehensive, real-time dashboard for tracking sales team Key Performance Indicators (KPIs) with Google Sheets integration.

## Features

### 📊 KPI Tracking
- **10 Closed-Won Per Month** - Track monthly closed deals
- **35% Meeting Closed Rate** - Monitor meeting completion rates
- **< 10% Missed Meeting Rate** - Track meeting attendance
- **100% of Deals with Open Task** - Ensure task coverage
- **75% Deals Moved Out of Limbo by 14 Days** - Monitor deal progression
- **5 Outbound Meetings Booked** - Track outbound activity
- **100% Meetings Booked & New Subs Called within 24 Hours** - Follow-up compliance

### 🎨 Dashboard Features
- **Real-time Updates** - Automatic data refresh every 5 minutes
- **Interactive Charts** - Trend analysis and performance visualization
- **Responsive Design** - Works on desktop, tablet, and mobile
- **Export Functionality** - Generate CSV reports
- **Google Sheets Integration** - Connect to your live data source
- **Activity Tracking** - Recent activity table with team member details

### 📈 Visualizations
- Progress bars with color-coded performance indicators
- Line charts for trend analysis
- Doughnut charts for monthly overview
- Status indicators (Excellent/Good/Needs Improvement)

## Quick Start

1. **Open the Dashboard**
   ```bash
   # Simply open index.html in your web browser
   open index.html
   ```

2. **View Sample Data**
   - The dashboard loads with sample data for demonstration
   - All KPIs are calculated and displayed with progress indicators

3. **Connect Google Sheets** (Optional)
   - Click "Connect Sheet" to integrate with your Google Sheets data
   - Follow the setup instructions in the modal

## Google Sheets Integration

### Setup Instructions

1. **Create Google Sheet Structure**
   Create a Google Sheet with these tabs:
   
   **Deals Tab:**
   | Deal_ID | Close_Date | Deal_Value | Team_Member | Stage |
   |---------|------------|------------|-------------|-------|
   | D001    | 2024-01-15 | 5000       | John Doe    | Closed Won |
   
   **Meetings Tab:**
   | Meeting_ID | Meeting_Date | Status    | Team_Member |
   |------------|--------------|-----------|-------------|
   | M001       | 2024-01-10   | Completed | John Doe    |
   
   **Tasks Tab:**
   | Task_ID | Deal_ID | Has_Task | Task_Date | Team_Member |
   |---------|---------|----------|-----------|-------------|
   | T001    | D001    | TRUE     | 2024-01-16| John Doe    |
   
   **Limbo Tab:**
   | Limbo_ID | Deal_ID | Limbo_Date | Moved_Date | Team_Member |
   |----------|---------|------------|------------|-------------|
   | L001     | D001    | 2024-01-01 | 2024-01-08 | John Doe    |
   
   **Outbound Tab:**
   | Outbound_ID | Outbound_Date | Team_Member |
   |-------------|---------------|-------------|
   | O001        | 2024-01-07    | John Doe    |
   
   **Subscriptions Tab:**
   | Sub_ID | First_Meeting_Date | Called_Within_24h | Team_Member |
   |--------|-------------------|-------------------|-------------|
   | S001   | 2024-01-02        | TRUE              | John Doe    |

2. **Make Sheet Public**
   - Go to File > Share > Change to "Anyone with the link can view"
   - Copy the Sheet ID from the URL (the long string between `/d/` and `/edit`)

3. **Get Google Sheets API Key**
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing
   - Enable Google Sheets API
   - Create credentials (API Key)
   - Restrict the key to Google Sheets API

4. **Connect in Dashboard**
   - Click "Connect Sheet" button
   - Enter your API key and Sheet ID
   - The dashboard will automatically fetch and process your data

## File Structure

```
Sales Dashboard/
├── index.html                    # Main dashboard page
├── styles.css                    # Dashboard styling
├── script.js                     # Main dashboard logic
├── google-sheets-integration.js  # Google Sheets API integration
├── README.md                     # This file
└── Sales KPI Coefficient Report.xlsx  # Sample data file
```

## KPI Calculations

### 1. Closed-Won Per Month
- **Formula**: Count of deals with `Close_Date` in current month
- **Target**: 10 deals
- **Status**: Excellent (100%+), Good (70-99%), Needs Improvement (<70%)

### 2. Meeting Closed Rate
- **Formula**: (Completed Meetings / Total Meetings) × 100
- **Target**: 35%
- **Status**: Based on percentage of target achieved

### 3. Missed Meeting Rate
- **Formula**: (Missed Meetings / Total Meetings) × 100
- **Target**: < 10%
- **Status**: Excellent if ≤10%, decreases as rate increases

### 4. Deals with Open Task
- **Formula**: (Deals with Tasks / Total Deals) × 100
- **Target**: 100%
- **Status**: Based on percentage of target achieved

### 5. Limbo Moved Rate
- **Formula**: (Deals moved within 14 days / Total Limbo Deals) × 100
- **Target**: 75%
- **Status**: Based on percentage of target achieved

### 6. Outbound Meetings Booked
- **Formula**: Count of outbound meetings in current month
- **Target**: 5 meetings
- **Status**: Based on percentage of target achieved

### 7. Meetings Booked & New Subs Called
- **Formula**: (Called within 24h / Total New Subs) × 100
- **Target**: 100%
- **Status**: Based on percentage of target achieved

## Customization

### Adding New KPIs
1. Add new KPI card to `index.html`
2. Create calculation method in `script.js`
3. Update dashboard update logic
4. Add to export functionality

### Styling Changes
- Modify `styles.css` for visual customization
- Color schemes, fonts, and layouts can be easily adjusted
- Responsive breakpoints can be modified for different screen sizes

### Data Sources
- Currently supports Google Sheets integration
- Can be extended to support other data sources (APIs, databases, etc.)
- Sample data structure provided for testing

## Troubleshooting

### Common Issues

1. **Google Sheets Not Loading**
   - Verify API key is correct and has Sheets API access
   - Check that sheet is publicly accessible
   - Ensure sheet ID is correct

2. **KPIs Showing 0 or Incorrect Values**
   - Verify data format in Google Sheets matches expected structure
   - Check date formats (should be YYYY-MM-DD)
   - Ensure column headers match exactly

3. **Charts Not Displaying**
   - Check browser console for JavaScript errors
   - Verify Chart.js library is loading
   - Ensure data is being passed to chart functions

### Browser Compatibility
- Modern browsers (Chrome, Firefox, Safari, Edge)
- JavaScript ES6+ features used
- Responsive design for mobile devices

## Support

For issues or questions:
1. Check the browser console for error messages
2. Verify your Google Sheets data format
3. Ensure all required files are present
4. Test with sample data first

## Future Enhancements

- [ ] Real-time notifications for missed targets
- [ ] Team member performance comparisons
- [ ] Historical data analysis
- [ ] Automated report scheduling
- [ ] Mobile app version
- [ ] Advanced filtering and search
- [ ] Integration with CRM systems
