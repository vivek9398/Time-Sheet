# Employee Attendance System - Setup Instructions

## Google Apps Script Setup

1. **Create a new Google Sheet**
   - Go to [Google Sheets](https://sheets.google.com)
   - Create a new blank spreadsheet
   - Name it "Employee Attendance"

2. **Add the Apps Script**
   - In your Google Sheet, go to `Extensions` → `Apps Script`
   - Delete the default code and paste the code from `Code.gs`
   - Save the project (Ctrl+S)

3. **Deploy as Web App**
   - Click `Deploy` → `New deployment`
   - Choose type: `Web app`
   - Execute as: `Me`
   - Who has access: `Anyone` (important for CORS)
   - Click `Deploy`
   - Copy the Web App URL

4. **Update the HTML file**
   - Replace the `SCRIPT_URL` in `index.html` with your Web App URL

## Troubleshooting "Failed to fetch" Error

The error occurs due to CORS restrictions. The fixes implemented:

1. **Added CORS headers** in Google Apps Script
2. **Added `doOptions` function** to handle preflight requests
3. **Improved error handling** in the HTML file
4. **Added `mode: 'cors'`** to the fetch request

## Testing

1. Deploy your HTML file to GitHub Pages
2. Test with a valid employee ID
3. Check the Google Sheet for recorded attendance

## Common Issues

- **Wrong deployment settings**: Ensure "Who has access" is set to "Anyone"
- **Old deployment**: After code changes, create a new deployment
- **Network issues**: Check browser console for detailed error messages