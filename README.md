# Enabling Google Sheets API and Google Drive API

Here's a step-by-step guide to enable the Google Sheets API and Google Drive API in the Google Cloud Console:

## Step-by-Step Instructions

1. **Go to Google Cloud Console**
   - Open your web browser and navigate to: [https://console.cloud.google.com/](https://console.cloud.google.com/)
   - Sign in with your Google account if you haven’t already.

2. **Create or Select a Project**
   - At the top of the page, you’ll see a project dropdown (it might say "Select a project" or show an existing project name).
   - Click the dropdown and then click **"New Project"** (or select an existing project if you want to use one).
   - If creating a new project:
     - Enter a **Project Name** (e.g., "PropertyAnalysis").
     - Leave the **Location** as "No organization" unless you’re part of a Google Cloud organization.
     - Click **Create**.
   - Wait a few seconds for the project to be created, then ensure it’s selected in the dropdown at the top.

3. **Navigate to the API Library**
   - In the left sidebar, hover over **"APIs & Services"** (it might be collapsed—click the hamburger menu ☰ if you don’t see it).
   - Click **"Library"** from the submenu that appears.

4. **Enable Google Sheets API**
   - In the API Library search bar, type **"Google Sheets API"**.
   - Click on the **Google Sheets API** result from the list.
   - On the API page, click the blue **Enable** button.
   - Wait for it to enable (this might take a few seconds). Once enabled, you’ll see a "Manage" button instead.

5. **Enable Google Drive API**
   - Go back to the API Library (you can click "API Library" in the left sidebar again or use the back button).
   - In the search bar, type **"Google Drive API"**.
   - Click on the **Google Drive API** result from the list.
   - Click the blue **Enable** button.
   - Wait for it to enable. Once enabled, you’ll see a "Manage" button.

6. **Verify APIs are Enabled**
   - To confirm, go to **"APIs & Services"** → **"Dashboard"** in the left sidebar.
   - You should see both **Google Sheets API** and **Google Drive API** listed under "Enabled APIs" with some usage statistics (initially zero).

## Next Steps After Enabling APIs

Now that the APIs are enabled, you need to create credentials (a service account key) to use them in your Python script. Here’s how to proceed:

1. **Create a Service Account**
   - In the left sidebar, go to **"APIs & Services"** → **"Credentials"**.
   - Click **"Create Credentials"** at the top and select **"Service Account"**.
   - Fill in:
     - **Service account name**: e.g., "PropertyAnalysisService"
     - **Service account ID**: This will auto-fill based on the name.
   - Skip the optional "Grant this service account access" and "Grant users access" steps (click "Continue").
   - Click **"Done"**.

2. **Generate a JSON Key**
   - On the Credentials page, under "Service Accounts," click the email address of the service account you just created.
   - Go to the **"Keys"** tab.
   - Click **"Add Key"** → **"Create new key"**.
   - Select **"JSON"** as the key type and click **"Create"**.
   - The JSON file will download automatically to your computer.

3. **Update Your Script**
   - Move the downloaded JSON file to the same directory as your Python script.
   - Rename it if desired (e.g., "property-analysis-credentials.json").
   - Update the `CREDENTIALS_FILE` variable in your script to match the filename:
     ```python
     CREDENTIALS_FILE = 'property-analysis-credentials.json'
     ```
   - Update the USER_EMAIL variable with your actual email address:
    ```python
    USER_EMAIL = 'your-email@gmail.com'
    ```

# Troubleshooting
- If "Enable" is grayed out: The API might already be enabled. Check the Dashboard to confirm.
- If you get errors later: Ensure your project has billing enabled (Google Cloud requires it for some API usage, though these APIs have a free tier).
- Lost in the interface?: Use the search bar at the top of Google Cloud Console and type "Sheets API" or "Drive API" to jump directly to those pages.

Once you’ve enabled the APIs and set up the credentials, your script should work with Google Sheets.
