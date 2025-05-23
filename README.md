# **The eHub Data + Systems Project: Scheduling, Matching, and Availability** 

## **Overview**

This project supports a suite of tools designed to:

1. **Schedule students** into tracks or events based on their availability.  
2. **Query students** available at a given time with particular interests.  
3. **Match students** into complementary teams using structured input data.  
4. **Integrate with HubSpot** to enhance student records with additional properties.
5. **Detect and visualize student communities** based on shared interests (via a standalone Python script).

Additionally, the project contains utility and helper functions that support general sheet operations, filtering, clustering, and data extraction.

The Google Apps Script code can be edited directly by going to [this sheet](https://docs.google.com/spreadsheets/d/10EtGplen7xTupr5-WehA1tfHkR8_mZJ3NQuv2WeQ0Yg/edit?usp=sharing) (access must be given first) and clicking Extensions \> Apps Script in the navigation bar.  

The project is organized into two main directories:
* `apps-script/` — Contains all Google Apps Script .gs and .html files managed via CLASP, grouped by functionality.
* `other-scripts/` — Contains supplemental tools, currently the Python script used for interest-based community detection.

---

## **File Structure and Functional Areas**

### **1\. Track/Event Scheduling (Availability-Based)**

These files contain logic to parse student availability from forms or sheets, determine optimal time slots, and write scheduling outputs:

* **scheduling-tracks.gs** – Contains core scheduling logic, including functions to extract availability from forms, identify optimal time slots, and write results to sheets. It supports both track-based and custom group-based scheduling, handles virtual student flags, and includes a main runner function for batch processing.  
  Major functions include:  
  * `getEmailsFromSheet()` and `getEmailsWithCondition()` – Extract email addresses from a sheet, with optional filtering by a "Yes"/"No" condition column.  
  * `getEmailsByTrack()` – Retrieves emails of students associated with a specific track (supports comma-separated multiple tracks).  
  * `findAvailableSlots()` and `findAvailableSlotsForEmails()` – Builds a mapping of available time slots to students, optionally filtering out slots below a threshold.  
  * `getMinimumTimeSlots()` – Selects the smallest number of time slots needed to cover all students at least once.  
  * `getIterativeTimeSlots()` – Iteratively selects the most popular available time slot, then removes those students and repeats until all students are covered.  
  * `writeSelectedSlotsToSheet()` – Writes availability information and selected time slots to a new or existing sheet.  
  * `processAvailabilityByTrack()` and `processAvailabilityBySheet()` – Convenience functions to run scheduling logic based on either a track or a manually specified list of students.  
  * `main()` – Example driver function for batch processing of common use cases.

### **2\. Time \+ Interest Queries**

This section supports finding students available at a specific time and day who also match one or more interests. This can be run directly from the sheet by clicking Activities \> Find Available Students in the navigation bar. Results can be output to a sheet, and each query is logged:

* **setup.gs** – Adds a custom menu to the spreadsheet UI and displays a sidebar for filtering students based on time and interest criteria.  
* **sidebar.html** – HTML interface rendered as a sidebar in the spreadsheet, allowing users to select a day, time, and one or more interests. Passes form data to `processFilters()` to generate query results.  
* **interest-availability-query.gs** – Handles backend logic for filtering students by time and interest. Contains:  
  * `preprocessStudentData()` to combine availability and interest data into a cached object.  
  * `processFilters()` to extract students based on sidebar inputs.  
  * `logQuery()` to record queries and generate unique sheet names for output.

### **3\. Team Matching (Skills \+ Interest-Based)**

* **data-driven-matchmaking.gs** – Implements the core matchmaking logic for forming teams:  
  * `updateMembershipColumn()` enriches student data with track membership and removes unmatched entries.  
  * `validateSheet()` ensures the dataset has all required columns for matchmaking.  
  * `separateLeadersAndMembers()` splits students into leaders (those seeking teammates) and members (those seeking to join a team).  
  * `findTopMatches()` and `calculateMatchScore()` identify top candidates for each leader based on shared interests and complementary skills.  
  * `runMatchmaking()` combines the full process and writes results to the "Matches" sheet.  
* **match-docs.gs** – Contains functions to generate Google Docs summarizing matches:  
  * `createMatchDocs()` generates a document for each team leader that includes their top 5 matches and detailed info about each teammate.  
  * `createMemberMatchDocs()` does the reverse: for each member, it summarizes matched team leaders.  
  * Both functions output document links to separate sheets for easy access.  
  * `docsMain()` runs both generation functions in sequence.

### **4\. HubSpot Integration**

* **hubspot-integration.gs** – Allows for enriching the student dataset with additional information pulled from HubSpot:  
  * `updatePropertiesInSheet(properties)` fetches values for specified HubSpot contact properties and writes them into a Google Sheet (assumes one row per student, matched by email).  
  * `getPropertiesFromHubSpot(email, properties)` makes a call to the HubSpot CRM API to retrieve the requested properties for a given contact email.  
  * `integrateMain()` is an example function call to populate major/minor fields.

### **5\. Community Detection (Python-based)**

* **community\_detection.py** (local Python script) – Provides logic for identifying natural student communities based on shared interests:  
  * Loads interest data from `student_data.json` and filters target students via a CSV.  
  * Builds a co-occurrence graph of shared interests and applies **Louvain community detection**.  
  * Assigns students to clusters based on which group of interests they align with most.  
  * Visualizes the results as:  
    * A network graph showing clustered interests  
    * A heatmap of co-occurrence  
    * Summary tables of top interests per group  
    * Email lists per community/track for outreach or scheduling  
* Outputs: `student_interest_clusters.csv` containing all cluster assignments.

### **6\. Miscellaneous Utilities**

* **kmeans.gs** – Provides basic K-Means clustering functionality:  
  * `kMeansClustering(data, k)` clusters vectors (e.g., numeric encodings of preferences or skills) into `k` groups.  
  * Internal helper functions (`initializeCentroids`, `assignCluster`, `updateCentroids`, `euclideanDistance`) support the core clustering logic.  
  * This can be used for exploratory grouping or segmenting students based on numeric profiles.  
* **useful-functions.gs** – Contains general helper functions for common operations:  
  * `getTopOccurrences()` – Returns the most common items in an array.  
  * `getRowByEmail()` / `getNameByEmail()` – Retrieve full student rows or names using email as a key.  
  * `getEmailsFromSheetCell()` – Parses comma-separated emails from a single cell.  
  * `exportStudentDataToDrive()` – Writes a saved JSON object to Google Drive for offline analysis.

---

## **Working with CLASP (Command Line Apps Script)**

To manage this Apps Script project with GitHub and edit locally, we use [CLASP](https://github.com/google/clasp).

### **1\. Setup**

```
npm install -g @google/clasp
clasp login
```

### **2\. Clone an Existing Project**

If you're starting from a Google Apps Script that already exists:

```
clasp clone <script-id>
```

You can get the `script-id` from the Apps Script project URL. Example:

```
https://script.google.com/home/projects/1abc1234567XYZ/edit
                                        ^^^^^^^^^^^^^^
```

### **3\. Pull Latest Changes**

```
clasp pull
```

This syncs code from Google Apps Script into your local folder.

### **4\. Push Local Changes**

```
clasp push
```

⚠️ If you see a "project too large" error, try splitting code into multiple files or reducing embedded data.

### **5\. Edit Code Locally \+ GitHub Integration**

* Make changes using your code editor.  
* Use `git add`, `git commit`, and `git push` to version control your script.  
* Use `clasp push` to update the script in Google Apps Script.

---

## **How to Use This Script**

1. **Set Up**  
   * Attach this script project to a Google Sheet with student availability and interest form responses.  
   * Ensure consistent column names across forms (e.g., "Email Address", "Interests", etc.). ⚠️ Some functions are highly dependent on sheet names, column names, and column locations within a sheet. Be sure to check for this before running scripts to prevent logical errors. ⚠️  
2. **Running Scheduling Logic**  
   * Use `processAvailabilityByTrack()` or `processAvailabilityBySheet()` to generate track- or group-specific availability slots.  
   * Use `selectTimeSlots()` or `getIterativeTimeSlots()` to generate the minimum number of slots or most efficient time slot assignments.  
   * View examples in the `main()` function for how to use batch commands.  
3. **Running Interest-Based Queries**  
   * Use the "Activities" menu in the spreadsheet and select **Find Available Students** to launch the sidebar UI.  
   * The sidebar (`sidebar.html`) allows filtering students by day, time, and interests.  
   * Results are written to a uniquely named sheet and logged via `logQuery()`.  
   * Use `preprocessStudentData()` in `interest-availability-query.gs` to cache updated student data from forms.  
4. **Running Team Matching**  
   * Run `runMatchmaking()` to generate leader-member matches.  
   * Use `createMatchDocs()` and `createMemberMatchDocs()` to generate detailed Google Docs summarizing the matches.  
5. **Syncing HubSpot Data**  
   * Run `integrateMain()` with the desired HubSpot properties to pull.  
   * Make sure your script has the correct API token for HubSpot authentication.  
6. **Updating Preprocessed Data**  
   * Run `preprocessStudentData()` to refresh stored availability and interest data.  
7. **Python Community Detection**  
   * A separate Python file (`community_detection.py`) can be used to detect natural student groupings based on overlapping interests.  
   * Outputs include visual graphs, cluster assignments, and CSV exports.
  
---

## **Managing Secrets (e.g., HubSpot API Key)**

To keep private keys like your HubSpot API token secure, avoid hardcoding them directly into version-controlled code.

### **✅ Recommended Approach: Use Script Properties**

Store secrets securely using Apps Script’s built-in property service:

1. **Set the key manually** (run once from the script editor):

```
PropertiesService.getScriptProperties().setProperty('HUBSPOT_API_KEY', 'your-api-key-here');
```

2.   
   **Access it securely in your script**:

```
const apiKey = PropertiesService.getScriptProperties().getProperty('HUBSPOT_API_KEY');
```

3.   
   **Remove hardcoded secrets** before committing to GitHub.

Script Properties are stored securely and not visible in your codebase, making this the ideal method for managing authentication tokens.
