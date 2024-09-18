# General Overview
This script is designed to automate the process of generating inspection reports based on data stored in a Google Sheet. It's capable of handling both Visual and Functional inspection reports for train systems.

## Key Components

### Configuration
The script starts with a set of configuration variables that control various aspects of the report generation process, including:
- Data start row in the sheet
- Batch size for processing
- Maximum execution time
- Column mappings for data extraction

### Main Functions
1. **Report Triggers**:  
   Two functions (`generateVisualReport` and `generateFunctionalReport`) serve as entry points for generating Visual and Functional reports respectively.
   
2. **Data Processing**:  
   The `processInspectionData` function is the core of the script. It:
   - Fetches data from the specified sheet
   - Processes data in batches
   - Handles time tracking and logging

3. **Report Generation**:  
   The `generateReportsFromData` function:
   - Creates a new Google Doc from a template
   - Populates the document with data
   - Saves the document in the appropriate folder structure

### Helper Functions
- `appendTableToDocument`: Adds data to the report table
- `replaceTrainNoPlaceholder`: Updates the document header
- `createDocumentFromTemplate`: Creates a new doc from a template
- `saveDocumentToFolder`: Organizes the generated reports in folders

## Key Features
- **Batch Processing**:  
  The script processes data in batches to handle large datasets efficiently.
  
- **Dynamic Document Creation**:  
  It creates new Google Docs based on a template, ensuring consistent formatting.
  
- **Data Organization**:  
  The script organizes generated reports into folders based on inspection type and train number.
  
- **Image Handling**:  
  It attempts to fetch and insert images into the report, with error handling for invalid URLs.
  
- **Performance Tracking**:  
  The script logs start and end times, calculating the duration of the report generation process.
  
- **User Feedback**:  
  It provides alerts and updates specific cells in the sheet to keep the user informed of progress.

## Workflow
1. User triggers report generation from the Google Sheet.
2. Script fetches data from the sheet in batches.
3. For each batch, a new Google Doc is created from a template.
4. Data is inserted into the document, including text and images.
5. The document is saved in the appropriate folder structure.
6. Process repeats until all data is processed or time limit is reached.
7. User is notified of completion and provided with document links.

This script demonstrates a robust approach to automating report generation, handling large datasets, and organizing output efficiently.

# Google Sheet Columns
The script uses a column mapping that corresponds to specific data fields. The columns in the Google Sheet are:

- **Column B (index 2)**: Inspection ID
- **Column E (index 5)**: UserName (PIC)
- **Column G (index 7)**: Train Number
- **Column H (index 8)**: Location
- **Column K (index 11)**: Car Body
- **Column M (index 13)**: Section Name
- **Column O (index 15)**: Subsystem Name
- **Column P (index 16)**: Serial Number
- **Column R (index 18)**: Subcomponent
- **Column S (index 19)**: Condition
- **Column T (index 20)**: Defect Type
- **Column U (index 21)**: Remarks
- **Column AA (index 27)**: Image URL

## Data Populated into Google Docs
The script creates a table in the Google Doc and populates it with the following information:

- **No (Item No)**: A dynamically generated item number
- **Loc**: Location (from Column H)
- **Car**: Car Body (from Column K)
- **PIC**: UserName (from Column E)
- **Section**: Section Name (from Column M)
- **Sub System**: Subsystem Name (from Column O)
- **Serial No**: Serial Number (from Column P)
- **Sub Component**: Subcomponent (from Column R)
- **Condition**: Condition (from Column S)
- **Defect**: Defect Type (from Column T)
- **Remarks**: Remarks (from Column U)
- **Image**: An image cell where the script attempts to insert an image from the URL provided in Column AA

## Header Replacement
Additionally, the script replaces a placeholder in the document header:

- `{{trainNo}}`: This is replaced with the actual train number (from Column G) and the inspection type (either "Functional Inspection" or "Visual Inspection" based on the report type).

## Report Organization
The script organizes the generated reports into folders based on the inspection type (Visual or Functional) and the train number. It also provides feedback to the user by updating specific cells in the Google Sheet with information about the generated report, such as the file name, document link, and the last processed item number.
