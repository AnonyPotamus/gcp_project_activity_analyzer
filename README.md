A Python script that finds the last activity date for all Google Cloud Platform projects within a specified folder hierarchy. The script checks activity across multiple GCP services and generates a comprehensive report to help identify inactive projects.
Features

Recursively scans all projects within a folder hierarchy
Checks for activity across multiple services:

Compute Engine (VM instances)
Cloud Storage (buckets and objects)
General API usage


Identifies the most recent activity date for each project
Calculates days since last activity
Reports access issues or disabled APIs
Outputs results to a color-coded Excel spreadsheet
Uses multi-threading for faster execution

Prerequisites

Python 3.6 or later
Google Cloud SDK installed and configured
Required Python packages:
google-cloud-storage
google-cloud-resource-manager
google-api-python-client
openpyxl


Installation

Ensure you have the Google Cloud SDK installed and configured
Install the required Python packages:
bashpip install google-cloud-storage google-cloud-resource-manager google-api-python-client openpyxl

Set up authentication:
bashgcloud auth application-default login


Usage
bashpython3 project-activity.py <folder_id> [--output OUTPUT_FILE] [--max_workers MAX_WORKERS]
Arguments

folder_id: Your GCP folder ID (required)
--output: (Optional) Output Excel file name (default: project_activity.xlsx)
--max_workers: (Optional) Maximum number of worker threads (default: 10)

Example
bashpython3 project-activity.py 123456789012 --output activity_report.xlsx --max_workers 20
This command will scan all projects in the folder with ID 123456789012 (and its subfolders), using 20 worker threads, and save the results to activity_report.xlsx.
Output
The script generates an Excel file with the following information:

Project ID
Last Activity Date
Activity Source (which service showed the most recent activity)
Days Since Activity
Access Issues (if any)

Projects with access issues or disabled APIs are highlighted in yellow for easy identification.
Troubleshooting
Authentication Issues
If you encounter authentication problems, ensure you've run:
bashgcloud auth application-default login
Permissions
The service account or user running the script needs the following permissions:

resourcemanager.folders.get
resourcemanager.folders.list
resourcemanager.projects.get
resourcemanager.projects.list
compute.instances.list
storage.buckets.list
storage.objects.list
serviceusage.services.list

These permissions need to be granted at the folder level to allow recursive scanning.
API Enablement
The script will attempt to check the following APIs in each project:

Cloud Resource Manager API
Compute Engine API
Cloud Storage API
Service Usage API

If these APIs are not enabled in a project, the script will report this as an access issue.
Limitations

The script may not detect all types of activity in a project. It focuses on compute, storage, and general API usage.
The script can't access projects where you don't have sufficient permissions.
For very large folders with many projects, the script may take a long time to run.
