# Microsoft Graph Toolbox

A Python package that simplifies using Microsoft Graph API. 
It provides an easy-to-use interface for authenticating with Microsoft Graph and using its capabilities.
The code expects a logger object with .debug/.info/.warning/.error methods.

## Current Features
- **Email Operations** (graph_email): Send emails with attachments, HTML content, and more
- **User Queries** (graph_users): Search and retrieve user profiles with flexible filters and projection
- **SharePoint Operations** (graph_sharepoint): Upload files, list folder contents

## General Prerequisites

- Python 3.12 or higher (lower versions may work but are not tested)
- Azure AD registered application with:
  - Client ID
  - Client Secret
- Required Microsoft Graph API permissions depending on feature:
  - Email operations: Mail.Send
  - User queries: User.Read, User.Read.All
  - SharePoint operations: Sites.Read.All, Sites.ReadWrite.All, Files.ReadWrite.All

## Installation

```bash
# Clone the repository
git clone https://github.com/runway28R/ms-graph.git
cd ms-graph

# Install required packages
pip install -r requirements.txt
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.


### Function graph_email

#### Prerequisites
  - Required Microsoft Graph API permissions (Mail.Send)

#### Features

- App-only authentication using client credentials flow
- Send emails with HTML or plain text content
- Support for attachments (both regular and inline)
- CC and BCC recipient support
- Email priority settings
- Logging capabilities

#### Example Usage

```bash
# Send HTML email
python -m examples.sending_email \
    --client-id CLIENT_ID \
    --client-secret CLIENT_SECRET \
    --tenant-id TENANT_ID \
    --sender from@someone.com \
    --to friend@another.com \
    --content-type HTML \
    --body "<h1>Hello!</h1><p>This is a test email.</p>"
```

#### Features in Detail

- **Content Types**: Support for both HTML and plain text
- **Attachments**: Send files and inline images
- **Recipients**: TO, CC, and BCC fields supported
- **Priority**: Set email importance (low, normal, high)
- **Logging**: Built-in logging for debugging and monitoring


### Function graph_users

#### Prerequisites
  - Required Microsoft Graph API permissions (User.Read, User.Read.All)

#### Features

- Search and retrieve user profiles with flexible filters and projection
- Get user details including email, name, and profile information
- Support for pagination
- Filtering by attributes (e.g., email, name, department)
- Projection of specific fields (e.g., email, name)

#### Example Usage

```bash
# Search for users
python -m examples.searching_users \
    --client-id CLIENT_ID \
    --client-secret CLIENT_SECRET \
    --tenant-id TENANT_ID \
    --search_name "John Doe" \
    --select_data givenName,displayName
```

#### Features in Detail

- **Filters**: Search by name, email, company, etc.
- **Fields**: Retrieve specific fields (e.g., displayName,givenName,surname,country,department,jobTitle,companyName,mail,accountEnabled)
- **Pagination**: Support for retrieving large datasets
- **Projection**: Select specific fields to return

### Function graph_sharepoint

#### Prerequisites
  - Required Microsoft Graph API permissions: Sites.Read.All, Sites.ReadWrite.All, Files.ReadWrite.All

#### Features

- Upload local files to SharePoint document libraries
- List contents of folders (separating files and subfolders)
- Support for specifying folder paths within document libraries
- Logging of folder content and upload results
- App-only authentication using client credentials flow
- Handles errors like missing files, invalid site URLs, or non-existing libraries
- Optional display of folder content before uploading

#### Example Usage

```bash
# Upload a file to SharePoint
python -m examples.upload_file \
    --client-id CLIENT_ID \
    --client-secret CLIENT_SECRET \
    --tenant-id TENANT_ID \
    --site_url "contoso.sharepoint.com:/sites/TeamSite" \
    --local_file_path "./report.xlsx" \
    --document_library "Documents" \
    --folder_path "Reports/2025"
```


#### Features in Detail

- **Folders vs Files**: Automatically lists top-level folders and files in the target library
- **Upload**: Uploads single files to specified folder, returns the SharePoint URL
- **Folder path**: Optional; if not provided, uploads to root of library
- **Error handling**: Returns clear messages if the file is missing or library/folder doesn’t exist
- **Logging**: Logs debug info for folder content and upload results
