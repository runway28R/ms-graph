# Microsoft Graph Toolbox

A Python package that simplifies using Microsoft Graph API. 
It provides an easy-to-use interface for authenticating with Microsoft Graph and using its capabilities.

## Current Features
- **Email Operations** (graph_email): Send emails with attachments, HTML content, and more
- **User Queries** (graph_users): Search and retrieve user profiles with flexible filters and projection

## General Prerequisites

- Python 3.12 or higher (lower versions may work but are not tested)
- Azure AD registered application with:
  - Client ID
  - Client Secret

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/ms-graph.git
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

