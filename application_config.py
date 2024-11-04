import os
AUTHORITY= os.getenv("AUTHORITY")

# Application (client) ID of app registration
CLIENT_ID = os.getenv("CLIENT_ID")
# Application's generated client secret: never check this into source control!
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

TENANT_ID = os.getenv("TENANT_ID")

THUMBPRINT = os.getenv("THUMBPRINT")
KEY_VAULT_URL = os.getenv('KEY_VAULT_URL')
 
REDIRECT_PATH = "/auth"  # Used for forming an absolute URL to your redirect URI.

ENDPOINT = 'https://graph.microsoft.com/v1.0/me'  
ENDPOINT_SHAREPOINT = 'https://accantec.sharepoint.com/.default'

# Endpoint for sending emails
#EMAIL_SEND_ENDPOINT = 'https://graph.microsoft.com/v1.0/me/sendMail'

# Scope for both user information and email sending permissions
SCOPE = ["User.Read","Sites.ReadWrite.All","Sites.Manage.All"]

# Tells the Flask-session extension to store sessions in the filesystem
SESSION_TYPE = "filesystem"
