from google_auth_oauthlib.flow import InstalledAppFlow
import os

# Define the scopes
SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels',
    'https://www.googleapis.com/auth/gmail.readonly'
]

def generate_token():
    """Generate the initial token for Gmail API access."""
    print("Starting token generation process...")
    credentials_path = 'resources/credentials.json'
    token_path = 'resources/token.json'
    
    # Ensure the resources directory exists
    print(f"Checking if resources directory exists...")
    os.makedirs('resources', exist_ok=True)
    
    if not os.path.exists(credentials_path):
        print(f"Error: {credentials_path} not found!")
        print("Please download your credentials.json file from Google Cloud Console and place it in the resources folder.")
        return
    else:
        print(f"Found credentials file at {credentials_path}")
    
    try:
        print("Creating OAuth flow...")
        # Create the flow using the client secrets file
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
        
        print("Starting local server for authentication...")
        print("Please check your web browser. If no browser opens, check the console for the URL...")
        # Run the OAuth flow locally
        creds = flow.run_local_server(
            port=0,
            authorization_prompt_message='Please visit this URL to authorize this application: {url}',
            success_message='The auth flow has completed! You may close this window.',
            open_browser=True
        )
        
        print("Authentication successful! Saving token...")
        # Save the credentials for future use
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
            
        print(f"Successfully generated token and saved to {token_path}")
        print("Token generation complete! You can now run your main script.")
        
    except Exception as e:
        print(f"An error occurred during token generation: {str(e)}")
        print("Debug info:")
        print(f"- Credentials path: {credentials_path}")
        print(f"- Token path: {token_path}")
        print(f"- Scopes: {SCOPES}")

if __name__ == "__main__":
    print("=== Gmail API Token Generator ===")
    generate_token() 