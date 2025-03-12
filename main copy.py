import re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import os
import base64
from googleapiclient.discovery import build
import openpyxl

# Gmail API Scopes (adjust as needed)
SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels',
    'https://www.googleapis.com/auth/gmail.readonly'
]

def get_gmail_service():
    print("Initializing Gmail service...")
    creds = None
    token_path = 'resources/token.json'
    credentials_path = 'resources/credentials.json'

    print(f"Checking for existing token at {token_path}")
    if os.path.exists(token_path):
        print("Found existing token, loading credentials...")
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    else:
        print("No existing token found.")
    
    if not creds or not creds.valid:
        print("Credentials are invalid or expired.")
        if creds and creds.expired and creds.refresh_token:
            print("Attempting to refresh expired token...")
            creds.refresh(Request())
            print("Token refreshed successfully.")
        else:
            print("Need to generate new token...")
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_path,
                SCOPES
            )
            creds = flow.run_local_server(port=0)
            print("New token generated.")
        
        print(f"Saving token to {token_path}")
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
    print("Building Gmail service...")
    service = build('gmail', 'v1', credentials=creds)
    print("Gmail service initialized successfully!")
    return service

def load_rules_from_excel(file_path):
    """Load rules from an Excel file.
    Column structure:
    1: Name (irrelevant)
    2: Description (irrelevant)
    3: Category (irrelevant)
    4: Rule Type
    5: Condition Type 1
    6: Condition Value 1
    7: Condition Type 2
    8: Condition Value 2
    9: Condition Type 3
    10: Condition Value 3
    11: Action
    12: Action Value
    """
    print("Loading rules from Excel...")
    rules = []
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):  # Skip header
        # Skip if rule type is None or empty
        rule_type = row[3]
        if not rule_type or (isinstance(rule_type, str) and rule_type.lower() == 'none'):
            continue

        rule = {
            'name': str(row[0]) if row[0] else f"Rule_{row_idx}",
            'rule_type': str(rule_type).lower() if rule_type else None,  # Convert to lowercase
            'conditions': [],
            'action': str(row[10]).lower() if row[10] and str(row[10]).lower() != 'none' else None,
            'action_value': str(row[11]) if row[11] and str(row[11]).lower() != 'none' else None
        }

        # Process conditions in pairs (type and value)
        condition_pairs = [
            (row[4], row[5]),  # Condition 1
            (row[6], row[7]),  # Condition 2
            (row[8], row[9])   # Condition 3
        ]

        # Add valid conditions
        for cond_type, cond_value in condition_pairs:
            if cond_type and cond_value and str(cond_type).lower() != 'none':
                rule['conditions'].append({
                    'type': str(cond_type).lower(),  # Convert to lowercase
                    'value': str(cond_value).lower()  # Convert to lowercase
                })
        
        print(f"Loaded rule: {rule['name']}")
        print(f"  Rule Type: {rule['rule_type']}")
        print(f"  Conditions: {len(rule['conditions'])}")
        print(f"  Action: {rule['action']} -> {rule['action_value']}")
        rules.append(rule)
    
    return rules

def evaluate_expression(expression, subject, sender):
    """Evaluate a logical expression with AND, OR, XOR, NOT."""
    # Replace placeholders with actual values
    expression = expression.replace('subject', f'"{subject.lower()}"')
    expression = expression.replace('sender', f'"{sender.lower()}"')

    # Evaluate the expression safely
    try:
        return eval(expression)
    except:
        return False

def parse_condition_value(condition_value, subject, sender):
    """Parse condition value with or without logical operators."""
    if not condition_value:
        return False

    # Check if the condition value contains logical operators
    if condition_value.startswith('[') and condition_value.endswith(']'):
        # Extract the expression inside the brackets
        expression = condition_value[1:-1]
        return evaluate_expression(expression, subject, sender)
    else:
        # Simple condition (no operators)
        return condition_value.lower() in subject.lower() or condition_value.lower() in sender.lower()

def evaluate_conditions(conditions, subject, sender, labels=None, service=None):
    """Evaluate all conditions for a rule (AND logic)."""
    # Convert email data to lowercase for comparison
    subject_lower = subject.lower()
    sender_lower = sender.lower()
    
    for condition in conditions:
        condition_type = condition['type'].lower()  # Ensure lowercase
        condition_value = condition['value'].lower()  # Ensure lowercase
        
        # Handle tag condition type
        if condition_type == 'tag':
            # If we don't have labels info, we can't evaluate tag conditions
            if labels is None or service is None:
                return False
                
            # Get label ID for the requested tag
            tag_label_id = get_or_create_label(service, condition_value)
            
            # Check if the email has this label
            if tag_label_id not in labels:
                return False
                
        # Check if this is a logical expression (enclosed in square brackets)
        elif condition_value.startswith('[') and condition_value.endswith(']'):
            # Remove brackets and evaluate the logical expression
            expression = condition_value[1:-1]
            field_text = subject_lower if condition_type == 'subject' else sender_lower
            
            if not evaluate_logical_expression(expression, field_text):
                return False
        else:
            # Regular single-value condition
            if condition_type == 'sender':
                if condition_value not in sender_lower:
                    return False
            elif condition_type == 'subject':
                if condition_value not in subject_lower:
                    return False
    
    return True  # All conditions passed

def evaluate_logical_expression(expression, text):
    """Evaluate a logical expression with AND, OR, XOR, NOT operators and parentheses."""
    # Simple case: just OR conditions
    if ' or ' in expression and ' and ' not in expression and ' xor ' not in expression and ' not ' not in expression and '(' not in expression:
        keywords = [k.strip().strip('"\'') for k in expression.split(' or ')]
        return any(keyword in text for keyword in keywords)
    
    # Simple case: just AND conditions
    if ' and ' in expression and ' or ' not in expression and ' xor ' not in expression and ' not ' not in expression and '(' not in expression:
        keywords = [k.strip().strip('"\'') for k in expression.split(' and ')]
        return all(keyword in text for keyword in keywords)
    
    # For more complex expressions, we need a proper parser
    # This is a recursive descent parser for logical expressions
    def parse_expression(expr):
        expr = expr.strip()
        
        # Base case: simple keyword search
        if '"' in expr and ' and ' not in expr and ' or ' not in expr and ' xor ' not in expr and ' not ' not in expr and '(' not in expr:
            keyword = expr.strip('"\'')
            return keyword in text
        
        # Handle parentheses
        if expr.startswith('('):
            # Find matching closing parenthesis
            count = 1
            for i in range(1, len(expr)):
                if expr[i] == '(':
                    count += 1
                elif expr[i] == ')':
                    count -= 1
                    if count == 0:
                        # Parse the subexpression
                        result = parse_expression(expr[1:i])
                        # If there's anything after the closing parenthesis, it must be an operator
                        if i+1 < len(expr):
                            rest = expr[i+1:].strip()
                            if rest.startswith('and '):
                                return result and parse_expression(rest[4:])
                            elif rest.startswith('or '):
                                return result or parse_expression(rest[3:])
                            elif rest.startswith('xor '):
                                return bool(result) != bool(parse_expression(rest[4:]))
                        return result
        
        # Handle NOT operator
        if expr.startswith('not '):
            return not parse_expression(expr[4:])
        
        # Find the first top-level operator
        in_quotes = False
        in_parens = 0
        for i in range(len(expr)):
            if expr[i] == '"':
                in_quotes = not in_quotes
            elif expr[i] == '(' and not in_quotes:
                in_parens += 1
            elif expr[i] == ')' and not in_quotes:
                in_parens -= 1
            elif in_parens == 0 and not in_quotes:
                # Check for operators
                if expr[i:].startswith('and '):
                    left = parse_expression(expr[:i].strip())
                    right = parse_expression(expr[i+4:].strip())
                    return left and right
                elif expr[i:].startswith('or '):
                    left = parse_expression(expr[:i].strip())
                    right = parse_expression(expr[i+3:].strip())
                    return left or right
                elif expr[i:].startswith('xor '):
                    left = parse_expression(expr[:i].strip())
                    right = parse_expression(expr[i+4:].strip())
                    return bool(left) != bool(right)
        
        # If we get here, it's a simple keyword
        keyword = expr.strip('"\'')
        return keyword in text
    
    # Start the parsing process
    return parse_expression(expression)

def get_or_create_label(service, label_name):
    """Get existing label ID or create a new one."""
    try:
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        
        for label in labels:
            if label['name'].lower() == label_name.lower():
                return label['id']

        # Create new label if not found
        new_label = {
            'name': label_name,
            'labelListVisibility': 'labelShow',
            'messageListVisibility': 'show'
        }
        created = service.users().labels().create(userId='me', body=new_label).execute()
        return created['id']
    
    except Exception as e:
        print(f"Error handling labels: {e}")
        return None

def handle_rule_action(service, message_id, rule):
    """Execute the action specified in the rule."""
    action = rule['action']
    action_value = rule['action_value']

    if not action or action.lower() == 'none':
        print("No action specified")
        return

    try:
        # Convert action to lowercase for comparison but keep action_value as-is
        # for display purposes and label creation
        action_lower = action.lower()
        
        if action_lower == 'label':
            label_id = get_or_create_label(service, action_value)
            if label_id:
                service.users().messages().modify(
                    userId='me',
                    id=message_id,
                    body={'addLabelIds': [label_id]}
                ).execute()
                print(f"Applied label: {action_value}")

        elif action_lower == 'move':
            label_id = get_or_create_label(service, action_value)
            if label_id:
                service.users().messages().modify(
                    userId='me',
                    id=message_id,
                    body={
                        'addLabelIds': [label_id],
                        'removeLabelIds': ['INBOX']
                    }
                ).execute()
                print(f"Moved email to: {action_value}")

        elif action_lower == 'mark_important':
            should_mark = str(action_value).lower() == 'true'
            service.users().messages().modify(
                userId='me',
                id=message_id,
                body={
                    'addLabelIds': ['IMPORTANT'] if should_mark else [],
                    'removeLabelIds': ['IMPORTANT'] if not should_mark else []
                }
            ).execute()
            print(f"Marked email as {'important' if should_mark else 'not important'}")

        elif action_lower == 'star':
            should_star = str(action_value).lower() == 'true'
            service.users().messages().modify(
                userId='me',
                id=message_id,
                body={
                    'addLabelIds': ['STARRED'] if should_star else [],
                    'removeLabelIds': ['STARRED'] if not should_star else []
                }
            ).execute()
            print(f"{'Starred' if should_star else 'Unstarred'} email")

        elif action_lower == 'archive':
            service.users().messages().modify(
                userId='me',
                id=message_id,
                body={'removeLabelIds': ['INBOX']}
            ).execute()
            print("Archived email")

        elif action_lower == 'delete':
            service.users().messages().trash(userId='me', id=message_id).execute()
            print("Moved email to trash")

    except Exception as e:
        print(f"Error performing action {action}: {str(e)}")

def process_emails():
    print("\n=== Starting Email Processing ===")
    
    # Get Gmail service
    service = get_gmail_service()
    
    # Load rules from Excel
    rules = load_rules_from_excel('resources/rules.xlsx')
    print(f"Loaded {len(rules)} rules from Excel")
    
    # Get Algorithm Reviewed label
    algorithm_reviewed_label = get_or_create_label(service, 'Algorithm Reviewed')
    print(f"DEBUG: Algorithm Reviewed label ID: {algorithm_reviewed_label}")
    
    # Get all inbox emails - we'll filter them in the loop
    query = 'in:inbox -label:"Algorithm Reviewed"'
    print(f"DEBUG: Query: {query}")
    results = service.users().messages().list(userId='me', q=query).execute()
    messages = results.get('messages', [])
    
    if not messages:
        print("No new emails to process")
        return
    
    total_messages = len(messages)
    print(f"Found {total_messages} emails to check")
    
    rules_applied_count = 0
    
    for index, message in enumerate(messages, 1):
        print(f"\nProcessing email {index}/{total_messages}")
        msg = service.users().messages().get(
            userId='me', id=message['id'], format='metadata'
        ).execute()
        headers = msg['payload']['headers']
        
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
        sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown Sender')
        
        print(f"Email from: {sender}")
        print(f"Subject: {subject}")
        
        # Get current labels to debug
        msg_labels = service.users().messages().get(userId='me', id=message['id'], format='metadata').execute().get('labelIds', [])
        print(f"DEBUG: Current message labels: {msg_labels}")
        
        # Check if this email has the pending label
        pending_label = get_or_create_label(service, 'Algorithm Reviewed [Pending]')
        has_pending_label = pending_label in msg_labels
        
        if has_pending_label:
            print("Email has 'Algorithm Reviewed [Pending]' label, only specific rules will be applied")
        
        # Evaluate rules
        rule_applied = False
        for rule_index, rule in enumerate(rules, 1):
            print(f"Checking rule {rule_index}: {rule['name']}")
            
            # If the email has pending label, only process rules that have a tag condition for it
            if has_pending_label:
                # Check if any condition in this rule is looking for the pending tag
                has_pending_condition = any(
                    c['type'].lower() == 'tag' and 
                    c['value'].lower() == 'algorithm reviewed [pending]'
                    for c in rule['conditions']
                )
                
                if not has_pending_condition:
                    print(f"  Skipping rule - email has pending label but rule doesn't look for it")
                    continue
            
            if evaluate_conditions(rule['conditions'], subject, sender, msg_labels, service):
                print(f"✓ Rule '{rule['name']}' matched!")
                
                handle_rule_action(service, message['id'], rule)
                rule_applied = True
                rules_applied_count += 1
                
                # Debug the condition values
                action_lower = rule['action'].lower()
                action_value_lower = rule['action_value'].lower()
                
                # Add "Algorithm Reviewed" label unless the action was to label as "Algorithm Reviewed [Pending]"
                if not (action_lower == 'label' and action_value_lower == 'algorithm reviewed [pending]'):
                    modify_result = service.users().messages().modify(
                        userId='me',
                        id=message['id'],
                        body={'addLabelIds': [algorithm_reviewed_label]}
                    ).execute()
                    print("Added 'Algorithm Reviewed' label")
                else:
                    print("Skipped adding 'Algorithm Reviewed' label due to '[Pending]' status")
                
                break  # Stop after the first matching rule
            else:
                print(f"✗ Rule '{rule['name']}' did not match")
        
        if not rule_applied:
            print("No rules matched for this email")
            
    print("\n=== Email Processing Complete ===")
    print(f"Applied rules to {rules_applied_count} emails")

if __name__ == "__main__":
    process_emails()
    print("Email processing complete!")