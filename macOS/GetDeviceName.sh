#!/bin/bash

XML_FILE="/Library/Application Support/MXB/Backup Manager/statusreport.xml"

if [[ ! -f "$XML_FILE" ]]; then
    echo "XML file not found: $XML_FILE"
    exit 1
fi

# Parse the account name using xmllint (assumes account name is in <Account> tag)
ACCOUNT_NAME=$(xmllint --xpath 'string(//Account)' "$XML_FILE" 2>/dev/null)

if [[ -z "$ACCOUNT_NAME" ]]; then
    echo "Account name not found in XML."
    exit 2
fi

echo "Cove Device Name: $ACCOUNT_NAME"

exit 0
