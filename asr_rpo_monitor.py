from azure.identity import ClientSecretCredential
from azure.mgmt.recoveryservicessiterecovery import SiteRecoveryManagementClient
import requests
import os
# ---------------------------------------------------
# ASR Service Principal Credentials (GitHub Secrets)
# ---------------------------------------------------
TENANT_ID = os.getenv("ASR_TENANT_ID")
CLIENT_ID = os.getenv("ASR_CLIENT_ID")
CLIENT_SECRET = os.getenv("ASR_CLIENT_SECRET")

# ---------------------------------------------------
# Microsoft Graph Mail Credentials (GitHub Secrets)
# ---------------------------------------------------
MAIL_TENANT_ID = os.getenv("MAIL_TENANT_ID")
MAIL_CLIENT_ID = os.getenv("MAIL_CLIENT_ID")
MAIL_CLIENT_SECRET = os.getenv("MAIL_CLIENT_SECRET")

# ---------------------------------------------------
# Email Configuration
# ---------------------------------------------------
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")
CC_EMAILS = os.getenv("CC_EMAILS")

# ---------------------------------------------------
# Azure Vault Details
# ---------------------------------------------------
SUBSCRIPTION_ID = "e8328d3b-7c5e-4aa5-b321-eeb887f1fc6b"
RESOURCE_GROUP = "AZW-RG-ASR"
VAULT_NAME = "AZW-RSV-ASR-01"


# ---------------------------------------------------
# Microsoft Graph Token
# ---------------------------------------------------
def get_graph_token():
    url = f"https://login.microsoftonline.com/{MAIL_TENANT_ID}/oauth2/v2.0/token"

    payload = {
        "client_id": MAIL_CLIENT_ID,
        "client_secret": MAIL_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }

    response = requests.post(url, data=payload)
    response.raise_for_status()

    return response.json()["access_token"]


# ---------------------------------------------------
# Send Email
# ---------------------------------------------------
def send_email(subject, body):
    token = get_graph_token()

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    to_recipients = [
        {
            "emailAddress": {
                "address": email.strip()
            }
        }
        for email in RECEIVER_EMAIL.split(",")
    ]

    cc_recipients = [
        {
            "emailAddress": {
                "address": email.strip()
            }
        }
        for email in CC_EMAILS.split(",")
    ]

    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body
            },
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients
        },
        "saveToSentItems": "true"
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload
    )

    if response.status_code == 202:
        print("Email sent successfully.")
    else:
        print("Email failed:", response.text)


# ---------------------------------------------------
# Azure Login
# ---------------------------------------------------
credential = ClientSecretCredential(
    tenant_id=TENANT_ID,
    client_id=CLIENT_ID,
    client_secret=CLIENT_SECRET
)

client = SiteRecoveryManagementClient(
    credential=credential,
    subscription_id=SUBSCRIPTION_ID,
    resource_group_name=RESOURCE_GROUP,
    resource_name=VAULT_NAME
)

print("Azure Login Successful")

# ---------------------------------------------------
# Fetch All Fabrics
# ---------------------------------------------------
fabrics = client.replication_fabrics.list()

email_body = f"""
<h2>Azure Site Recovery Alert Report</h2>
<p><strong>Vault Name:</strong> {VAULT_NAME}</p>
<hr>
"""

found = False

# ---------------------------------------------------
# Scan All Fabrics / Containers / Regions
# ---------------------------------------------------
for fabric in fabrics:
    fabric_name = fabric.name

    try:
        containers = client.replication_protection_containers.list_by_replication_fabrics(
            fabric_name
        )

        for container in containers:
            container_name = container.name

            replicated_items = client.replication_protected_items.list_by_replication_protection_containers(
                fabric_name=fabric_name,
                protection_container_name=container_name
            )

            for item in replicated_items:
                try:
                    vm_name = getattr(
                        item.properties,
                        "friendly_name",
                        item.name
                    )

                    replication_health = str(
                        item.properties.replication_health
                    ).strip().lower()

                    provider_details = item.properties.provider_specific_details

                    rpo_seconds = getattr(
                        provider_details,
                        "last_rpo_in_seconds",
                        None
                    )

                    if rpo_seconds is None:
                        rpo_seconds = getattr(
                            provider_details,
                            "rpo_in_seconds",
                            None
                        )

                    if rpo_seconds is None:
                        continue

                    rpo_hours = round(rpo_seconds / 3600, 2)

                    active_location = getattr(
                        provider_details,
                        "primary_azure_location",
                        None
                    )

                    if active_location is None:
                        active_location = getattr(
                            provider_details,
                            "recovery_azure_location",
                            None
                        )

                    if active_location is None:
                        active_location = item.properties.active_location

                    # Alert Condition
                    if rpo_hours >= 1 and replication_health == "critical":
                        found = True

                        print(f"VM Name: {vm_name}")
                        print(f"Fabric: {fabric_name}")
                        print(f"Container: {container_name}")
                        print(f"Replication Health: {item.properties.replication_health}")
                        print(f"Status: {item.properties.protection_state}")
                        print(f"Active Location: {active_location}")
                        print(f"RPO (Hour): {rpo_hours}")
                        print(f"Failover Health: {item.properties.failover_health}")
                        print("-" * 80)

                        email_body += f"""
                        <h3>VM Name: {vm_name}</h3>
                        <ul>
                            <li><strong>Fabric:</strong> {fabric_name}</li>
                            <li><strong>Container:</strong> {container_name}</li>
                            <li><strong>Replication Health:</strong> {item.properties.replication_health}</li>
                            <li><strong>Status:</strong> {item.properties.protection_state}</li>
                            <li><strong>Active Location:</strong> {active_location}</li>
                            <li><strong>RPO (Hour):</strong> {rpo_hours}</li>
                            <li><strong>Failover Health:</strong> {item.properties.failover_health}</li>
                        </ul>
                        <hr>
                        """

                except Exception as item_error:
                    print(f"Error processing item {item.name}: {str(item_error)}")

    except Exception as fabric_error:
        print(f"Error scanning fabric {fabric_name}: {str(fabric_error)}")


# ---------------------------------------------------
# Send Email
# ---------------------------------------------------
if found:
    send_email(
        subject=f"Critical Azure ASR Alert - Vault {VAULT_NAME}",
        body=email_body
    )
else:
    print("No critical ASR alerts found.")
