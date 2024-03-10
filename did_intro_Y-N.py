import os
import pandas as pd
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Retrieve your Slack API token and channel ID from environment variables
slack_token = os.getenv("SLACK_TOKEN")
channel_id = os.getenv("CHANNEL_ID")
client = WebClient(token=slack_token)

# Path to your Excel file
excel_path = "user_record.xlsx"

# Load the Excel file
df = pd.read_excel(excel_path)

try:
    # Fetch all user data in the channel for efficiency
    response = client.users_list()
    users_data = {u["id"]: u for u in response["members"] if "email" in u["profile"]}

    # Map email and real_name for quick lookup
    email_to_user_info = {
        u["profile"]["email"].lower(): {
            "real_name": u["profile"]["real_name"],
            "user_id": u["id"],
        }
        for u in users_data.values()
    }

    # Fetch messages from the channel
    messages_result = client.conversations_history(channel=channel_id)
    messages = messages_result["messages"]

    # Pre-process messages to identify introductions
    introductions = {
        message["user"]: message["text"]
        for message in messages
        if "introduction" in message.get("text", "").lower()
    }

    # Iterate through DataFrame to update information based on pre-fetched data
    for index, row in df.iterrows():
        user_info = email_to_user_info.get(row["email"].lower(), {})
        slack_accountname = user_info.get("real_name", "")
        user_id = user_info.get("user_id", "")

        # Check for introductions
        did_intro = "yes" if user_id in introductions else "no"
        df.at[index, "did_intro_Y-N"] = did_intro

        # Account name correctness and update Slack account name
        df.at[index, "correct_accountname_Y-N"] = (
            "yes"
            if row["Accountname"].strip().lower() == slack_accountname.lower()
            else "no"
        )
        df.at[index, "Accountname_slack"] = (
            slack_accountname  # Populate with actual Slack username
        )

    # Save the DataFrame back to the Excel file
    df.to_excel(excel_path, index=False)

except SlackApiError as e:
    print(f"Error: {e}")
