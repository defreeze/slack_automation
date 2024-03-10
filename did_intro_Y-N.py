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
    # Fetch messages from the channel
    messages_result = client.conversations_history(channel=channel_id)
    messages = messages_result["messages"]

    # Iterate through each row in the DataFrame to check introduction status
    for index, row in df.iterrows():
        # Check if 'did_intro_Y-N' is not filled (NaN, empty, or some placeholder like 'TBD')
        if pd.isna(row["did_intro_Y-N"]) or row["did_intro_Y-N"].strip().lower() in [
            "",
            "tbd",
        ]:
            excel_email = row["email"].lower()
            account_name = row["Accountname"].lower()
            did_intro = "no"  # Default assumption

            # Iterate through messages to check for introductions by this user
            for message in messages:
                if "introduction" in message.get("text", "").lower():
                    user_id = message.get("user")
                    # Fetch the user's profile to get their email and real name
                    user_info = client.users_info(user=user_id)
                    user_email = user_info["user"]["profile"].get("email", "").lower()
                    user_name = (
                        user_info["user"]["profile"].get("real_name", "").lower()
                    )

                    if user_email == excel_email or user_name == account_name:
                        did_intro = "yes"
                        break  # Introduction found, no need to check further messages

            # Update the 'did_intro_Y-N' column in the DataFrame
            df.at[index, "did_intro_Y-N"] = did_intro

    # Save the DataFrame back to the Excel file
    df.to_excel(excel_path, index=False)

except SlackApiError as e:
    print(f"Error: {e}")
