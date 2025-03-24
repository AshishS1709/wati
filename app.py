import pandas as pd
import spacy
from datetime import datetime

# Load spaCy NLP model
nlp = spacy.load("en_core_web_sm")

# Load Excel file
file_path = "AI_Chatbot_Shift_Requests_Sample_Updated.xlsx"
df = pd.read_excel(file_path)

# Convert date column to standard format
df["Requested Date"] = pd.to_datetime(df["Requested Date"], errors="coerce").dt.date

# Store user session data
user_name = None

# Function to extract shift details
def get_shift_details(worker_name, date=None):
    worker_name = worker_name.lower()
    shift_data = df[df["Worker Name"].str.lower() == worker_name]

    if date:
        shift_data = shift_data[shift_data["Requested Date"] == date]

    if not shift_data.empty:
        return shift_data.iloc[0].to_dict()
    return None

# NLP function to extract name and date from user input
def extract_details(user_message):
    doc = nlp(user_message)
    name, date = None, None

    # Handle "My name is ..." patterns manually
    words = user_message.lower().split()
    if "my" in words and "name" in words and "is" in words:
        name_index = words.index("is") + 1
        if name_index < len(words):
            name = " ".join(words[name_index:]).title()  # Convert to title case for names

    # Try extracting from NLP entities
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            name = ent.text
        elif ent.label_ in ["DATE", "TIME"]:
            try:
                parsed_date = pd.to_datetime(ent.text, errors="coerce")
                if not pd.isna(parsed_date):
                    date = parsed_date.date()
            except ValueError:
                date = None

    return name, date

# Function to determine chatbot response
def process_message(user_message):
    global user_name

    # Extract potential name and date
    extracted_name, date = extract_details(user_message)

    # If user provides a name, validate it
    if extracted_name:
        if extracted_name.lower() in df["Worker Name"].str.lower().values:
            user_name = extracted_name
            return f"Got it, {user_name}! How can I assist you with your shift today?"
        else:
            return f"Sorry, {extracted_name} is not found in our records. Please check your name and try again."

    # Ensure a name is provided before answering queries
    if not user_name:
        return "Please provide your name first."

    # Normalize input message for better processing
    user_message = user_message.lower()

    # Handle different queries
    if any(keyword in user_message for keyword in ["schedule", "shift", "work today"]):
        shift = get_shift_details(user_name, date)
        if shift:
            return f"Your next shift is on {shift['Requested Date']} from {shift['Requested Time']} at {shift['Outlet']}."
        return "No shifts found."

    elif any(keyword in user_message for keyword in ["how many hours", "shift length"]):
        shift = get_shift_details(user_name)
        if shift:
            start, end = shift["Requested Time"].split("-")
            start_time = datetime.strptime(start.strip(), "%H:%M")
            end_time = datetime.strptime(end.strip(), "%H:%M")
            hours_worked = (end_time - start_time).seconds // 3600
            return f"Your shift is {hours_worked} hours long, from {start.strip()} to {end.strip()}."
        return "No shift details available."

    elif "swap" in user_message or "change shift" in user_message:
        return "Please provide the date & time. I'll notify your manager."

    elif "late" in user_message or "traffic" in user_message:
        return "How long will the delay be? I'll inform your manager."

    elif "who is my manager" in user_message:
        shift = get_shift_details(user_name)
        if shift and shift["Manager Response"]:
            return f"Your manager's response: {shift['Manager Response']}."
        return "Manager response not found."

    elif "final decision" in user_message:
        shift = get_shift_details(user_name)
        if shift and shift["Final Decision"]:
            return f"The final decision on your request is: {shift['Final Decision']}."
        return "No final decision available."

    elif "cancel" in user_message:
        return "Cancellations require manager approval. I'll notify them now."

    elif "approve" in user_message and "shift" in user_message:
        return "Approved. Workers have been notified."

    elif "reject" in user_message:
        return "Rejected. Worker has been informed."

    elif "overtime" in user_message:
        return "Please confirm the worker and duration of overtime."

    elif "pay rate" in user_message:
        return "Your pay rate is [Rate/Hour]. For more details, contact HR."

    elif "availability" in user_message:
        return "Please provide the dates and times you're available. I'll update the system."

    elif "more workers" in user_message:
        return "How many workers do you need and for which date/time? I'll process your request."

    elif "booking status" in user_message:
        return "Your booking is confirmed. Let me know if you need any changes."

    elif "emergency" in user_message:
        return "Please stay safe. I'll notify your manager immediately."

    elif "feedback" in user_message or "report problem" in user_message:
        return "Please describe the issue. I'll escalate it to the relevant team."

    return "I didn't understand. Ask about shifts, swaps, lateness updates, or bookings."

# Run chatbot in console
print("Shift Bot: Hello! Ask me about your schedule, shift swaps, or lateness updates. Type 'exit' to stop.")

while True:
    # Ask for name until a valid one is provided
    if not user_name:
        user_input = input("You: ")
        
        if user_input.lower() == "exit":
            print("Shift Bot: Goodbye! Have a great day!")
            break

        response = process_message(user_input)
        print(f"Shift Bot: {response}")

        # If user name was not found, continue asking for name
        if "not found" in response:
            continue
        else:
            continue  # Wait for next user input (skip the follow-up question for now)

    # Main query loop
    while True:
        user_input = input("You: ")
        
        if user_input.lower() == "exit":
            print("Shift Bot: Goodbye! Have a great day!")
            break

        response = process_message(user_input)
        print(f"Shift Bot: {response}")

        # Ask if the user needs anything else **only after answering a valid query**
        if "I didn't understand" not in response:
            follow_up = input("Shift Bot: Can I help you with anything else? (yes/no)\nYou: ").strip().lower()

            if follow_up not in ["yes", "y"]:
                print("Shift Bot: Alright! Have a great day!")
                break
