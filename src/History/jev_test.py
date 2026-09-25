from dotenv import load_dotenv
from typesafe_sdk import TypeSafeClient, Choice

load_dotenv()
client = TypeSafeClient()

response = client.system_one(
    state="Office: Paris. Amount written: '1,500'. Message: Please send me the updated invoice for the group booking. The customer is American",
    questions={
        "amount_format": Choice(
            instructions="How should the amount '1,500' be read?",
            criteria={
                "english": "One thousand five hundred (comma is a thousands separator)",
                "french": "One point five (comma is a decimal separator)",
            },
        ),
    },
)

answer = response.answers["amount_format"]
print("choice:", answer.choice)
print("probabilities:", answer.probabilities)
print("confidence:", answer.confidence)