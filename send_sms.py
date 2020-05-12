from twilio.rest import Client
from credentials import account_sid, auth_token, my_cell, my_twilio
from retrieve_from_excel import user_dict

client = Client(account_sid, auth_token)

my_msg = "Hi. This is KAP coffee chat. Are you free for coffee next week? If you reply \"yes\", we will pair you up with another person tomorrow. Please reply by tonight 12am."

if __name__ == "__main__":
    for i, (key, item) in enumerate(user_dict.items()):
        message = client.messages.create(to=item[3], from_=my_twilio, body=my_msg)
