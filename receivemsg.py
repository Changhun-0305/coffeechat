from flask import Flask, request, redirect
from twilio.twiml.messaging_response import MessagingResponse
from retrieve_from_excel import user_dict, loc
import openpyxl

app = Flask(__name__)

@app.route("/sms", methods=['GET', 'POST'])
def sms_reply():
    """Respond to incoming calls with a simple text message."""
    # Start our TwiML response
    resp = MessagingResponse()
    number = request.form['From']
    message_body = request.form['Body']
    write_response(number, message_body)
    # Add a message
    #resp.message("Hello {}, you said: {}".format(number, message_body))

    return str(resp)


def write_response(number, answer):
    """Writes responses to excel"""
    for i, (key, item) in enumerate(user_dict.items()):
        if item[3] == number:
            print(key, number, answer)
            cell = sheet.cell(i+1, 6)
            cell.value = answer
            book.save(loc)

if __name__ == "__main__":
    book = openpyxl.load_workbook(loc)
    sheet = book.active
    app.run(debug=True)

