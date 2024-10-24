import random
import json

from datetime import datetime
import torch
from model import NeuralNet
from nltk_utils import bag_of_words, tokenize
import sqlite3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')

with open('intents.json', 'r', encoding='utf-8') as json_data:
    intents = json.load(json_data)

FILE = "data.pth"
data = torch.load(FILE, weights_only=True)

input_size = data["input_size"]
hidden_size = data["hidden_size"]
output_size = data["output_size"]
all_words = data['all_words']
tags = data['tags']
model_state = data["model_state"]

model = NeuralNet(input_size, hidden_size, output_size).to(device)
model.load_state_dict(model_state)
model.eval()


bot_name = "Qui Meo"


previous_user_message = None
awaiting_confirmation = False

def get_response(msg):
    global previous_user_message, awaiting_confirmation

    if awaiting_confirmation:
        user_response = msg.lower()

        if user_response == "yes":
            if previous_user_message:
                send_question_and_save("bichqui1212@gmail.com","2100007862", previous_user_message, "nguyenddqui@gmail.com")
                awaiting_confirmation = False
                return "I've notified the manager about your question."
            else:
                awaiting_confirmation = False
                return "There was an issue retrieving your original question."
        elif user_response == "no":
            awaiting_confirmation = False
            previous_user_message = None
            return "Okay, let me know if you have any other questions."
        else:
            return "Please respond with 'yes' or 'no'."

    else:
        sentence = tokenize(msg)
        X = bag_of_words(sentence, all_words)
        X = X.reshape(1, X.shape[0])
        X = torch.from_numpy(X).to(device)

        output = model(X)
        _, predicted = torch.max(output, dim=1)

        tag = tags[predicted.item()]
        probs = torch.softmax(output, dim=1)
        prob = probs[0][predicted.item()]

        if prob.item() > 0.75:
            for intent in intents['intents']:
                if tag == intent["tag"]:
                    previous_user_message = None
                    awaiting_confirmation = False
                    return f"{random.choice(intent['responses'])}"
        else:
            previous_user_message = msg
            awaiting_confirmation = True
            return "I'm sorry, I couldn't find an answer to your question. Would you like me to notify the manager? (yes/no)"

def send_email(receiver_email, subject, body):
    sender_email = "bichqui1212@gmail.com"
    password = "lmbp reme uxeo gumj"

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    server = smtplib.SMTP("smtp.gmail.com", 587)
    try:

        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        print(f"Email sent to {receiver_email}")
    except Exception as e:
        print(f"Failed to send email to {receiver_email}. Error: {e}")
    finally:
        server.quit()

def insert_question(student_id, question_text, email):
    conn = sqlite3.connect('students.db')
    cursor = conn.cursor()

    submission_datetime = datetime.now()
    submission_datetime = datetime(submission_datetime.year,
                                                   submission_datetime.month,
                                                   submission_datetime.day,
                                                   submission_datetime.hour,
                                                   submission_datetime.minute)

    cursor.execute('''
        INSERT INTO questions (StudentID, QuestionText, SubmissionDateTime, Status, Email)
        VALUES (?, ?, ?, ?, ?)
    ''', (student_id, question_text, submission_datetime, 'Pending', email))

    question_id = cursor.lastrowid

    conn.commit()

    cursor.close()
    conn.close()

    return question_id


def send_question_and_save(from_email, student_id,question_text, receiver_email):

    question_id = insert_question(student_id, question_text, from_email)

    subject = f"New question: {question_id} - {student_id}"
    body = f"Student ID: {student_id}\nQuestion: {question_text}"

    send_email(receiver_email, subject, body)


