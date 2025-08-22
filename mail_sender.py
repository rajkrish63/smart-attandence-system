from flask import Flask
from flask_mail import Mail, Message
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime
import os

app = Flask(__name__)

# Email Configuration
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'your_email@gmail.com'  # your email
app.config['MAIL_PASSWORD'] = 'your_app_password'     # use app password if using Gmail

mail = Mail(app)

# Function to send email
def send_scheduled_email():
    with app.app_context():
        msg = Message('Scheduled Report',
                      sender=app.config['MAIL_USERNAME'],
                      recipients=['receiver_email@example.com'])
        msg.body = "This is an automated email sent at a scheduled time."
        mail.send(msg)
        print(f"Email sent at {datetime.now()}")

# Setup the scheduler
scheduler = BackgroundScheduler()
scheduler.add_job(func=send_scheduled_email, trigger="cron", hour=9, minute=0)  # Sends at 9:00 AM daily
scheduler.start()

@app.route('/')
def home():
    return "Flask app running with scheduled email!"

if __name__ == '__main__':
    app.run(debug=True)
