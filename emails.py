import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import sensitive_info

def send_emails(data_by_person):
    sender_email = sensitive_info.test_sender_email
    password = sensitive_info.test_sender_password

    for student in data_by_person.keys():
        receiver_email = sensitive_info.test_receiver_email
        subject = "Your AP Exam Schedule - " + student

        # Convert dictionary to JSON string
       # Convert data to HTML table
        body = f"""
        <html>
        <body>
            <h2>AP Exam Schedule for {student}</h2>
            <table>
                <tr>
                    <th>Exam</th>
                    <th>Date</th>
                    <th>Number</th>
                    <th>Room Number</th>
                    <th>Proctor</th>
                </tr>
        """
        for exam in data_by_person.get(student):
            body += f"""
                <tr>
                    <td>{exam['exam']}</td>
                    <td>{exam['date'].strftime('%Y-%m-%d %H:%M:%S')}</td>
                    <td>{exam['number']}</td>
                    <td>{exam['room_number']}</td>
                    <td>{exam['proctor']}</td>
                </tr>
            """
        body += """
            </table>
            <style>
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                th, td {
                    border: 1px solid #dddddd;
                    text-align: left;
                    padding: 8px;
                }
                tr:nth-child(even) {
                    background-color: #f2f2f2;
                }
                th {
                    background-color: #4CAF50;
                    color: white;
                }
            </style>
        </body>
        </html>
        """

        # Create the email
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message.attach(MIMEText(body, "html"))

        
        smtp_server = "smtp.gmail.com"
        port = 587
        
        try:
            # Connect to the SMTP server
            server = smtplib.SMTP(smtp_server, port)
            server.starttls()  # Secure the connection
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message.as_string())
            print("Email sent successfully!")
        except Exception as e:
            print(f"Error: {e}")
        finally:
            server.quit()