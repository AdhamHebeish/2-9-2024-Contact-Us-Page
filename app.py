from flask import Flask, render_template, request
from openpyxl import load_workbook

app = Flask(__name__)
wsgi_app = app.wsgi_app

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/excelify',methods=['POST','GET'])
def toexcel():
    err = ""
    mail = request.form['email']
    mail = mail.strip()
    tel = request.form['phone']
    tel = tel.strip()
    note = request.form['message']
    note = note.strip()
    goodmail = mail.endswith("@gmail.com") or mail.endswith("@yahoo.com") or mail.endswith("@microsoft.com") or mail.endswith("@outlook.com")
    if mail == "" or tel == "" or note == "":
        err = "You have not filled in all the fields"
    elif goodmail == False:
        err = "This e-mail is not from the lsit of accepted e-mails."
    elif tel.startswith("+") == False:
        err = "Please remember to include the country code in the phone number."
    else:
        contactusexcel = 'Contact Us Excel.xlsx'
        info = [[mail, tel, note]]
        xlsx = load_workbook(contactusexcel)
        sheet = xlsx.active
        for x in info:
            sheet.append(x)
        xlsx.save(contactusexcel)
        err = "Your info has been successfully uploaded to our server."
    return render_template('index.html',errcode=err)

if __name__ == '__main__':
    app.run()
