from flask import Flask, render_template,request,redirect,url_for,send_from_directory,g
import pandas as pd
import smtplib 
app = Flask(__name__)


@app.route('/')
def home():
	return render_template('index.html')

@app.route('/upload',methods=['GET','POST'])
def upload():
	try :
		if request.method == "POST":
			
			# Getting All Requested Data
			sender_email =  request.form['sender_mail']
			sender_password =  request.form['sender_pass']
			sender_msg = request.form['msg']
			sender_password = str(sender_password)
			FILE = request.files['myFile']

			# Extracting Data
			data = pd.read_excel(FILE)
			list_of_email = data['email'].tolist()
			print(sender_password,sender_email,list_of_email)

			# SMTP Server Setup
			server = smtplib.SMTP('smtp.gmail.com',587)
			server.starttls()
			server.login(sender_email,sender_password)
			print("Login ok")

			#Sending Emails
			for email in range(0,len(list_of_email)):
				server.sendmail(sender_email,list_of_email[email],str(sender_msg))
			print('Email sended')

			# Rendering
			return render_template('success.html',data=list_of_email)
	except:
		return render_template('error.html')


if __name__ == '__main__':
	app.run(debug=True)