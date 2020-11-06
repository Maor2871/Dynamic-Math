import os
from sys import exit
from random import choice, shuffle
from docx import Document
from docx.shared import Pt
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email import encoders
import smtplib


##### ---------- GUI ---------- #####

pass

##### ---------- LUI ---------- #####

def abort():
	"""
		The function aborts the script.
	"""
	
	exit()

def user_input(operations_dict):
	"""
		The function receives and validates all the inputs that the user enters.
	"""
	
	# Ask for a minimum number.
	range_min = input("Please enter a minimum number: ")
	
	# Make sure that the received input is a number.
	if not range_min.isdigit():
		print("Next time enter a number please.")
		abort()

	# Ask for a maximum number.
	range_max = input("\nPlease enter a maximum number: ")
	
	# Make sure that the received input is a number.
	if not range_max.isdigit():
		print("Next time enter a number please.")
		abort()
	
	# Convert range min and range max to int numbers.
	range_min = int(range_min)
	range_max = int(range_max)
	
	# Ask for exercises count.
	amount = input("\nPlease indicate the number of exercises you wish to generate: ")
	
	# Make sure that the received input is a number > 0.
	if not amount.isdigit() or amount.isdigit() and not int(amount) > 0:
		print("Next time enter a number greater than zero please.")
		abort()
		
	# Convert amount to int number.
	amount = int(amount)
	
	# Ask for operations type.
	operations = input("\nPlease indicate what operations you would like to have in the exercises.\n\nm-multiplication\nd-division.\n\n- You can specify more than one operation, for example: dm (both ':' and 'x').\n\nOperations: ")
	
	# Make sure all the received operations are available.
	for chr in operations:	
		if chr not in operations_dict:
			print("The operation: " + chr + " is not available.\nPlease read again the instructions.")
			abort()
	
	# Ask for the document name.
	file_name = input("\nworksheet name: ")
	
	# If the file name has a dot cut the name until the dot.
	if '.' in file_name: file_name = file_name[:file_name.index('.')]

	# The path of the new word document.
	file_path = os.path.join("Exercises", file_name + '.docx')
	
	# If the file name already exist, replace it with the new one.
	if os.path.exists(file_path):
		os.remove(file_path)
	
	# Ask for the desired output: p for print or m for mail.
	action = input("\nPlease indicate in what form you want to get the file.\n\np- direct print\nm- Send document to mail.\n\n- You can specify more than one action, for example: mp (both printing and mailing).\n\nActions: ")
	
	# Make sure an available action received.
	if not 'p' in action and not 'm' in action:
		print("You must enter either 'p' or 'm' to indicate if you want the file to be sent to the printer or to your mail address.")
		abort()
	
	# If the user wants the file to be sent by mail, ask him for his mail address.
	if 'm' in action:
		
		# Get the mail address of the user.
		mail_address = input("\nPlease enter your mail address: ")
		
		# Make a minimal check that a valid mail address was entered.
		if not '@' in mail_address:
			print("A mail address must have '@' in it.")
			abort()
	
	# If not set mail address to None.
	else:
	
		mail_address = None
	
	# Indicate the user that the all the input was received.
	print('\n\nAll input received successfully.\n')
	
	# Return all the inputs.
	return range_min, range_max, amount, operations, file_name, file_path, action, mail_address

	
##### ---------- Functions ---------- #####
	
	
def generate_riddles(amount, operation, range_min, range_max):
	"""
		The function receives operation type and amount of riddles to create and returns a list of operations.
	"""
	
	# Will contain all the exercises.
	riddles = []
	
	# Iterate over the amount of exercises expected.
	for i in range(amount):
	
		riddles.append(generate_riddle(operation, range_min, range_max))
	
	return riddles

def generate_riddle(operation, range_min, range_max):
	"""
		The function receives operation type and returns a riddle of that type.
	"""
	
	# Generate the numbers.
	num1 = choice(range(range_min, range_max))
	num2 = choice(range(range_min + 1, range_max))
	
	# If else operation type. Return the riddle of the received operation.
	if operation == "m":
	
		return str(num1) + ' x ' + str(num2) + '  = '
		
	elif operation == "d":
	
		return str(num1 * num2) + ' : ' + str(num1) + '  = '
		

def create_exercises(operations, amount, range_min, range_max):
	"""
		The function creates and randomizes the desired exercises.
	"""
	
	# The list with all the exercises.
	riddles = []
	
	# Check if more than one operation selected.
	if len(operations) > 1:
		amount = int(round(amount / len(operations)))

	# Check if the operation of the exercises is division.
	if "d" in operations:
		
		riddles += generate_riddles(amount, 'd', range_min, range_max)

	# Create multiplication exercises as default.
	if 'm' in operations:
		
		riddles += generate_riddles(amount, 'm', range_min, range_max)
	
	# Shuffle all the exercises.
	shuffle(riddles)
	
	# Return all the exercises as a list.
	return riddles


def format_riddles(riddles):
	"""
		The function receives exercises list and formats them to a string.
	"""
	
	# Format the riddles.
	riddles_formatted = '\n\n\n'.join(riddles[i] + '\t\t\t\t\t\t\t' + riddles[i - 1] for i in range(1, len(riddles), 2))

	# If odd number of exercises add the last one.
	if not len(riddles_formatted) %2 == 0: riddles_formatted += '\n\n\n' + riddles[-1]
	
	# Return the exercises string.
	return riddles_formatted
	

def save_as_word(operations_dict, operations, range_min, range_max, riddles_formatted, file_path):
	"""
		The function saves the exercises as a word document.
	"""
	
	# Create the word document.
	document = Document()

	# Add header to the document.
	document.add_heading((', '.join(operations_dict[operation] for operation in operations)).replace(', ', " and ", -1) + ' of numbers from ' + str(range_min) + ' to ' + str(range_max) + '\n\n', 0)

	# Define the font size.
	style = document.styles['Normal']
	font = style.font
	font.size = Pt(14)

	# Add the exercises to the document.
	document.add_paragraph(riddles_formatted, style='Normal')

	# Save the document.
	document.save(file_path)
	
	
def send_mail(mail_address, file_name, file_path):
	"""
		The function sends the document by mail.
	"""
	
	print("Preparing mail to be sent... ", end="")
	
	fromaddr = 'classifiedanonymouse@gmail.com'
	toaddr = mail_address


	msg = MIMEMultipart()
	msg['From'] = fromaddr
	msg['To'] = toaddr
	msg['Subject'] = file_name
	msg.attach(MIMEText('Dynamic math worksheet created ' + file_name + " for you.\nIt is attached to the mail as a word document."))


	attachment = open(file_path, 'rb')
	part = MIMEBase("application", "octet-stream")
	part.set_payload(attachment.read())
	encoders.encode_base64(part)
	part.add_header("Content-Disposition", "attachment; filename= " + file_name + ".docx")
	msg.attach(part)
	msg = msg.as_string()

	try:
		server = smtplib.SMTP('smtp.gmail.com:587')
		server.ehlo()
		server.starttls()
		server.login(fromaddr, 'ThisIsClassified4@')
		# server.sendmail(fromaddr, toaddr, msg)
		server.quit()
		
		print('Email sent successfully.')

	except Exception as e:
		print("Email couldn't be sent.")
		
		
def execute_actions(action, mail_address, file_name, file_path):
	"""
		receive the word document and choose what to do with it.
	"""
	
	# If the user want the document by mail send it to his mail address.
	if 'm' in action:

		# Send the document to mail.
		send_mail(mail_address, file_name, file_path)

	# If the user want the document to be printed, send it to the printer.
	if 'p' in action:
		
		print("Preparing document to print... ", end="")
		
		try:
		
			# Print the word document.
			#os.startfile(os.path.join("Exercises", file_name + '.docx'), "print")
		
			print("Document Sent to printer successfully.")
			
		except Exception as e:
		
			print("Document couldn't be sent to printer.")
	

##### ---------- Create exercises ---------- #####


# Indicates what operations are available and their short-cuts.
operations_dict = {'m': "Multiplication", 'd': "Division"}

# Get all the necessary inputs from the user.
range_min, range_max, amount, operations, file_name, file_path, action, mail_address = user_input(operations_dict)

print("Creating exercises... ", end="")

# Will contain all the exercises.
riddles = create_exercises(operations, amount, range_min, range_max)

print("complete.\n")
print("formatting exercises... ", end="")

# Format all the exercises to a string.
riddles_formatted = format_riddles(riddles)

print("complete.\n")
print("\nSaving exercises to a word document... ", end="")

# Save the exercises as a word document.
save_as_word(operations_dict, operations, range_min, range_max, riddles_formatted, file_path)

print("complete.\n")
print("Executing action:")

# Execute the actions that the user wishes to do with the document.
execute_actions(action, mail_address, file_name, file_path)

print("\nAll actions are done successfully, thank you for using Dynamic Math.\nCome again soon.")