from flask import Flask
from flask import render_template
from flask import request
import pandas as pd
from docx import Document
from docx.shared import Inches
from docxtpl import DocxTemplate
from docx2pdf import convert
import sys
import os
if sys.version_info >= (3, 6):
    import zipfile
else:
    import zipfile36 as zipfile
app = Flask(__name__)


@app.route('/')
def index():
	return render_template('index.html')


@app.route('/result', methods=['POST'])
def result():
	# Read the data.xlsx file and convert it to a dataframe
	data = pd.read_excel('data.xlsx')
	# Read user_input.xlsx file and convert it to a dataframe
	user_input = pd.read_excel('user_input.xlsx')
	print(data)
	print(user_input)
#  Merge the two dataframes on the 'roll' column
	merged_data = pd.merge(data, user_input, on='roll')
	print(merged_data)
	# number of rows in the merged dataframe
	n = merged_data.shape[0]
	# Fill the docx template with the data from the merged dataframe
	for j in range(n):
		context = get_context(merged_data, j)
		doc = DocxTemplate("template.docx")
		doc.render(context)
		#save the docx file in static folder
		doc.save("result"+str(j)+".docx")
		#convert the docx file to pdf
		converter = convert('result'+str(j)+'.docx', 'result/'+str(merged_data['name'][j])+' '+str(merged_data['roll'][j])+'.pdf')
	#zip the pdf files
	zipf = zipfile.ZipFile('result.zip', 'w', zipfile.ZIP_DEFLATED)
	zipdir('result/', zipf)
	zipf.close()
	# Delete the docx files
	for j in range(n):
		os.remove('result'+str(j)+'.docx')
	return render_template('result.html')


def getGrade(marks, max_marks):
	percentage = (marks/max_marks)*100
	if (percentage >= 90):
		return 'A'
	elif (percentage >= 80):
		return 'B'
	elif (percentage >= 70):
		return 'C'
	elif (percentage >= 60):
		return 'D'
	elif (percentage >= 50):
		return 'E'
	else:
		return 'F'


def get_context(df, j):
	context = {}
	context['row_contents'] = []
	sub = 0
	flag = 0
	sno = 0
	marks_obtained = 0
	max_marks = 0
	for i in df.columns:
		if (i == 'subject1' or flag == 1):
			flag = 1
			if (sub % 3 == 0):
				context['row_contents'].append({'sno': sno, 'scode': df[i][j]})
				sno += 1
				sub += 1
			elif (sub % 3 == 1):
				dict = context['row_contents'][sno-1]
				dict['maxmarks'] = df[i][j]
				max_marks += df[i][j]
				context['row_contents'][sno-1] = dict
				sub += 1
			else:
				dict = context['row_contents'][sno-1]
				dict['smarks'] = df[i][j]
				marks_obtained += df[i][j]
				context['row_contents'][sno-1] = dict
				sub += 1
		else:
			if (flag == 0):
				# context[i] = list of values in the column
				context[i] = df[i][j]
	context['max_marks'] = max_marks
	context['marks_obtained'] = marks_obtained
	context['grade'] = getGrade(marks_obtained, max_marks)
	# print(context)
	return context

def zipdir(path, ziph):
	# ziph is zipfile handle
	for root, dirs, files in os.walk(path):
		for file in files:
			ziph.write(os.path.join(root, file))
	

if __name__ == '__main__':
	app.run(debug=True)
