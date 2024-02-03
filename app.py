from flask import Flask, render_template, request,send_file
import pandas as pd
import re
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    # Get the uploaded files
    file1 = request.files['file1']
    file2 = request.files['file2']

    # Read data from a specific sheet in the Excel file
    df = pd.read_excel(file1, sheet_name='Sheet1')
  

    # Extract the 8th column (index 7)
    rawMedicine = df.iloc[:, 0]
    rawConsumedStrips = df.iloc[:, 1]
    rawClosingStrips = df.iloc[:, 2]

    consumedStrips = list(rawConsumedStrips)
    closingStrips = list(rawClosingStrips)
    medicine = list(rawMedicine)
    # print(consumedStrips)
    # print(closingStrips)


    # Function to replace matched strings
    def replace_amount(match):
        # Extract the matched amount and convert it to an integer
        num = match.group('num')
        s = match.group(0)
        
        # Replace with a new value (e.g., double the amount)
        new_value = s.replace(num,f" {num}")
        return new_value

    pattern = r'(mf|m|xr|cv|as|mv)(?P<num>\d+)'

    data = {}
    for i,m in enumerate(medicine):
        m = m.strip().lower()
        m = m.replace('tab.','')
        m = m.replace('tab','')
        m = m.replace('-','')
        m = m.strip()
        m = m.strip('.')
        m = m.strip()
        m = re.sub(pattern,replace_amount,m)
        data[m]={
            'consumed': consumedStrips[i],
            'closing': closingStrips[i]
        }


    df2 = pd.read_excel(file2, sheet_name='Sheet1')

    rawStocklistCode = df2.iloc[:, 0]
    saqibConsumed = df2.iloc[:, 3]
    saqibClosing = df2.iloc[:, 4]

    stocklistCode = []
    for s in rawStocklistCode:
        s = s.strip().lower()
        s = s.replace('15 tab.','')
        s = s.replace('15 tab','')
        s = s.replace('tab.','')
        s = s.replace('tab','')
        s = s.strip()
        s = s.strip('.')
        s = s.strip()

        stocklistCode.append(s)


    newConsumed = []
    newClosing = []
    for i,slc in enumerate(stocklistCode):
        if slc in data:
            newConsumed.append(data[slc]['consumed'])
            newClosing.append(data[slc]['closing'])
        else:
            newConsumed.append(saqibConsumed[i])
            newClosing.append(saqibClosing[i])


    print(newConsumed)
    print(newClosing)

    df2['JK1004'] = newConsumed
    df2['JK1004.1'] = newClosing

    # Create BytesIO object to store the Excel content
    excel_output = BytesIO()
    with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
        df2.to_excel(writer, index=False)

    # Set the BytesIO object's position to the beginning
    excel_output.seek(0)

    # Return the Excel file as a downloadable attachment
    return send_file(
        excel_output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='edited_doc.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True)
