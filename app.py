from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import datetime
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'file1' not in request.files or 'file2' not in request.files:
        return redirect(request.url)

    file1 = request.files['file1']
    file2 = request.files['file2']

    if file1 and file2:
        df = pd.read_excel(file1, header=1)
        person = pd.read_excel(file2, header=1)
        print(df)
        print(person)
        df_final = pd.DataFrame(columns=['Name', 'Record date', 'Time in', 'Time out', 'Late Time'])
        date = datetime.date(1, 1, 1)
        check_in = datetime.time(8, 30, 0)
        check_out = datetime.time(18, 0, 0)

        df_date = df['Record Date'].unique()
        for i in range(len(person['First Name'])):
            for j in range(len(df_date)):
                df_result = df[(df['First Name'] == person['First Name'][i]) & (df['Record Date'] == df_date[j])]
                if df_result.empty:
                    Name = person['First Name'][i]
                    Date = df_date[j]
                    timein = ""
                    timeout = ""
                    status = "no tapping"
                else:
                    time_in = df_result['Earliest Time'].loc[df_result.index[0]].split(":")
                    time_out = df_result['Latest Time'].loc[df_result.index[0]].split(":")
                    person_in = datetime.time(int(time_in[0]), int(time_in[1]), int(time_in[2]))
                    person_out = datetime.time(int(time_out[0]), int(time_out[1]), int(time_out[2]))

                    if person_in <= check_in:
                        status = ""
                    else:
                        datetime1 = datetime.datetime.combine(date, check_in)
                        datetime2 = datetime.datetime.combine(date, person_in)
                        status = str(datetime2 - datetime1)

                    Name = person['First Name'][i]
                    Date = df_result['Record Date'].loc[df_result.index[0]]
                    timein = df_result['Earliest Time'].loc[df_result.index[0]]
                    timeout = df_result['Latest Time'].loc[df_result.index[0]]

                new_row = {'Name': Name, 'Record date': Date, 'Time in': timein, 'Time out': timeout, 'Late Time': status}
                new_df = pd.DataFrame([new_row])
                df_final = pd.concat([df_final, new_df], ignore_index=True)
                print(df_final, "final")

        output_file = os.path.join(app.config['DOWNLOAD_FOLDER'], 'Attendance-Cvox.csv')
        df_final.to_csv(output_file, encoding='utf-8-sig', index=False)

        # Redirect to the download page with a query parameter
        return redirect(url_for('download_file', filename='Attendance-Cvox.csv'))
    print("out")
    return redirect(request.url)

@app.route('/download')
def download_file():
    filename = request.args.get('filename')
    if not filename:
        return redirect(url_for('home'))
    output_file = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
    if os.path.exists(output_file):
        return send_file(output_file, as_attachment=True)
    return redirect(url_for('home'))

if __name__ == '__main__':
    app.run(debug=True)