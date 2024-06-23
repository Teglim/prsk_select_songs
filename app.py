from flask import Flask, request, render_template
import pandas as pd
import random

app = Flask(__name__)

# Excelファイルを読み込む
df = pd.read_excel("prsk.xlsx")

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/filter', methods=['POST'])
def filter_data():
    writing = request.form['writing']
    unit = request.form['unit']
    expert = request.form['expert']
    master = request.form['master']
    range_input = request.form['range']

    filtered_df = df.copy()

    if writing == "T":
        filtered_df = filtered_df[filtered_df.iloc[:, 1].astype(str).str.startswith('2')]
    elif writing != "p":
        filtered_df = filtered_df[filtered_df.iloc[:, 1].astype(str).str.startswith('1')]

    if unit != "p":
        filtered_df = filtered_df[filtered_df.iloc[:, 1].astype(str).str[1] == unit]

    if expert != "p":
        filtered_df = filtered_df[filtered_df.iloc[:, 8] == int(expert)]

    if master != "p":
        filtered_df = filtered_df[filtered_df.iloc[:, 9] == int(master)]

    if range_input != "p":
        range_start, range_end = map(int, range_input.split('-'))
        filtered_df = filtered_df[
            (filtered_df.iloc[:, 8].between(range_start, range_end)) | 
            (filtered_df.iloc[:, 9].between(range_start, range_end))
        ]

    if not filtered_df.empty:
        random_row = filtered_df.sample(n=1)
        output = random_row.iloc[:, [3, 4, 8, 9, 11, 12, 14]].values.flatten()
        output_columns = ['name', 'unit', 'expert', 'master', 'expert_notes', 'master_notes', 'length']
        output_dict = dict(zip(output_columns, output))
        return render_template('index.html', result=output_dict)
    else:
        return render_template('index.html', result="Cannot find.")

if __name__ == '__main__':
    app.run(debug=True)
