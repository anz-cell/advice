from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
import os
from docx import Document
from docx.shared import Inches
import google.generativeai as genai
from backend import generate_recommendations_english, create_report_english, generate_recommendations_arabic,create_report_arabic
from database import Recommendation_English,Recommendation_Arabic

app = Flask(__name__, static_url_path='/static')
# Configure Google API
os.environ['API_KEY'] = 'AIzaSyCVVe2FwYmaaDG61RAQ-e8pOvIs8CzsrME'
genai.configure(api_key=os.environ['API_KEY'])

@app.route('/')
def language_selection():
    return render_template('language_selection.html')

@app.route('/index/<language>')
def index(language):
    if language == 'arabic':
        Recommendations = list(Recommendation_Arabic.keys())
    else:
        Recommendations = list(Recommendation_English.keys())
    return render_template('index.html', language=language, Recommendation = Recommendations)

@app.route('/generate_report', methods=['POST'])
def generate_report():
        data = request.form.to_dict()
        language = data.pop('language').strip().lower()
        if language == 'arabic':
            recommendations = generate_recommendations_arabic(data)
            create_report_arabic(data, recommendations)
            filename = f'Manzili_Energy_Audit_Report_{data["رقم_التقرير"]}.docx'
        else:
            recommendations = generate_recommendations_english(data)
            create_report_english(data, recommendations)
            filename = f'Manzili_Energy_Audit_Report_{data["report_number"]}.docx'
        return send_file(filename, as_attachment=True, download_name=filename)

@app.route('/delete_file', methods=['POST'])
def delete_file():
    data = request.json
    filename = data.get('filename')
    if filename:
        filename= os.path.join(os.path.dirname(__file__), filename)
        os.remove(filename)
        #since the function doesnt return it will say there is a warning it doesnt return but we dont want to return anything so its fine, the warning doesnt cause any problems.

if __name__ == '__main__':
    app.run()
