from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
import os
from docx import Document
from docx.shared import Inches
import google.generativeai as genai
from backend import generate_recommendations_english, create_report_english, generate_recommendations_arabic,create_report_arabic
from database import Recommendation_English,Recommendation_Arabic
import logging

# Initialize Flask app
app = Flask(__name__, static_url_path='/static')

# Configure Google Generative AI API key
genai.configure(api_key=os.environ['API_KEY'])

# Route for language selection page
@app.route('/')
def language_selection():
    """Renders the language selection page."""
    return render_template('language_selection.html')

# Route for the main application page
@app.route('/index/<language>')
def index(language):
    """
    Renders the main application page.

    Args:
        language (str): The selected language ('arabic' or 'english').
    
    Returns:
        Rendered HTML template for the index page.
    """
    # Load recommendations based on the selected language
    if language == 'arabic':
        Recommendations = list(Recommendation_Arabic.keys())
    else:
        Recommendations = list(Recommendation_English.keys())

    return render_template('index.html', language=language, Recommendation=Recommendations)

# Route for generating the energy audit report
@app.route('/generate_report', methods=['POST'])
def generate_report():
    """
    Generates and sends the energy audit report.

    Receives form data, generates recommendations, creates the report,
    and sends it as a downloadable file.
    """
    # Get form data as a dictionary
    data = request.form.to_dict()
    # Extract language information
    language = data.pop('language')

    # Generate recommendations and create report based on language
    if language == 'arabic':
        recommendations = generate_recommendations_arabic(data)
        create_report_arabic(data, recommendations)
        filename = f'Manzili_Energy_Audit_Report_{data["رقم_التقرير"]}.docx'
    else:
        recommendations = generate_recommendations_english(data)
        create_report_english(data, recommendations)
        filename = f'Manzili_Energy_Audit_Report_{data["report_number"]}.docx'

    # Send the generated report file
    return send_file(filename, as_attachment=True, download_name=filename)

# Run the Flask application if the script is executed directly
if __name__ == '__main__':
    app.run(debug=True)