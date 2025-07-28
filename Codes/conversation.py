"""
Script to merge and convert multiple JSON files containing chatbot conversation data
into a single, readable PDF file for qualitative analysis.

Author: Pejman Ebrahimi
Date: 2025-07-23
"""
import json
import glob
import os
from datetime import datetime
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch

# Folder with JSON files
json_folder = "Json_Data"
json_files = glob.glob(os.path.join(json_folder, "*.json"))

# Load and merge all conversation data
all_data = []
for file in json_files:
    with open(file, 'r', encoding='utf-8') as f:
        data = json.load(f)
        all_data.extend(data)

# Create PDF document
pdf_filename = "merged_conversations_reportlab.pdf"
doc = SimpleDocTemplate(pdf_filename, pagesize=LETTER,
                        rightMargin=72, leftMargin=72,
                        topMargin=72, bottomMargin=72)

styles = getSampleStyleSheet()
normal_style = styles["Normal"]
heading_style = styles["Heading2"]

elements = []

for session in all_data:
    session_id = session.get("sessionId", "Unknown Session")
    elements.append(Paragraph(f"=== Session: {session_id} ===", heading_style))
    elements.append(Spacer(1, 0.1*inch))
    
    for msg in session["messages"]:
        role = msg.get("role", "unknown").capitalize()
        # Parse timestamp, fallback to raw string if error
        try:
            timestamp = datetime.fromisoformat(msg["time"].replace("Z", "+00:00")).strftime("%Y-%m-%d %H:%M")
        except Exception:
            timestamp = msg["time"]
        content = msg.get("content", "")
        
        # Format message text
        text = f"<b>[{role}] {timestamp}</b><br/>{content}"
        elements.append(Paragraph(text, normal_style))
        elements.append(Spacer(1, 0.1*inch))
    
    elements.append(Spacer(1, 0.3*inch))  # Space between sessions

# Build the PDF
doc.build(elements)

print(f"PDF generated successfully: {pdf_filename}")

