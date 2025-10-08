import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime
import os

# Input and Output paths
DATA_FILE = "data/invoices_data.xlsx"
OUTPUT_FOLDER = "output/"

# Ensure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Load data
df = pd.read_excel(DATA_FILE)

def generate_invoice(row):
    """Generate a PDF invoice from one row of the dataframe."""
    invoice_id = row["Invoice ID"]
    name = row["Name"]
    email = row["Email"]
    date = pd.to_datetime(row["Date"]).strftime("%d-%b-%Y")
    amount = row["Amount"]
    desc = row["Description"]

    filename = f"Invoice_{invoice_id}_{name.replace(' ', '')}.pdf"
    filepath = os.path.join(OUTPUT_FOLDER, filename)

    c = canvas.Canvas(filepath, pagesize=A4)
    width, height = A4

    # Header
    c.setFont("Helvetica-Bold", 20)
    c.drawString(200, 800, "INVOICE")

    # Metadata
    c.setFont("Helvetica", 12)
    c.drawString(50, 760, f"Invoice ID: {invoice_id}")
    c.drawString(50, 740, f"Date: {date}")

    # Customer Info
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, 700, f"Billed To:")
    c.setFont("Helvetica", 12)
    c.drawString(50, 680, f"{name}")
    c.drawString(50, 665, f"{email}")

    # Description
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, 620, "Description:")
    c.setFont("Helvetica", 12)
    c.drawString(50, 600, desc)

    # Amount
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, 550, f"Amount: ₹{amount:,.2f}")

    # Footer
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, 100, "Thank you for your business!")

    c.save()
    print(f"✅ Generated {filename}")

# Generate invoices for each row
for _, row in df.iterrows():
    generate_invoice(row)
