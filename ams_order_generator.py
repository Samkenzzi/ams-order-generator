import pandas as pd
from fuzzywuzzy import process
from flask import Flask, request, send_file, render_template_string
import tempfile
import os

app = Flask(__name__)

# Load static reference files
UPC_FILE = "UPC CODES.xlsx"
DELIVERY_FILE = "Delivery addresses v2.xlsx"

def load_reference_data():
    upc_df = pd.read_excel(UPC_FILE)
    upc_df.columns = ["Model Number", "Item_Code", "Item Description"]

    delivery_df = pd.read_excel(DELIVERY_FILE)
    delivery_df.rename(columns={"Store": "Mapped Store"}, inplace=True)
    delivery_df["Mapped Store"] = delivery_df["Mapped Store"].str.strip().str.title()

    return upc_df, delivery_df

def fuzzy_map_stores(order_df, delivery_df):
    store_names = delivery_df["Mapped Store"].unique()
    order_df["Location"] = order_df["Location"].str.strip().str.title()
    order_df["Mapped Store"] = order_df["Location"].apply(
        lambda loc: process.extractOne(loc, store_names)[0] if pd.notnull(loc) else None
    )
    return order_df

def generate_ams_file(order_file):
    upc_df, delivery_df = load_reference_data()
    order_df = pd.read_excel(order_file)

    order_df = fuzzy_map_stores(order_df, delivery_df)
    merged_df = pd.merge(order_df, upc_df, on="Model Number", how="left")
    full_df = pd.merge(merged_df, delivery_df, on="Mapped Store", how="left")

    full_df["Sales Order Number"] = full_df["Document Number"]
    full_df["Date of Order"] = full_df["Date"]

    final_df = full_df[[
        "Phone", "Sales Order Number", "Date of Order", "Item_Code", "Item Description", "Quantity",
        "Ship_Addressee", "Ship_Address Line 1", "Ship_Address Line 2",
        "Ship_City", "Ship_State", "Ship_Postcode", "Ship_Country",
        "Ship_Delivery Instructions", "Despatch_Method"
    ]]

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    final_df.to_excel(temp_file.name, index=False)
    return temp_file.name

UPLOAD_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>AMS Order File Generator</title>
</head>
<body>
    <h2>Upload Shaver Shop Order File</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx" required>
        <input type="submit" value="Generate AMS File">
    </form>
</body>
</html>
"""

@app.route("/", methods=["GET"])
def index():
    return render_template_string(UPLOAD_HTML)

@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    output_path = generate_ams_file(file)
    return send_file(output_path, as_attachment=True, download_name="AMS_Order.xlsx")

if __name__ == "__main__":
    app.run(debug=True)
