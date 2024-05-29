from flask import Flask, request, jsonify
from dotenv import load_dotenv
import xlwings as xw
import os

# Load environment variables from .env file
load_dotenv()

# Get the port number from the environment variable
PORT = int(os.getenv("PORT", default=5000))

app = Flask(__name__)

@app.route('/modify_excel', methods=['GET'])
def modify_excel():
    try:
        # Get file_name and vba_script from URL parameters
        file_path = request.args.get('filePath')
        vba_script = request.args.get('scriptName')

        # Open the Excel file
        wb = xw.Book(file_path)

        # Run VBA script
        macro_to_run = wb.macro(vba_script)
        macro_to_run()

        # Save the workbook
        # wb.save()
        # wb.close()

        return jsonify({"status": "success"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=PORT)