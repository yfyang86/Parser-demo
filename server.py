from flask import Flask, request, jsonify
import json
from sophon_doc_parser import SophonDocParser

app = Flask(__name__)
parser = SophonDocParser('configure.json')

@app.route('/vlm_pdf__parser', methods=['POST'])
def vlm_pdf_parser():
    data = request.json
    if 'file' not in data:
        return jsonify({"error": "No file part in the request"}), 400

    input_base64 = data['file']

    try:
        output_json = parser.parse_pdf(
            input=input_base64,
            output_dir='./',
            prompt=None,
            api_key=None,
            base_url=None,
            model='gpt-4o',
            verbose=False,
            gpt_worker=1
        )
        json_data = json.loads(output_json)
        return jsonify(json_data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/parser_unified', methods=['POST'])
def unified_parser():
    data = request.json
    if 'file' not in data:
        return jsonify({"error": "No file part in the request"}), 400

    if 'type' not in data:
        return jsonify({"error": "No type specified"}), 400

    input_base64 = data['file']
    file_type = data['type']

    try:
        if file_type == 'pdf':
            output_json = parser.parse_pdf(
                input=input_base64,
                output_dir='./',
                prompt=None,
                api_key=None,
                base_url=None,
                model='gpt-4o',
                verbose=False,
                gpt_worker=1
            )
        else:
            output_json = parser.parser_unified(input_base64, file_type)

        json_data = json.loads(output_json)
        return jsonify(json_data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    with open('configure.json', 'r') as f:
        config = json.load(f)
    app.run(debug=True, port=config['flask_port'], host = config['flask_host'])