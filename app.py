from flask import Flask, request, render_template
import send

app = Flask(__name__)

@app.route('/')
def send_mail():
    return render_template('index.html')

@app.route('/upload_xlsx', methods=['POST'])
def upload_xlsx():
    xlsx_file = request.files['file_path']
    xlsx_file.save("temp.xlsx")
    result = send.load_data("temp.xlsx")
    if result is None:
        return render_template('404.html')
    return render_template('result.html', result = result)

if __name__=='__main__':
    app.run(debug=True)