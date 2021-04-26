import os
from flask import Flask, flash, render_template, request, redirect, url_for
from backend import write
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = '/uploads'
ALLOWED_EXTENSIONS = {'xlsx'}


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("website.html")

@app.route("/about.html")
def about():
    return render_template("about.html")

@app.route("/courselisting.html")
def courselisting():
    return render_template("courselisting.html")

@app.route("/howto.html")
def howto():
    return render_template("howto.html")

@app.route("/registration.html", methods = ["GET", "POST"])
@app.route('/upload')
def upload_file():
   return render_template('registration.html')

@app.route('/output.html', methods = ['GET', 'POST'])
def uploaded_file():
   if request.method == 'POST':
      f = request.files['file']
      f.save(secure_filename(f.filename))
      write(f.filename)
      return render_template('output.html')
    

if __name__ == '__main__':
    app.run(debug=True)
