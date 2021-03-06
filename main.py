import os
from flask import Flask, flash, render_template, request, redirect, url_for
from backend import write
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
ALLOWED_EXTENSIONS = {'xlsx'}
def allowed_file(filename):
   return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/index.html")
def home():
    return render_template("/index.html")

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
def registration():
   """
   modified from 
   https://blog.miguelgrinberg.com/post/handling-file-uploads-with-flask#:~:text=The%20actual%20file%20data%20can,then%20the%20file%20is%20discarded.
   """
   if request.method == 'POST':
       if 'file' not in request.files:
           print('No file attached in request')
           return redirect(request.url)
       file = request.files['file']
       if file.filename == '':
           print('No file selected')
           return redirect(request.url)
       if file and allowed_file(file.filename):
           filename = secure_filename(file.filename)
           file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
           write(os.path.join(app.config['UPLOAD_FOLDER'], filename))
           return render_template('output.html')
   return render_template('registration.html')


if __name__ == '__main__':
    app.run(debug=True)
