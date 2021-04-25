from Flask import Flask, render_template, request, redirect
from backend import fill_pdf

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("website.html")

@app.route("/registration/", methods = ["GET", "POST"])
def open():
    if request.method == "POST":
        data = request.form["file"]
        return render_template("registration.html", error=None)

    else:    
        return render_template("registration.html", error=True)
    return render_template("registration.html", error=None)

if __name__ == '__main__':
    app.run
