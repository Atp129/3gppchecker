from flask import Flask


app = Flask(__name__)

@app.route("/")
def index():
    return "this is the very first page"

if __name__ == "__main__":
    app.run()


