from flask import Flask, render_template, url_for

base = Flask(__name__)


@base.route('/')
@base.route('/home')
def index():
    return render_template('index.html')


@base.route('/about')
def about():
    return render_template('about.html')


@base.route('/user/<string:name>/<int:id>')
def user(name, id):
    return "User page:" + name + ' - ' + str(id)


if __name__ == '__main__':
    base.run(debug=True)
