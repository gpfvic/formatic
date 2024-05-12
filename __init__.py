# conding: utf-8
import os
from flask import Flask, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = os.pardir
ALLOWED_EXTENSIONS = {'docx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/hello/')
@app.route('/hello/<name>')
def hello(name=None):
    return render_template('hello.html', name=name)



@app.route('/upload')
def upload():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['POST'])
def upload_file():
   if 'file' in request.files: 
      f = request.files['file']
      f.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename)))
      return 'file uploaded successfully'



@app.route('/list')
def list_files():
    path = os.path.expanduser(UPLOAD_FOLDER)
    tree = dict(name=os.path.basename(path), children=[])
    try: lst = os.listdir(path)
    except OSError:
        pass #ignore errors
    else:
        for name in lst:
            fn = os.path.join(path, name)
            if os.path.isdir(fn):
                tree['children'].append(dict(name=name))

    return render_template('dirtree.html', tree=tree)