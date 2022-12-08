from csv import excel_tab
from re import sub
import re
from flask import Flask,request,send_file
from flask.templating import render_template
import os
import shutil
import range_roll_generator
import generate_full

app = Flask(__name__)
ALLOWED_EXTS = {"csv"}
ALLOWED_PICS = {"jpeg","jpg","png"}
def check_file(file):
    return '.' in file and file.rsplit('.',1)[1].lower() in ALLOWED_EXTS

def check_img_ext(file):
    return '.' in file and file.rsplit('.',1)[1].lower() in ALLOWED_PICS


not_present = []
@app.route("/not_present_list")
def not_present_function():
    
    return render_template("not_present.html", np = not_present)


@app.route("/download")
def download_transcripts():
    path = ".\\Transcripts.zip"
    return send_file(path, as_attachment=True)

def refresh_server():
    output_path = ".\\outputs"
    signature_path = "./signature_photo"
    stamp_path = "./stamp_photo"
    uploads_path = ".\\uploads"
    try:
        if os.path.exists(output_path):
            shutil.rmtree(output_path)
        if os.path.exists(uploads_path):
            shutil.rmtree(uploads_path)
        
        if os.path.exists(signature_path):
            shutil.rmtree(signature_path)
        if os.path.exists(stamp_path):
            shutil.rmtree(stamp_path)
        if os.path.exists(".\Transcripts.zip"):
            os.remove(".\Transcripts.zip")
        not_present = []
        return True

    except Exception:
        return False

def check_range(range1, range2):

    if len(range1)!=8 or len(range2)!=8:
        return False

    year1 = range1[0:2]
    course1 = range1[2:4]
    branch1  = range1[4:6]
    roll1 = range1[6:8]

    year2 = range2[0:2]
    course2= range2[2:4]
    branch2  = range2[4:6]
    roll2 = range2[6:8]

    if (year1.isnumeric()
     and course1.isnumeric()
     and branch1.lower().isalpha() 
     and roll1.isnumeric() 
     and year2.isnumeric() 
     and course2.isnumeric() 
     and branch2.lower().isalpha() 
     and roll2.isnumeric()
     and year1.lower()==year2.lower()
     and branch1.lower()==branch2.lower()
     and course1.lower()==course2.lower()
     and (int(roll1) <= int(roll2))):
        return True
    return False




@app.route("/", methods = ["GET", "POST"])
def build():

    error = None
    success = None
    if request.method == "POST":
        if 'submit_files' in request.form:
            if 'grades' not in request.files or 'namesroll' not in request.files or 'subjectmaster' not in request.files:
                error = "Please upload all files together!"
                return render_template("index.html", error = error)

            grade_file = request.files["grades"]
            grade_filename = grade_file.filename

            namesroll_file = request.files["namesroll"]
            namesroll_filename = namesroll_file.filename

            subjectmaster_file = request.files["subjectmaster"]
            subjectmaster_filename = subjectmaster_file.filename

            if grade_filename=='' or namesroll_filename=='' or subjectmaster_filename == '':
                error = "Please upload all files!"
                return render_template("index.html", error = error)
            if not (check_file(grade_filename) and check_file(namesroll_filename) and check_file(subjectmaster_filename)):
                error = "Please upload csv type files!"
                return render_template('index.html', error = error)

            upload_path = "./uploads"
            if os.path.exists(upload_path):
                shutil.rmtree(upload_path)
                os.mkdir(upload_path)
            else:
                os.mkdir(upload_path)
            
            grade_file.save(os.path.join(upload_path,"grades.csv"))
            namesroll_file.save(os.path.join(upload_path,"names-roll.csv"))
            subjectmaster_file.save(os.path.join(upload_path,"subjects_master.csv"))
            
            return render_template('index.html', error=error, success = "Upload success" )
        
        if 'submit_signature' in request.form:
            
            error_signature = None
            signature_file = request.files["signature"]
            signature_filename = signature_file.filename

            if signature_filename == '':
                error_signature = "Please upload the signature file!"
                return render_template('index.html',error_signature = error_signature)

            if not (check_img_ext(signature_filename)):
                error_signature = "Please upload PNG,JPEG,JPG type files!"
                return render_template('index.html', error_signature = error_signature)

            upload_path = "./signature_photo"

            if os.path.exists(upload_path):
                shutil.rmtree(upload_path)
                os.mkdir(upload_path)
            else:
                os.mkdir(upload_path)
            signature_file.save(os.path.join(upload_path,"signature.png"))
            return render_template('index.html', error_signature=error_signature, success_signature = "Upload success" )
        
        if 'submit_stamp' in request.form:

            error_stamp = None
            stamp_file = request.files["stamp"]
            stamp_filename = stamp_file.filename

            if stamp_filename == '':
                error_stamp = "Please upload the stamp file!"
                return render_template('index.html',error_stamp = error_stamp)
            
            
            if not (check_img_ext(stamp_filename)):
                error_stamp = "Please upload PNG,JPEG,JPG type files!"
                return render_template('index.html', error_stamp = error_stamp)

            upload_path = "./stamp_photo"

            if os.path.exists(upload_path):
                shutil.rmtree(upload_path)
                os.mkdir(upload_path)
            else:
                os.mkdir(upload_path)
            stamp_file.save(os.path.join(upload_path,"stamp.png"))
            return render_template('index.html', error_stamp=error_stamp, success_stamp = "Upload success" )


        if 'generate_range' in request.form:
            
            if not (check_range(range1 = request.form['range1'], range2 = request.form['range2'])):
                error_range = "Please enter a valid range"
                return render_template('index.html', error_range = error_range)
            try:
                range_roll = []
                template_roll = request.form['range1'][0:6]
                start_roll = int(request.form['range1'][6:8])
                end_roll = int(request.form['range2'][6:8])

                for i in range(start_roll,end_roll+1):
                    if i/10 < 1:
                        range_roll.append(template_roll.upper() + "0" + str(i))
                    else:
                        range_roll.append(template_roll.upper() + str(i))
                
                global not_present 
                not_present = range_roll_generator.range_generate(range_roll)
            except Exception:
                return render_template("index.html", error_range = "Error in manipulating files! Please upload correct files!")
            
            if os.path.exists(".\\outputs"):
                shutil.make_archive("Transcripts",'zip',".\\outputs" )
                return render_template("index.html", success_range = "The files have been generated sucessfully in the output folder!",
                    not_present = not_present
                )
            return render_template("index.html", error_range = "Could not generate files in the given range",
                    not_present = not_present
                )
        
        
        if 'generate_full' in request.form:
            try:
                if not os.path.exists(".\\uploads"):
                    return render_template("index.html", error_full = "Please upload the files first!")
                generate_full.generate_all_pdf()
            except Exception:
                return render_template("index.html", error_full = "Error in manipulating files! Please upload correct files!")

            if os.path.exists(".\\outputs"):
                shutil.make_archive("Transcripts",'zip',".\\outputs" )    
                return render_template("index.html", success_full = "The files have been generated sucessfully in the output folder!")
            return render_template("index.html", error_full = "No roll number found!")
        
        if 'refresh' in request.form:
            if refresh_server():
                return render_template('index.html',success_refresh = '1')
            return render_template('index.html',error_refresh = "There was some error in refreshing, please restart the server!")

    return render_template("index.html")

app.run(debug=True)