from flask import Flask , request ,render_template,flash,redirect,session,url_for,Response
from sqlalchemy import create_engine
from sqlalchemy.orm import scoped_session,sessionmaker
from passlib.hash import sha256_crypt
from flask_mail import Mail, Message
from time import sleep
import random
import xlwt
import io
import xlrd
import math
import logging

logger = logging.getLogger(__name__)
logging.basicConfig(filename='test.log', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

engine = create_engine("mysql://root:@localhost/mydb")
db = scoped_session(sessionmaker(bind=engine))
app = Flask(__name__)
app.config['SECRET_KEY'] = 'thisissecretkey'
app.config['MAIL_SERVER'] = "smtp.gmail.com"
app.config['MAIL_PORT'] = 465
app.config['MAIL_USE_SSL'] = True
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USERNAME'] = 'aasthsaxenadummy@gmail.com'
app.config['MAIL_PASSWORD'] = 'Victory@1947'
app.config['MAIL_DEFAULT_SENDER'] = 'aasthsaxenadummy@gmail.com'

mail = Mail(app)

@app.route("/" , methods=["GET","POST"])
def myApp():
    if request.method == "POST":
        name = request.form.get("name")
        password = request.form.get("password")

        userdata = db.execute("SELECT name FROM user WHERE name=:name",{"name":name}).fetchone()
        passdata = db.execute("SELECT password FROM user WHERE name=:name", {"name": name}).fetchone()
        # passw = str(passdata)
        if userdata == None:
            flash("No user found", "danger")
            logger.warning('No user found in the database')
            return render_template("myApp.html")
        else:
            for data in passdata:
                if sha256_crypt.verify(password, data):
                    session['name'] = name
                    flash('You are now logged in','success')
                    return redirect(url_for('home'))
                else:
                    flash("Incorrect Password", "danger")
                    logger.error('Incorrect Password')
                    return render_template("myApp.html")

    return render_template("myApp.html")

@app.route('/register', methods = ['GET','POST'])
def register():
    if request.method =='POST':
        name = request.form.get('name')
        email = request.form.get('email')
        password = request.form.get('password')
        role = request.form.get('role')
        hash = sha256_crypt.hash(str(password))
        db.execute("INSERT INTO user(name,email,password,role) VALUES(:name,:email,:password,:role)",
                                    {"name":name,"email":email,"password":hash,"role":role})
        db.commit()
        flash("Registered Successfully!","success")
        logger.info('Registered Successfully : {} - {}'.format(name,email))
        return redirect(url_for('myApp'))

    return render_template("register.html")

@app.route('/myApp/admin/home')
def admin():
    return render_template('admin.html')

@app.route('/myApp/user/home')
def home():
    return render_template('home.html')

@app.route('/logout')
def logout():
    session.clear()
    flash("You are logged out", "success")
    logger.info('Successfully Logged out')
    return redirect(url_for("myApp"))

@app.route('/create',methods=['GET','POST'])
def create():
    if request.method =='POST':
        name = request.form.get('name')
        password = request.form.get('password')
        role = request.form.get('role')
        email = request.form.get('email')
        hash = sha256_crypt.hash(str(password))
        db.execute("INSERT INTO user(name,email,password,role) VALUES(:name,:email,:password,:role)",
                                    {"name":name,"email":email , "password":hash,"role":role})
        db.commit()
        flash("User created", "success")
        logger.info('User created Successfully : {} - {}'.format(name, email))
        return redirect(url_for('admin'))

    return render_template("admin.html")

@app.route('/delete/<string:id>/<int:page>' , methods=['POST'])
def delete(id,page):
    db.execute("DELETE FROM user WHERE id=:id", {"id": id})
    db.commit()
    flash("User Deleted", "success")
    pages = (page-1)*5
    users = db.execute(f"SELECT * FROM user LIMIT 5 OFFSET {pages}")
    count = db.execute("SELECT COUNT(*) FROM user").scalar()
    html_page = math.ceil(count/5)
    next = page+1
    previous = page-1
    total = html_page+1
    db.commit()
    logger.info('user deleted successfully')
    return render_template('view.html',  values=users, html_page=html_page , page=page, total=total , next=next , previous=previous)

@app.route('/view/<int:page>')
def view(page):
    pages = (page-1)*5
    users = db.execute(f"SELECT * FROM user LIMIT 5 OFFSET {pages}")
    count = db.execute("SELECT COUNT(*) FROM user").scalar()
    html_page = math.ceil(count/5)
    next = page+1
    previous = page-1
    total = html_page+1
    db.commit()
    return render_template("view.html", values=users, html_page=html_page , page=page, total=total , next=next , previous=previous)

@app.route('/download')
def download():
    try:
        result = db.execute("SELECT * FROM user")
        output = io.BytesIO()
        workbook = xlwt.Workbook()
        sh = workbook.add_sheet('User Details')

        sh.write(0,0, 'id')
        sh.write(0,1, 'name')
        sh.write(0,2, 'email')
        sh.write(0,3, 'role')
        sh.write(0,4, 'token')


        idx=0
        for row in result:
            sh.write(idx+1,0,str(row['id']))
            sh.write(idx + 1, 1, str(row['name']))
            sh.write(idx + 1, 2, str(row['email']))
            sh.write(idx + 1, 3, str(row['role']))
            sh.write(idx + 1, 4, str(row['token']))
            idx+=1

        workbook.save(output)
        output.seek(0)
        return Response(output,mimetype="user/ms-excel", headers={"Content-Disposition":"attachment;filename=user_list.xls"})
    except Exception as e:
        print(e)


@app.route('/forgot')
def forgot():
    token = random.randint(1,10000)
    return render_template('forgot.html', value=token)

@app.route('/mail')
def send_mail():
    msg = Message('Hey There', recipients=['hayix54617@aprimail.com'])
    mail.send(msg)
    return 'msg sent'

@app.route('/reset_password/<string:token>', methods=['POST'])
def reset_pass(token):
    email = request.form.get('email')
    emaildata = db.execute("SELECT email FROM user WHERE email=:email", {"email": email}).fetchone()
    if emaildata == None:
        flash("No user found!")
        logger.warning('No user found in the database')
        return render_template("myApp.html")
    else:
        db.execute("UPDATE user SET token = (:token) WHERE email = (:email)", {"token": token, "email": email})
        db.commit()
        return render_template('reset.html', value=token)
    return render_template('reset.html')

@app.route('/reset/<string:token>', methods= ['POST'])
def reset(token):
    password = request.form.get('password')
    confirm = request.form.get('confirm')
    hashed = sha256_crypt.encrypt(str(password))
    if password != confirm:
        flash("Couldn't match passwords","danger")
        return render_template('reset.html')
    else:
        db.execute("UPDATE user SET password = (:password)  WHERE token = (:token)", {"password": hashed,"token": token})
        db.execute("UPDATE user SET token = 0 WHERE token = (:token)",
                   {"token": token})
        db.commit()
        flash("Reset successful","success")
        logger.info('Password Reset Successful')
        return render_template("myApp.html")
    return render_template('reset.html')


@app.route('/upload' , methods=['POST'])
def uplaod():
    book = xlrd.open_workbook("static/upload_dir/Book1.xlsx")
    sheet = book.sheets()[0]
    for row in range(1,sheet.nrows):
        # id = sheet.cell(row,0).value
        name = sheet.cell(row, 1).value
        email = sheet.cell(row, 2).value
        password = sheet.cell(row, 3).value
        hash = sha256_crypt.hash(str(password))
        role = sheet.cell(row, 4).value
        token = sheet.cell(row, 5).value
        db.execute('INSERT INTO user (name,email,password,role,token) VALUES (:name,:email,:password,:role,:token)',{ "name":name,"email":email,"password":hash,"role":role,"token":token})
        db.commit()

    flash("File uploaded",'success')
    logger.info('File upload Successful')
    return render_template('myApp.html')


@app.route('/search',methods=['POST'])
def search():
    name = request.form.get('name')
    users = db.execute("SELECT * FROM user WHERE name=:name", {"name": name})
    return render_template("search.html", values=users)

@app.route('/stream')
def stream():
    def generate():
        with open('test.log') as f:
            while True:
                yield f.read()
                sleep(1)
    return app.response_class(generate(),mimetype='text/plain')



if __name__ == '__main__':
    app.run(debug=True)