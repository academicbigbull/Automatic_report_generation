from flask import Flask,render_template#导入jinja模块

app = Flask(__name__)

class User:
    def __init__(self,username,email):
        self.username=username
        self.email=email


@app.route('/')
def hello_world():  # put application's code here
    person={#传字典类
        'username':'梁爽',
        'email':'321'
    }
    user=User(username='向豪杰',email='123')#传类对象
    return render_template('index.html',user=user,person=person)

@app.route('/blog/<blog_id>')
def blog_detail(blog_id):
    return render_template('blog_detail.html',blog_id=blog_id,username='梁爽')#将id和username传到blog_id。html页面

@app.route('/filter')
def filter():
    user = User(username='向豪杰', email='123')  # 传类对象
    return render_template('filter.html',user=user)

#jinja模块使用控制语句
@app.route('/control')
def control_statement():
    age=19
    return render_template('control.html',age=age)

#模块的继承
@app.route('/child')
def extend_statement():
    return render_template('child.html')

@app.route('/child1')
def extend_statement1():
    return render_template('child1.html')

#加载静态文件
@app.route('/static')
def static_statement():
    return render_template('static.html')

if __name__ == '__main__':
    app.run()
