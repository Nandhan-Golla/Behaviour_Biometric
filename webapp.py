import flask
import _data_run as dr
from sklearn.model_selection import train_test_split
import pandas as pd
from sklearn.linear_model import LogisticRegression

app = flask.Flask(__name__)
model = LogisticRegression()
"""data = pd.DataFrame(pd.read_excel('D:\local-code\Data.xlsx'))

X = data.drop(columns='User')
Y = data['User']

X_train, X_test, y_train, y_test = train_test_split(X,Y, test_size=0.2, random_state=1423)"""



@app.route("/", methods=['POST', 'GET'])
def main():

    if flask.request.method == 'POST':

        user = flask.request.form['name']
        dr.process(user)

        print("check1")

       

        
    return flask.render_template('index.html')

        
#final_result = ''        

@app.route('/target')
def main1():
    #dr.process(user)
     data = pd.DataFrame(pd.read_excel('D:\local-code\Data.xlsx'))
     X = data.drop(columns='User')
     Y = data['User']

     X_train, X_test, y_train, y_test = train_test_split(X,Y, test_size=0.2, random_state=1423)
     model.fit(X_train, y_train)

     


     return f"Welcome {model.predict(X_test)}"




if __name__ == '__main__':
    app.run(debug=True)