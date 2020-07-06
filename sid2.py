import xlsxwriter
import datetime
import schedule
import time
import requests
import xlrd
import xlsxwriter
import openpyxl
from xlutils.copy import copy
import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.screenmanager import ScreenManager, Screen, FadeTransition


class ScreenManagement(ScreenManager):
    def _init_(self, **kwargs):
        super(ScreenManagement, self)._init_(**kwargs)

class Current_Window(Screen):
    def _init_(self, **kwargs):
        super(Current_Window, self)._init_(**kwargs)
        self.add_widget(Label(text='City', size_hint=(.45, .1), pos_hint={'x': .05, 'y': .85}))
        self.City = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .85})
        self.add_widget(self.City)
        self.add_widget(Label(text='Wind(m/s)', size_hint=(.45, .1), pos_hint={'x': .05, 'y': .4}))
        self.wind = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .4})
        self.add_widget(self.wind)
        self.add_widget(Label(text='Time', size_hint=(.45, .1), pos_hint={'x': .05, 'y': .55}))
        self.time = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .55})
        self.add_widget(self.time)
        self.add_widget(Label(text='Power(kw/h)', size_hint=(.45, .1), pos_hint={'x': .05, 'y': .25}))
        self.pawar = TextInput(multiline=False, size_hint=(.45, .1), pos_hint={'x': .5, 'y': .25})
        self.add_widget(self.pawar)
        self.btn5 = Button(text='calculate', size_hint=(.9, .1), pos_hint={'center_x': .5, 'y': .7})
        self.add_widget(self.btn5)
        self.btn5.bind(on_press = self.calculate)
        self.btn6 = Button(text='Fuckk Goo back!', size_hint=(.9, .1), pos_hint={'center_x': .5, 'y': .1})
        self.add_widget(self.btn6)
        self.btn6.bind(on_press = self.screen_transition)
    def calculate(self, *args):
        city = self.City.text
        workbk_out = xlsxwriter.Workbook("Wind.xlsx")
        sheet_out = workbk_out.add_worksheet()
        sheet_out.write("A1", "Time")
        sheet_out.write("B1", "Wind Speed")
        sheet_out.write("C1", "Power In Watt")
        n1 = 2
        n2 = 2
        n3 = 2
        now = datetime.datetime.now()
        a = now.strftime("%H:%M:%S")
        api_address='http://api.openweathermap.org/data/2.5/weather?appid=1222ec2c19edb278b4e39377e4138b42&q='
        url = api_address + city
        json_data = requests.get(url).json()
        wind = json_data['wind']['speed']
        # print('Wind Speed: {}'.format(wind))
        # print("Current Time is:", a )
        sheet_out.write("B2",wind )
        sheet_out.write("A2", a )
        sheet_out.write("C2",(0.5*1.23*2826*wind*wind*wind/1000))
        powar = 0.5*1.23*2826*wind*wind*wind/1000
        self.time.text = str(a)
        self.pawar.text = str(powar)
        self.wind.text = str(wind)
        workbk_out.close()
    def screen_transition(self, *args):
        self.manager.current = 'login'

class Prediction_Window(Screen):
    def _init_(self, **kwargs):
        super(Prediction_Window, self)._init_(**kwargs)
        self.add_widget(Label(text='City', size_hint=(.45, .08), pos_hint={'x': .05, 'y': 0.85}))
        self.City = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': 0.85})
        self.add_widget(self.City)
        self.add_widget(Label(text='Date', size_hint=(.45, .08), pos_hint={'x': .05, 'y': .75}))
        self.date = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .75})
        self.add_widget(self.date)
        self.add_widget(Label(text='Day', size_hint=(.45, .08), pos_hint={'x': .05, 'y': .65}))
        self.day = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .65})
        self.add_widget(self.day)
        self.add_widget(Label(text='Power Generated(kw/h)', size_hint=(.45, .08), pos_hint={'x': .05, 'y': .40}))
        self.pawar = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .40})
        self.add_widget(self.pawar)
        self.add_widget(Label(text='Deficiet', size_hint=(.45, .08), pos_hint={'x': .05, 'y': .27}))
        self.Deficiet = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .27})
        self.add_widget(self.Deficiet)
        self.add_widget(Label(text='Solar', size_hint=(.45, .08), pos_hint={'x': .05, 'y': .14}))
        self.solarX = TextInput(multiline=False, size_hint=(.45, .08), pos_hint={'x': .5, 'y': .14})
        self.add_widget(self.solarX)
        self.btn7 = Button(text='predict', size_hint=(.9, .1), pos_hint={'center_x': .5, 'y': .50})
        self.add_widget(self.btn7)
        self.btn7.bind(on_press = self.pressed)
        self.btn8 = Button(text='Fuckk Goo back!', size_hint=(.9, .08), pos_hint={'center_x': .5, 'y': .03})
        self.add_widget(self.btn8)
        self.btn8.bind(on_press = self.screen_tronsition)
    def pressed(self, instance):
        #city = self.city.text
        A = self.City.text
        B = self.date.text
        C = 'https://www.wunderground.com/hourly/in/'
        D = self.day.text
        E = ('Data_'+D)
        url = C + A +'/date/' +B
        print(url)
        workbook = xlrd.open_workbook(E+'.xlsx')
        sheet = workbook.sheet_by_index(0)
        row_count = sheet.nrows
        lis1=[]
        lis2=[]
        lis3=[]
        lis4=[]
        dict1={}
        dict2={}
        dicttime={}
        time=1
        time2=1
        rw_no=1
        n1=1
        n2=0

        xyz=[]
        def solar():
            n1 = -1
            n2 = -1
            n = -1 #index of lis3
            for i in range (1,row_count):
                solar5 = sheet.cell_value(i,1)#lis2=theory
                lis2.append(solar5)
            for i in range (1,row_count):
                solar3 = sheet.cell_value(i,0)
                lis3.append(solar3)#lis3=times

        for i in range (1,row_count):
            sv = sheet.cell_value(i,2)
            lis1.append(sv)
        for speed in lis1:
            dict2[time]=speed
            time+=1
        for i in dict2:
            if dict2[i] < 3.3:
                dict2[i] = 0
            elif dict2[i] > 20:
                dict2[i] = 0
            else:
                dict2[i] =  (0.5 * 1.23 *2826 * dict2[i] * dict2[i] * dict2[i])/1000
        rb = xlrd.open_workbook(E+'.xlsx')
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        ml = 1
        for j in range(1,25):
            w_sheet.write(j,4,dict2[ml])
            wb.save(E+'.csv')
            ml += 1
        need = 15000
        sum = 0
        for i in dict2:
            sum += dict2[i]
            self.pawar.text = str(sum)
        if sum <  need:
            self.Deficiet.text = str(need - sum)
            solar()
            for i in lis2:
                if i == "Partly Cloudy":
                    xyz.append(n1)
                n1+=1
        elif sum > need:
            print('Excess Power Generated:', sum - need)
            n1 += 1
        else:
            print('Requirements Stisfied:')
            n1 += 1
        w_sheet.write(27,4,sum)
        wb.save(E+'.csv')
        # print(lis2[n1])
        print("Use solar at {} O'clock".format(xyz))
        self.solarX.text = str(dict2)
    def screen_tronsition(self, *args):
        self.manager.current = 'login'

class LoginWindow(Screen):
        def _init_(self, **kwargs):
            super(LoginWindow, self)._init_(**kwargs)
            self.add_widget(Label(text="Welcome!", size_hint=(.9, .3), pos_hint={'x': .06, 'y': .7}))
            self.btn2 = Button(text='current', size_hint=(.9, .3), pos_hint={'center_x': .5, 'y': .4})
            self.add_widget(self.btn2)
            self.btn2.bind(on_press = self.screen_transition)
            self.btn3 = Button(text='prediction', size_hint=(.9, .3), pos_hint={'center_x': .5, 'y': .09})
            self.add_widget(self.btn3)
            self.btn3.bind(on_press = self.screen_tronsition)

        def screen_tronsition(self, *args):
            self.manager.current = 'Prediction'
        def screen_transition(self, *args):
            self.manager.current = 'Current'

class Application(App):
    def build(self):
        sm = ScreenManagement(transition=FadeTransition())
        sm.add_widget(LoginWindow(name='login'))
        sm.add_widget(Current_Window(name='Current'))
        sm.add_widget(Prediction_Window(name='Prediction'))
        return sm

if __name__ == "__main__":
    Application().run()