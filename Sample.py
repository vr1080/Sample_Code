####################################################
#Author: Vema Reddy, 2015
#Purpose: GUI interface to work with ETS770 IC Automated Handler

#Basics of program
#this program will send a start command to the machine that contains microprocessor to begin test
#It will test microprofessor for 8 different tests
#If pass, it will go into the Pass bin
#If fail, it will go into the Fail bin
#if it loses connection the machine, it will stop testing after 10 failed attempts
#The machine will continue to test microprocessor until their is none left

#LED lights
#After each test pass, The LED light will change fromyellow to green
#After each test fails, The LED light will change from yellow to red
#yellow light means - been reset
#purple light means - lost connection

#Data
#There is data at the bottom telling how many parts pass, fail, and yield rate
#This data can also be sent txt file for furthur analysis
####################################################
#Library needed
import collections
import timeit
import Tkinter as tkinter
import serial
import tkMessageBox
import time
import datetime
import webbrowser
import winsound, sys
#####################################################
#Global Variables Needed for Serial Communication

start = 'y'
dummy = 'j'

#Command code
T2 = "T2\r\n"
F2 = "F2\r\n" #failed at continuity #tested

T3 = "T3\r\n"
F3 = "F3\r\n" #failed at chip erase

T4 = "T4\r\n"
F4 = "F4\r\n" #failed at Erase test

T5 = "T5\r\n"
F5 = "F5\r\n" #failed at Write test 

T6 = "T6\r\n"
F6 = "F6\r\n" #failed at Wires 

T7 = "T7\r\n"
F7 = "F7\r\n" #Failures at Write All 

#repeat of chip erase test
T8 = "T8\r\n"
#F3 failure possible

P1 = "P1\r\n" #passed

#translate command code into message
dict_codeL = {'F2\r\n': 'Continuity Failure', 'F3\r\n': 'Chip Erase Failure', 'F4\r\n': 'Erase Test Failure',
              'F5\r\n': 'Write Test Failure', 'F6\r\n': 'Wire Failure', 'F7\r\n': 'Write All Failure',
              'P1\r\n': "Passed"}

####################################################
#define GUI interface
def main():
    root = Application()
    root.setup()
    root.resizable(width=False, height=False) #can't resize
    root.title('7602A_Functional_Test') #title
    #root.geometry('200x300') #window size
    root.protocol('WM_DELETE_WINDOW', root.exit_app) #double check if want to delete
    try:
        root.iconbitmap('Tekmos.ico')
        #r'C:/Users/Tekmos/Documents/IPython Notebooks/Interface/favicon.ico'
    except:
        try:
            root.iconbitmap(r'C:/Program Files/Tekmos_7602/Tekmos.ico')
            #r'C:/Users/Tekmos/Documents/IPython Notebooks/Interface/favicon.ico'        
        except:
            root.textframe.insert(tkinter.END,"Icon error")
            root.textframe.insert(tkinter.END,"\n") 
            
    root.mainloop()

#GUI Application Class
class Application(tkinter.Tk):

    def setup(self):
###########################################################################################
#Frames
        self.root = self
    
        #top frame
        tf = self.__top_frame = tkinter.Frame(self.root)
        self.__top_frame.grid()
        #middle frame
        mf = self.__middle_frame = tkinter.Frame(self.root)
        self.__middle_frame.grid()
        #bottom frame
        bf = self.__bot_frame = tkinter.Frame(self.root)
        self.__bot_frame.grid()
        #bottomest frame
        bbf = self.__bbot_frame = tkinter.Frame(self.root)
        self.__bbot_frame.grid()
        
####################################################### 
#Menu
        self.menubar = tkinter.Menu(self.root)
        self.menubar.add_command(label="About",command=self.about)
        self.menubar.add_command(label="Quit",command=self.exit_app)
        self.config(menu=self.menubar)
###########################################################################################
#Important variables

#Variables
        self.__chip_number = 0
        self.__chip_pass_num = 0
        self.__chip_fail_num = 0
        self.__chip_yield_num = 0
        self.__unknown_num = 0   #failing due to connection issues
        self.__poll_ser = 3   #time to poll ser(in sec)
        self.__poll_T6_ser = 5.2
        self.__baud_rate = 19200
        self.__wait_0_5s = 500 #1/2 second
#Bools
        self.test_start = True #control if able to start test
        self.test_done = False #tells if test is finished
        self.PASS = True #individual test in serial communciation
        self.T6_state = False
#list
        self.list_results = []
#filename
        self.filename = "test_7602.txt"
###########################################################################################
#Top Frame       
        Port_Name = tkinter.Label(tf, text='Port:').grid(row=0, column=0)
        self.patt = tkinter.IntVar()
        
        self.set_arrow = tkinter.Spinbox(tf, from_=0, to=20, width=5, textvariable=self.patt).grid(row=0, column=2)
        self.set_button = tkinter.Button(tf, text ='Set', command=self.__setup_serial)
        self.set_button.grid(row= 0, column=3,padx=1)  
###########################################################################################
#Middle Frame
        self.__widgets = collections.OrderedDict((
            ('COT', 'Continuity Test'), ('CHE', 'Chip Erase'),
            ('ERT', 'Erase Test'), ('WRT', 'Write Test'),
            ('WIRT', 'Wire Reading Test'), ('WIT', 'Wires Test'),
            ('WRAT', 'Write All Test'), ('DO', 'Done')))

        for row, (key, value) in enumerate(self.__widgets.items()):
            label = tkinter.Label(mf, text=value+':')
            label.grid(row=row, column=0, sticky=tkinter.E)
            canvas = tkinter.Canvas(mf, bg='green', width=10, height=10)
            canvas.grid(row=row, column=1)
            self.__widgets[key] = label, canvas

        self.__cn = tkinter.Label(mf, text='Chip Number:')
        self.__cn.grid(row=8, column=0, sticky=tkinter.E)
        
        self.__display = tkinter.Label(mf)
        self.__display.grid(row=8, column=1, sticky=tkinter.E)
        
        #print self.__widgets
#######################################################################################
#Bottom Frame

        #start button
        self.__button = tkinter.Button(bf, text='START', command=self.__start_pre)
        self.__button.grid(row=0, column=1,sticky=tkinter.E)
        self.__button['state'] = tkinter.DISABLED
        
        #stop button
        self.__button2 = tkinter.Button(bf, text='STOP', command=self.__stop_pre)
        self.__button2.grid(row=0, column=2,sticky=tkinter.E)
        self.__button2['state'] = tkinter.DISABLED

        #RESULTS button
        self.__results_button = tkinter.Button(bf, text='Results', command=self.__result_action)
        self.__results_button.grid(row=0, column=3,sticky=tkinter.E)
        self.__results_button['state'] = tkinter.DISABLED

#important numbers
        #PASSED
        Passed_Label = tkinter.Label(bf, text='Passed').grid(row=2, column=0, sticky = tkinter.E)
        self.__passed_display = tkinter.Label(bf)
        self.__passed_display.grid(row=2, column=1, sticky=tkinter.E)
        #self.__passed_display['text'] = "Passed:   " + str(self.__chip_pass_num)
        
        #FAILED
        Failed_Label= tkinter.Label(bf, text='Failed').grid(row=2, column=2, sticky = tkinter.E)
        self.__failed_display = tkinter.Label(bf)
        self.__failed_display.grid(row=2, column=3, sticky=tkinter.E)
        #self.__failed_display['text'] = "Failed:   "+ str(self.__chip_fail_num)
        
        #Yield
        Yield_Label= tkinter.Label(bf, text='Yield').grid(row=2, column=4, sticky = tkinter.E)
        self.__yield_display = tkinter.Label(bf)
        self.__yield_display.grid(row=2, column=5, sticky=tkinter.E)
        #self.__yield_display['text'] = "Yield:   " + str(self.__chip_yield_num)        
############################################################################################
#Bottomest Frame!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!Fix Me
        text = self.textframe = tkinter.Text(bbf, height=20, width=50)
        text.grid()
        scrl = tkinter.Scrollbar(bbf, command=text.yview)
        text.config(yscrollcommand=scrl.set)
        scrl.grid(row=0, column=1, sticky='ns')
        x = "Welcome" + "\n"
        self.textframe.insert(tkinter.END,x)
        #self.textframe.delete(1.0, tkinter.END)
############################################################################################
#Functions         
##################################################
#STOP Button Action
    def __stop_pre(self):
        
        self.test_start = False
        self.__button2['state'] = tkinter.DISABLED #TURN OFF STOP BUTTON
        
        if self.test_done:
            self.__button['state'] = tkinter.NORMAL #Turn ON START Button
            self.test_start = True
            self.__print_to_file() #print to file      
        else:
            self.after(self.__wait_0_5s, self.__stop_pre)
###################################################
#Start Button Action
#Set the Right condition when the Start Button has been pressed
    def __start_pre(self):
        #reset to correct condition
        self.__button['state'] = tkinter.DISABLED #Turn OFF Start Button
        self.__button2['state'] = tkinter.NORMAL  #Turn ON Stop Button
        self.__results_button['state'] = tkinter.DISABLED
        self.textframe.delete(1.0, tkinter.END)
        
        self.__chip_number = 0
        self.__chip_pass_num = 0
        self.__chip_fail_num = 0
        self.__chip_yield_num = 0
        
        self.__display['text'] = self.__chip_number
        self.__failed_display['text'] = self.__chip_fail_num
        self.__passed_display['text'] = self.__chip_pass_num
        self.__yield_display['text'] =  self.__chip_yield_num
        
        self.root.update()
        
        self.test_start = True
        self.test_done = False
        
        #start test
        self.__start_button()

	#Going throught each test and updating the GUI interface	
    def __start_button(self):
        
        if self.test_start:
            self.__reset_color() #change color to yellow(RESET STATUS)
            self.__chip_number += 1
            self.__display['text'] = str(self.__chip_number)
            
            self.__run_serial_com() #Go through Serial Communication
            
            if self.PASS:   #PASSED Part
                self.after(500, self.__start_button) # 1/2 sec til start again
            else:           #Failed Part
                self.after(3000, self.__start_button) # 3 sec til start again
        else:
            self.__button['state'] = tkinter.NORMAL
            self.__results_button['state'] = tkinter.NORMAL            
            self.test_done = True
			
#########################################################
#Result button action
    def __result_action(self):
        webbrowser.open(self.filename)
#########################################################
#Serial Part

#Establish Serial Connection
    def __setup_serial(self):
	
#Open port
        try:
            self.ser = serial.Serial()
            self.ser.baudrate = self.__baud_rate
            self.ser.port = self.patt.get() - 1
            self.ser.timeout = self.__poll_ser
            self.ser
            self.ser.open()
            self.__button['state'] = tkinter.NORMAL #Turn ON Start Button
            self.set_button['state'] = tkinter.DISABLED  #Turn OFF Stop Button
        
            #dummy test
            #self.ser.write(dummy)
            #test_j = self.ser.readline() #expect j             
            #print test_j
            
            self.__reset_color()
            self.__chip_number = 0
            
            self.test_start = True
            self.test_done = False
        
		#in Case of Error    
        except serial.SerialException:
            print "No connection to the device could be established"
            self.Wrong_port()    
        self.ser.close()

    def __run_serial_com(self):       
#Send the Start Command and go through the 8 test

        #start up commands
        self.ser.open()
        #print self.ser.isOpen() #is port on
        
        #send command
        self.ser.write(start)
        y = self.ser.readline() #expect y        
        self.PASS = True
        
        startTimer = timeit.default_timer() #start timer
        
		self.__serial_test(T2,F2,'COT')
        end = timeit.default_timer()- startTimer
        
		#print end
        self.__serial_test(T3,F3,'CHE')
        end = timeit.default_timer()- startTimer
        
		#print end
        self.__serial_test(T4,F4,'ERT')
        end = timeit.default_timer()- startTimer

        #print end     
        self.__serial_test(T5,F5,'WRT')
        end = timeit.default_timer()- startTimer

        #print end
        self.__serial_test(T6,F6,'WIRT')
        end = timeit.default_timer()- startTimer

        #print end
        self.__serial_test(T7,F7,'WIT')
        end = timeit.default_timer()- startTimer

        #print end
        self.__serial_test(T8,F3,'WRAT')
        end = timeit.default_timer()- startTimer

        #print end            
        self.__serial_test(P1,F2,'DO')
        end = timeit.default_timer()- startTimer
        #print end
        
        self.ser.close()
        #print self.ser.isOpen()

        #Display Yield
        if self.__chip_number > 0:
            self.__chip_yield_num = int(100*(float(self.__chip_pass_num)/float(self.__chip_pass_num+self.__chip_fail_num)))
            self.__yield_display['text'] = str(self.__chip_yield_num) +'%'
            
    #check each individual test
    def __serial_test(self,a,b,c):

#a is Pass message
#b is Fail message
#c is LED
		
		#change timeout variable on T6, this test takes longer
        if self.PASS:
            self.ser.timeout = self.__poll_ser
            if a == T6:
                self.ser.timeout = self.__poll_T6_ser
            x=self.ser.readline()
            
            #check if passed test
            if (x == a): 
                self.__widgets[c][1]['bg'] = 'green'
                #self.__widgets[c][1].update_idletasks()
                #check if passed all test
                if (x == P1): 
                    self.__chip_pass_num += 1 
                    self.__passed_display['text'] = self.__chip_pass_num
                    self.list_results.append(["Chip Number", self.__chip_number, dict_codeL[x]])
                    self.textframe.insert(tkinter.END,self.list_results[-1])
                    self.textframe.insert(tkinter.END,"\n")
            
            #check if failed test
            elif (x == b): 
                self.__widgets[c][1]['bg'] = 'red'
                self.__chip_fail_num += 1
                self.__failed_display['text'] = self.__chip_fail_num
                self.ser.timeout = 0.2
                error_mess = self.ser.readlines() #extra crap
                self.list_results.append(["Chip Number",self.__chip_number, dict_codeL[x], error_mess])

                self.textframe.insert(tkinter.END,self.list_results[-1])
                self.textframe.insert(tkinter.END,"\n")
                
                
                self.ser.timeout = self.__poll_ser
                self.PASS = False
                self.ser.flushInput()
                self.ser.flushOutput()
            
            #lost sync or connection
            else: 
                self.list_results.append(["SYNC ERROR"])
                self.textframe.insert(tkinter.END,self.list_results[-1])
                self.textframe.insert(tkinter.END,"\n")
                
                self.__chip_number -= 1
                self.__display['text'] = str(self.__chip_number)
                self.PASS = False
                self.__stall_color()
                
            self.root.update()    #update GUI
            
###################################################################  
#print chip information to txt file
#shows all pass and fails of chip into a txt file

    def __print_to_file(self):
        f = open(self.filename,"w") #opens file with name of "test.txt"
        for x in self.list_results:
            f.write(str(x) + '\n')
        f.close()
        self.list_results = []
###################################################################
#LED lights on GUI frame
#Reset
#change all the lights to Yellow
    def __reset_color(self):    
        
        self.__widgets['COT'][1]['bg'] = 'yellow'
        #self.__widgets['COT'][1].update_idletasks()

        self.__widgets['CHE'][1]['bg'] = 'yellow'
        #self.__widgets['CHE'][1].update_idletasks()
     
        self.__widgets['ERT'][1]['bg'] = 'yellow'
        #self.__widgets['ERT'][1].update_idletasks()
        
        self.__widgets['WRT'][1]['bg'] = 'yellow'
        #self.__widgets['WRT'][1].update_idletasks()
         
        self.__widgets['WIRT'][1]['bg'] = 'yellow'
        #self.__widgets['WIRT'][1].update_idletasks()
  
        self.__widgets['WIT'][1]['bg'] = 'yellow'
        #self.__widgets['WIT'][1].update_idletasks()
        
        self.__widgets['WRAT'][1]['bg'] = 'yellow'
        #self.__widgets['WRAT'][1].update_idletasks()
       
        self.__widgets['DO'][1]['bg'] = 'yellow'
        #self.__widgets['DO'][1].update_idletasks()
        
        self.root.update()    #update GUI
        
#Stall
    def __stall_color(self):
	#if stall, change all the lights to purple
	
        #Play Sound
        try:
            self.root.update()    #update GUI, #this is necessary to have
            winsound.PlaySound('air_horn.wav',winsound.SND_FILENAME) #there is an issue 
            #'C:/Users/Tekmos/Documents/IPython Notebooks/Interface/music.wav'
        except:
            self.textframe.insert(tkinter.END,"sound error")
        
        self.__widgets['COT'][1]['bg'] = 'purple'
        #self.__widgets['COT'][1].update_idletasks()
        
        self.__widgets['CHE'][1]['bg'] = 'purple'
        #self.__widgets['CHE'][1].update_idletasks()
        
        self.__widgets['ERT'][1]['bg'] = 'purple'
        #self.__widgets['ERT'][1].update_idletasks()
        
        self.__widgets['WRT'][1]['bg'] = 'purple'
        #self.__widgets['WRT'][1].update_idletasks()
        
        self.__widgets['WIRT'][1]['bg'] = 'purple'
        #self.__widgets['WIRT'][1].update_idletasks()
        
        self.__widgets['WIT'][1]['bg'] = 'purple'
        #self.__widgets['WIT'][1].update_idletasks()
        
        self.__widgets['WRAT'][1]['bg'] = 'purple'
        #self.__widgets['WRAT'][1].update_idletasks()
        
        self.__widgets['DO'][1]['bg'] = 'purple'
        #self.__widgets['DO'][1].update_idletasks()
        
        self.root.update()    #update GUI
####################################################################       
#Messages

#WRONG PORT OR NEEDS RECONNECTION
    def Wrong_port(self):
        tkMessageBox.askokcancel("COM PORT", "Wrong Port")     
#About Program    
    def about(self):
        tkMessageBox.showinfo("About","Author: Vema Reddy, Dat Ho, Michael Robinson 2015 FOR TK7602A")
#CLOSE PROGRAM
    def exit_app(self):
        if tkMessageBox.askokcancel("Quit", "Do you really want to quit?"):
            self.root.destroy()         
			
#####################################################################
#Start Program
if __name__ == '__main__':
    main()