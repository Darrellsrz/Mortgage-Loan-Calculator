# Name : Siow Run Ze
# Kelas : 3 Red
# Guru : Pn.Lai
#import module and python features from system shell or command prompt
import tkinter 
from tkinter import*
from tkinter.ttk import*
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import time
import webbrowser
from PIL import Image
from pathlib import Path 
exportclick=True
firstclick=True
# date and time
def search(path:str)-> Path : 
    return Path("./assets")/Path(path)
def clock():
        hour=time.strftime("%I")
        minute=time.strftime("%M")
        second=time.strftime("%S")
        day = time.strftime("%A")
        month=time.strftime("%b")
        am_pm = time.strftime("%p")
        time_zone=time.strftime("%Z")
        digital.config(text="   "+hour + ":" + minute + ":" + second+" "+am_pm+"\n"+month+" "+day )
        digital.after(1000,clock)


def checkresult():
    global LoanUI , DpUI,InterestUI,yearUI,MR,scrollbar
    Loan=str(LoanUI.get())
    try :
       Loan=float(Loan)
       if Loan<=0 :
            messagebox.showerror('Invalid input',"No negatif value or Zero Loan Amount")
            LoanUI.set(" ")
            Loanpy.state(["invalid"])
            print(Loanpy.state())
            
            
       elif Loan >=1000000000 :
            messagebox.showerror('Invalid input',"Please input the Loan less than RM 100 Million")
            LoanUI.set(" ")
            Loanpy.state(["invalid"])
            print(Loanpy.state())
            
    except ValueError:    
            messagebox.showerror('Invalid Input',"Please input only number for Loan Amount")
            LoanUI.set(" ")
            Loanpy.state(["invalid"])
            print(Loanpy.state())
            
            
    Dp=str(DpUI.get())
    try :
        Dp=float(Dp)
        if Dp<=0 :
            messagebox.showerror('Invalid input',"No negatif value or Zero Down Payment")
            DpUI.set(" ")
            Dppy.state(["invalid"])
            print(Dppy.state())
            
        elif Dp>=100 :
            messagebox.showerror('Invalid input',"Please input the down payment less than 100%")
            DpUI.set(" ")
            Dppy.state(["invalid"])
            print(Dppy.state())
            
    except ValueError:    
            messagebox.showerror('Invalid Input',"Please input only number for Down Payment ")
            DpUI.set(" ")
            Dppy.state(["invalid"])
            print(Dppy.state())
            
    Interest=str(InterestUI.get())
    try :
        Interest=float(Interest)
        if Interest<=0 :
            messagebox.showerror('Invalid input',"No negatif value or Zero Interest Rate")
            InterestUI.set(" ")
            Interestpy.state(["invalid"])
            print(Interestpy.state())

        elif Interest>=100 :
            messagebox.showerror('Invalid input',"Please input the interest rate less than 100%")
            InterestUI.set(" ")
            Interestpy.state(["invalid"])
            print(Interestpy.state())
            
    except ValueError:    
            messagebox.showerror('Invalid Input',"Please input only number for Interest Rate")
            InterestUI.set(" ")
            Interestpy.state(["invalid"])
            print(Interestpy.state())
            
    year=str(yearUI.get())
    try :
        year=float(year)
        if year<=0 :
            messagebox.showerror('Invalid input',"No negatif value or Zero Year ")
            yearUI.set(" ")
            Yearpy.state(["invalid"])
            print(Yearpy.state())
        
        elif year >= 200 :
            messagebox.showerror('Invalid Input',"Please input the year below 200 years to avoid error")
            yearUI.set(" ")
            Yearpy.state(["invalid"])
            print(Yearpy.state())
            
    except ValueError :
        messagebox.showerror('Invalid Input',"Please input only number for year")
        yearUI.set(" ")
        Yearpy.state(["invalid"])
        print(Yearpy.state())
            
   
def button_clicked():
    global LoanUI,DpUI,InterestUI,yearUI,MR,scrollbar,checkresult,Period_list,Principle_list,Interest_list,Balance_list,total_loan,button_export,total_int,firstclick,exportclick
    
    checkresult()
    Loan=str(LoanUI.get())
    Dp=str(DpUI.get())
    Interest=str(InterestUI.get())
    year=str(yearUI.get())
    
    Period_list = []
    Principle_list=[]
    Interest_list=[]
    Balance_list=[]
    Loan=float(Loan)
    x=((1+float(Interest)/100/12)**(-12*float(year)))
    Loan=float(Loan)*((100-float(Dp))/100)
    mr=((float(Loan)*(float(Interest)/100/12))/(1-float(x)))
    total_loan=float(Loan)
    total_m=float(year)*12
    
    Label(root,text=format(mr,".2f")).place(x=130,y=225)
    # print out the analysis
    for a in range(int(total_m)):
        Period_list.append(a+1)
        monthly_int=float(Loan)*(float(Interest)/12/100)
        Interest_list.append(format(monthly_int,".2f"))
        principle=float(mr)-monthly_int
        Principle_list.append(format(principle,".2f"))
        Loan=float(Loan)-principle
        if Loan<0 :
            Loan=0
        Balance_list.append(format(Loan,".2f"))
        if a % 2 == 0 :
            tree.insert(parent = '', index = 'end', iid = a, text = "parent",
                     values = (Period_list[a], Principle_list[a], Interest_list[a], Balance_list[a]),tags=("even",))
        else :
            tree.insert(parent = '', index = 'end', iid = a, text = "parent",
                     values = (Period_list[a], Principle_list[a], Interest_list[a], Balance_list[a]),tags=("odd",))
        tree.tag_configure("even",foreground="black",background="#D6EAF8")
        tree.tag_configure("odd",foreground="black",background="#EAFAF1")

    total_int=(float(mr)*(12*float(year)))-total_loan
    tree.insert(values=('Total',format(total_loan,"0.2f"),format(total_int,"0.2f"),' '), tags=("total"),parent='', index='end', iid=a+1, text="")
    tree.tag_configure("total",foreground="black",background="#839192")
    
    # disable the button
    Loanpy.configure(state=DISABLED)
    Dppy.configure(state=DISABLED)
    Interestpy.configure(state=DISABLED)
    Yearpy.configure(state=DISABLED)
    firstclick=False
    exportclick=True
    # show the export button
    button_export.place(x=635,y=490,width=156,height=121)

def function_export() :
    global Period_list,Principle_list,Interest_list,Balance_list,total_loan,total_int
    Period_list.append("      Total :")
    Principle_list.append(format(total_loan,".2f"))
    Interest_list.append(format(total_int,".2f"))
    Balance_list.append(" ")
    excel = pd.DataFrame({
            'Period': Period_list,
            'Principle': Principle_list,
            'Interest': Interest_list,
            'Balance': Balance_list
    })
    filename = filedialog.asksaveasfilename()
    excel.to_excel(filename + ".xlsx", index=False)
    reset()
    exportclick = False

def reset() :
    global reset
    #reset the functionality of the button 
    Loanpy.configure(state=NORMAL)
    Dppy.configure(state=NORMAL)
    Interestpy.configure(state=NORMAL)
    Yearpy.configure(state=NORMAL)

    Loanpy.delete(0,END)
    Dppy.delete(0,END)
    Interestpy.delete(0,END)
    Yearpy.delete(0,END)
    for i in tree.get_children():
        tree.delete(i)
    Mr = Label(root,text="             ")
    Mr.place(x=130,y=225)
    button_export.place_forget()
    exportclick=True

def on_closing():
    global button_export,firstclick,exportclick 
    if firstclick == True or exportclick == True :
        root.after_cancel(digital)
        root.destroy()
    else :  
        quit=messagebox.askyesnocancel("Quit","Do you want to generate the excel file before quit ?") 
        if quit == True :
            function_export()
            root.destroy()
        elif quit == False :
            digital.destroy()
            root.destroy()
##MAIN PROGRAM

#basic set for the user interface
root=Tk()
root.title("Mortgage Loan Calculator")
root.geometry("920x625")
root.maxsize (920,625)
root.minsize (920,625)
root.tk.call('wm', 'iconbitmap', root._w, 'home.ico')
root.iconbitmap("home.ico")

# theme for gui
style = ttk.Style(root)
root.tk.call('source', 'theme/azure dark/azure dark.tcl')
style.theme_use('azure')
    
digital = Label(root,text=" ",font=("DS-Digital",48) , foreground ="white",background="#333333")
digital.place (x=535,y=10)
clock()

#input
Label(root,text="Home Loan Calculator",font=("Showcard Gothic",25),foreground="#5499C7").place(x=10,y=5)
Label(root,text="Loan Amount (RM) :").place(x=10,y=55)
Label(root,text="Down Payment (%) :").place(x=10,y=95)
Label(root,text="Interest Rate (%)      :").place(x=10,y=135)
Label(root,text="Year                           :").place(x=10,y=180)
Label(root,text="Monthly Repayment :").place(x=10,y=225)

#entry 
LoanUI=StringVar()
DpUI=StringVar()
InterestUI=StringVar()
yearUI=StringVar()

Loanpy = Entry(root,textvariable=LoanUI)
Dppy = Entry(root,textvariable=DpUI)
Interestpy = Entry(root,textvariable=InterestUI)
Yearpy=Entry(root,textvariable=yearUI)

Loanpy.place(x=130,y=50)
Dppy.place(x=130,y=90)
Interestpy.place(x=130,y=130)
Yearpy.place(x=130,y=170)

#style
style = ttk.Style()
style.configure("bordercancel.TLabel", background="#E7B88D",borderwidth= 0)
#button
imgexport = PhotoImage(file=search("Excel1.png"))

button_export=Button(root,image=imgexport,style="bordercancel.TLabel",command=function_export)
button_export.place_forget()

imgFrame = Frame(root, width=350, height=350)
img1 = PhotoImage(file=search("HLimg3.png"))
imgFrame.place(x=550, y=125)
lbl = Label(imgFrame, image=img1)
lbl.place(x=0,y=0)

imghelp=PhotoImage(file=("web3.png"))
btnhelp=Button(root,image=imghelp,style="bordercancel.TLabel")
btnhelp.place(x=430,y=10)
btnhelp.bind("<Button-1>", lambda e:webbrowser.open_new_tab('https://github.com/Darrellsrz/Mortgage-Loan-Calculator'))

btncalculate = PhotoImage(file=search('Calculate_button.png'))
button_calculate=Button(root,image=btncalculate,command=button_clicked,style="bordercancel.TLabel")
button_calculate.place(x=300,y=170)

button_reset=Button(root,text="Reset",style='Accentbutton',command=reset)
button_reset.place(x=300,y=110)

button_exit=Button(root,text="Exit",style='Accentbutton',command=on_closing)
button_exit.place(x=300,y=70)


#style for the Treeview
style=Style()
style.configure("Treeview",background="#D3D3D3",foreground="black",fieldbackground = "#D3D3D3" ,rowheight=25)
style.map("Treeview", background=[('selected', ' #')])

#treeview for result
tree_frame = Frame(root)
tree_frame.place(x = 10, y= 270,width=500,height=340)

tree_scroll=Scrollbar(tree_frame,orient=VERTICAL)
tree_scroll.pack(side=RIGHT,fill=Y)

tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set,selectmode="extended")
tree['columns'] = ("Period", "Principle", "Interest", "Balance")
tree.config(height=300)
tree.place(x=0,y=0,width=485,height=340)
# columns
tree.column("#0",width=0, stretch=NO)
tree.column("Period", anchor=CENTER,width=100,minwidth=100)
tree.column("Principle", anchor=CENTER, width = 100,minwidth=100)
tree.column("Interest", anchor=CENTER, width=100,minwidth=100)
tree.column("Balance", anchor=CENTER, width=100,minwidth=100)
# headings
tree.heading("#0")
tree.heading("Period", text = "Period ", anchor = CENTER)
tree.heading("Principle", text = "Principle ", anchor = CENTER)
tree.heading("Interest", text = "Interest ", anchor = CENTER)
tree.heading("Balance", text = "Balance ", anchor = CENTER)
# Configure Scrollbar
tree_scroll.config(command=tree.yview)
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()




            
 
    

  
