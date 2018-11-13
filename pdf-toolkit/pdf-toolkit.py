import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from pathlib import Path as P
import re
from pikepdf import Pdf
import pikepdf
import tempfile
import shutil
import os

def tab_change(*args):
    tab=n.index('current')
    input_label.config(state=('disabled' if tab==1 else 'enabled'))
    input_entry.config(state=('disabled' if tab==1 else 'enabled'))
    input_button.config(state=('disabled' if tab==1 else 'enabled'))
def go_cmd(*args):
    tab=n.index('current')
    if tab!=1:
        infile=input_file.get()
        if infile=='':
            messagebox.showerror(title='Input Error!',message='Must specify an input file!')
            return
        elif not P(infile).exists():
            messagebox.showerror(title='Input Error!',message='Input file does not exist!')
            return
    outfile=output_file.get()
    if outfile=='':
        outfile=filedialog.asksaveasfilename(title='Select output file...',filetypes=(('PDF Files','*.pdf'),('All files','*.*')))
        if outfile!='':
            output_file.set(outfile)
        else:
            messagebox.showerror(title='Output Error!',message='Must specify output file!')
            return
    if tab==0: #Rotate
        fp,tp=int(from_page.get())-1,int(to_page.get())-1
        if tp<fp:
            tp=fp
            to_page.set(str(fp+1))
        d=int(direction.get())
        rotate(infile,fp,tp,d,outfile)
    elif tab==1: #Merge
        merge(filelist,outfile)
    elif tab==2: #Reorder
        pl=pagelist.get()
        match=re.findall(r'\d+\-\d+|\d+',pl)
        if pl.strip()!='' and len(match):
            pl=[]
            for pr in match:
                if pr.find('-')>0:
                    t=tuple(map(int,pr.split('-')))
                else:
                    t=int(pr)
                pl.append(t)
        elif splitval.get():
            pl=[(from_entry.config()['from'][4],to_entry.config()['to'][4])]
        else:
            messagebox.showerror(title='Input Error!',message='No valid page range found!')
        reorder(infile,pl,outfile)
    return
def rotate(infile,fp,tp,d,outfile):
    global password
    infile_stream=open(infile,'rb')
    try:
        pdf_in=Pdf.open(infile_stream)
    except pikepdf.PasswordError:
        if not password:
            password=get_password(P(infile).stem)
        pdf_in=Pdf.open(infile_stream,password)
    if outfile!=infile:
        pdf_out=Pdf.new()
        pdf_out.pages.extend(pdf_in.pages)
    else:
        pdf_out=pdf_in
    for page in pdf_out.pages[fp:tp+1]:
        page.Rotate=(page.Rotate+d)%360
    pdf_out.save(outfile)
    infile_stream.close()
    del pdf_out
def merge(filelist,outfile):
    if outfile in filelist:
        messagebox.showerror(title='Filename Error',message='Input and output filenames must be different for this operation.')
        return
    pdf_out=Pdf.new()
    for f in filelist:
        try:
            pdf=Pdf.open(f)
        except pikepdf.PasswordError:
            password=get_password(P(f).stem)
            pdf=Pdf.open(f,password)
        pdf_out.pages.extend(pdf.pages)
    pdf_out.save(outfile)
    del pdf_out
def reorder(infile,pl,outfile):
    global password
    infile_stream=open(infile,'rb')
    try:
        pdf_in=Pdf.open(infile_stream)
    except pikepdf.PasswordError:
        if not password:
            password=get_password(P(infile).stem)
        pdf_in=Pdf.open(infile_stream,password)
    if splitval.get():
        max_pagenum=0
        for i in pl:
            if isinstance(i,int) and i>max_pagenum:
                max_pagenum=i
            elif isinstance(i,tuple) and max(i)>max_pagenum:
                max_pagenum=max(i)
        path=P(outfile)
        outfiletemplate=str(P(path.parent,path.stem))+'_{:0'+str(len(str(52)))+'d}.pdf'
        for i in pl:
            if isinstance(i,int):
                pdf_out=Pdf.new()
                pdf_out.pages.extend([pdf_in.pages.p(i)])
                pdf_out.save(outfiletemplate.format(i))
            elif isinstance(i,tuple):
                for j in range(i[0],i[1]+1):
                    pdf_out=Pdf.new()
                    pdf_out.pages.extend([pdf_in.pages.p(j)])
                    pdf_out.save(outfiletemplate.format(j))
            del pdf_out
    else:
        pdf_out=Pdf.new()
        for i in pl:
            if isinstance(i,int):
                pdf_out.pages.extend([pdf_in.pages.p(i)])
            elif isinstance(i,tuple):
                pdf_out.pages.extend(pdf_in.pages[(i[0]-1):i[1]])
        pdf_out.save(outfile)
        del pdf_out
    infile_stream.close()
def get_input():
    global password
    infile=filedialog.askopenfilename(title='Select input file...',filetypes=(('PDF Files','*.pdf'),('All files','*.*')))
    input_file.set(infile)
    infile_stream=open(infile,'r+b')
    try:
        pdf_in=Pdf.open(infile_stream)
    except pikepdf.PasswordError:
        password=get_password(P(infile).stem)
        pdf_in=Pdf.open(infile_stream,password=password)
    from_entry.config(from_=1)
    from_entry.set('1')
    from_entry.config(to=len(pdf_in.pages))
    to_entry.config(to=len(pdf_in.pages))
    infile_stream.close()
def get_output():
    outfile=filedialog.asksaveasfilename(title='Select output file...',filetypes=(('PDF Files','*.pdf'),('All files','*.*')))
    if outfile[-4:]!='.pdf':
        outfile+='.pdf'
    output_file.set(outfile)
def add_merge_file():
    infile=filedialog.askopenfilenames(title='Select input file...',filetypes=(('PDF Files','*.pdf'),('All files','*.*')))
    for f in infile:
        path=P(f)
        if f!='' and path.exists() and path.suffix=='.pdf':
            filelist.append(f)
            filelistbox.insert('end',path.stem)
def del_merge_file():
    index=filelistbox.index('active')
    filelist.remove(filelist[index])
    filelistbox.delete(index)
def move_mergefile_up():
    index=filelistbox.index('active')
    if index==0:
        return
    filelistbox.delete(index)
    filelistbox.insert(index-1,P(filelist[index]).stem)
    filelist[index],filelist[index-1]=filelist[index-1],filelist[index]
    filelistbox.selection_anchor(index-1)
def move_mergefile_down():
    index=filelistbox.index('active')
    if index==filelistbox.index('end'):
        return
    filelistbox.delete(index)
    filelistbox.insert(index+1,P(filelist[index]).stem)
    filelist[index],filelist[index+1]=filelist[index+1],filelist[index]
    filelistbox.selection_anchor(index+1)
def get_password(filename):
    return simpledialog.askstring('Password','Enter password for {}'.format(filename),show='*')

root=tk.Tk()
root.title('PDF Tools')

#root.option_add('*tearOff',False)
#menubar=tk.Menu(root)
#root['menu']=menubar
#menu_file=tk.Menu(menubar)
#menubar.add_cascade(menu=menu_file,label='File')

#create and configure main window
mainframe=ttk.Frame(root)
frame=ttk.Frame(mainframe,borderwidth=5,padding=(10,10,10,10))
mainframe.grid(column=0,row=0,columnspan=3,rowspan=4,sticky=(tk.N,tk.E,tk.S,tk.W))
mainframe.columnconfigure(1,weight=1)
mainframe.rowconfigure(1,weight=1)
root.columnconfigure(0,weight=1)
root.rowconfigure(0,weight=1)

#define variables
input_file=tk.StringVar()
output_file=tk.StringVar()

#create input file widgets
input_label=ttk.Label(mainframe,text='Input')
input_label.grid(column=0,row=0,sticky=tk.E,padx=10,pady=10)
input_entry=ttk.Entry(mainframe,textvariable=input_file,state='disabled')
input_entry.grid(column=1,row=0,sticky=(tk.W,tk.E),pady=10)
password=''
input_button=ttk.Button(mainframe,text='Browse...',command=get_input)
input_button.grid(column=2,row=0,sticky=tk.W,padx=10,pady=10)

#create notebook and frames
n=ttk.Notebook(mainframe)
n.enable_traversal()
rotatepage=ttk.Frame(n,padding=(10,10,10,10))
mergepage=ttk.Frame(n,padding=(10,10,10,10))
reorderpage=ttk.Frame(n,padding=(10,10,10,10))

#configure rotate page
rotatepage.grid(column=0,row=0,columnspan=4,rowspan=4,sticky=(tk.N,tk.E,tk.S,tk.W))
rotatepage.columnconfigure(1,weight=1)
rotatepage.columnconfigure(3,weight=1)
rotatepage.rowconfigure(0,weight=1)
rotatepage.rowconfigure(1,weight=1)
rotatepage.rowconfigure(2,weight=1)
rotatepage.rowconfigure(3,weight=1)
#create rotate widgets
direction=tk.StringVar()
direction.set('270')
ttk.Label(rotatepage,text='Direction').grid(column=0,row=0,rowspan=3,sticky=tk.E,padx=10)
counterclockwise=ttk.Radiobutton(rotatepage,text='Counterclockwise',variable=direction,value='270',underline=3)
counterclockwise.grid(column=1,row=0,columnspan=3,sticky=(tk.E,tk.W))
clockwise=ttk.Radiobutton(rotatepage,text='Clockwise',variable=direction,value='90',underline=0)
clockwise.grid(column=1,row=1,columnspan=3,sticky=(tk.E,tk.W))
oneeighty=ttk.Radiobutton(rotatepage,text='Upside Down',variable=direction,value='180',underline=0)
oneeighty.grid(column=1,row=2,columnspan=3,sticky=(tk.E,tk.W))
ttk.Label(rotatepage,text='Range').grid(column=0,row=3,sticky=tk.E,padx=10)
from_page=tk.StringVar()
from_page.set('0')
to_page=tk.StringVar()
to_page.set('0')
from_entry=ttk.Spinbox(rotatepage,from_=0,textvariable=from_page,wrap=True)
from_entry.grid(column=1,row=3,sticky=(tk.E,tk.W))
ttk.Label(rotatepage,text='to').grid(column=2,row=3,sticky=(tk.E,tk.W),padx=10)
to_entry=ttk.Spinbox(rotatepage,from_=0,textvariable=to_page,wrap=True)
to_entry.grid(column=3,row=3,sticky=(tk.E,tk.W))

#configure merge page
mergepage.grid(column=0,row=0,columnspan=4,rowspan=4,sticky=(tk.N,tk.E,tk.S,tk.W))
mergepage.columnconfigure(1,weight=1)
mergepage.rowconfigure(0,weight=1)
mergepage.rowconfigure(1,weight=1)
mergepage.rowconfigure(2,weight=1)
mergepage.rowconfigure(3,weight=1)
#create merge widgets
ttk.Label(mergepage,text='Files').grid(column=0,row=0,rowspan=4,sticky=tk.E,padx=10)
filelist=[]
filelistbox=tk.Listbox(mergepage,height=10,listvariable=filelist)
filelistbox.grid(column=1,row=0,rowspan=4,sticky=(tk.N,tk.E,tk.S,tk.W))
s=ttk.Scrollbar(mergepage,orient=tk.VERTICAL,command=filelistbox.yview)
s.grid(column=2,row=0,rowspan=4,sticky=(tk.N,tk.S))
ttk.Button(mergepage,text='Add...',command=add_merge_file,underline=0).grid(column=3,row=0,sticky=(tk.N,tk.E,tk.W),padx=10)
ttk.Button(mergepage,text='Move Up',command=move_mergefile_up,underline=5).grid(column=3,row=1,sticky=(tk.E,tk.S,tk.W),padx=10)
ttk.Button(mergepage,text='Move Down',command=move_mergefile_down,underline=7).grid(column=3,row=2,sticky=(tk.N,tk.E,tk.W),padx=10)
ttk.Button(mergepage,text='Delete',command=del_merge_file,underline=0).grid(column=3,row=3,sticky=(tk.E,tk.S,tk.W),padx=10)

#configure reorder page
reorderpage.grid(column=0,row=0,columnspan=2,rowspan=2,sticky=(tk.N,tk.E,tk.S,tk.W))
reorderpage.columnconfigure(1,weight=1)
reorderpage.rowconfigure(0,weight=1)
reorderpage.rowconfigure(1,weight=1)
#create reorder widgets
ttk.Label(reorderpage,text='Page List').grid(column=0,row=0,sticky=tk.E,padx=10)
pagelist=tk.StringVar()
page_entry=ttk.Entry(reorderpage,textvariable=pagelist)
page_entry.grid(column=1,row=0,sticky=(tk.E,tk.W))
splitval=tk.IntVar()
splitval.set(0)
splitcheck=tk.Checkbutton(reorderpage,text='Split pages as separate files',underline=0,variable=splitval).grid(column=0,row=1,columnspan=2,sticky=(tk.E,tk.W))

#create tab interface
n.add(rotatepage,text='Rotate',underline=0)
n.add(mergepage,text='Merge',underline=0)
n.add(reorderpage,text='Reorder/Split',underline=2)
n.grid(column=0,columnspan=3,row=1,sticky=(tk.W,tk.E,tk.N,tk.S),padx=10)

#create output file widgets
ttk.Label(mainframe,text='Output').grid(column=0,row=2,sticky=tk.E,padx=10,pady=10)
output_entry=ttk.Entry(mainframe,textvariable=output_file)
output_entry.grid(column=1,row=2,sticky=(tk.W,tk.E),pady=10)
ttk.Button(mainframe,text='Browse...',command=get_output).grid(column=2,row=2,sticky=tk.W,padx=10,pady=10)

#create go widget
go=ttk.Button(mainframe,text='Go!',command=go_cmd)
go.grid(column=2,row=3,sticky=tk.E,padx=(0,10),pady=(0,10))

#set default focus
input_entry.focus()
#bind return key
root.bind('<Return>', go_cmd)
root.bind('<<NotebookTabChanged>>',tab_change)
#center window
root.withdraw()
root.update_idletasks()
x=(root.winfo_screenwidth()-root.winfo_reqwidth())/2
y=(root.winfo_screenheight()-root.winfo_reqheight())/2
root.geometry('+{:.0f}+{:.0f}'.format(x,y))
root.minsize(root.winfo_width(),root.winfo_height())
root.resizable(False,True)
root.deiconify()
#start application
root.mainloop()
