#from doc_to_pdf import *
#si lo descargas porfavor complementalo y luego subelo, usualo para fines didacticos.
#gracias
from Tkinter import *
import tkFileDialog as dialog
import tkMessageBox
from os import chdir,  path
from time import strftime
from win32com import client



#linie de functions
def message_box(title,message):
	tkMessageBox.showinfo(title,message)

def capture_file():
	#I catch the string with 
	name_doc=''
	file = dialog.askopenfilename()
	if len(file)>0 :
		
		#here become below the string in a list of python
		listl=file.split('/')	
		lenl = len(listl)
		l = lenl-1
		name_doc = listl[l]
		listl[l]="" 
		
		strl = '/'.join(str(e) for e in listl)
		#print strl
		#chdir(strl)
		folder= strl
        c=0

        try:
           word = client.DispatchEx("Word.Application")
           if  name_doc.endswith(".docx"):# determinar si termina en .docx e archivo

               new_name = name_doc.replace(".docx", r".pdf")#rempaza la extencion archivo
               in_file = path.abspath(folder + "\\" + name_doc)
               new_file = path.abspath(folder + "\\" + new_name)
               #print new_file
               
               doc = word.Documents.Open(in_file)
               print strftime("%H:%M:%S"), " docx -> pdf ", path.relpath(new_file)
               doc.SaveAs(new_file, FileFormat = 17)
               doc.Close()
               c=1
               messge_doc =  path.relpath(new_file)
           if name_doc.endswith(".doc"):

               new_name = name_doc.replace(".doc", r".pdf")
               in_file = path.abspath(folder + "\\" + name_doc)
               new_file = path.abspath(folder + "\\" + new_name)
               doc = word.Documents.Open(in_file)
               #print strftime("%H:%M:%S"), "El archivo doc  -> pdf ", path.relpath(new_file)
               doc.SaveAs(new_file, FileFormat = 17)
               doc.Close()
               c=1
               messge_doc =  path.relpath(new_file)
        except Exception, e:
             print e
        finally:
             word.Quit()
          

        

        if c==1:
           m = 'Conexion establecida. Proceso finalizado, el documento '+name_doc+' ha sido convertido a '+messge_doc+'.'
           message_box('Mensaje',m)
           label_message= Label(form,text= messge_doc)
           label_message.grid(row=2,column=0,padx=40)
           label_message.pack()
           label_message= Label(form,text= 'Ruta :'+strl)
           label_message.grid(row=2,column=0,padx=40)
           label_message.pack()

        	
        else:
            message_box('Mensaje','Conexion serrada')

            
           
       
        
base = Tk()
base.title('CONVERTIDOR .DOC(X) A .PDF')
base.geometry("500x100+500+0")
form = Frame(base,width=500,height=500)
form.pack()

label= Label(form,text= "Cheking the file a convert")
label.grid(row=0,column=0,padx=40)
label.pack()
button=Button(form, text= "Click me",command=capture_file)
button.grid(row=1,column=0,padx=0)
button.pack()


base.mainloop()