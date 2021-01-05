
#===================================================================================================
#IMPORTING GENERAL MODULES
#===================================================================================================

import os
import PySimpleGUI as sg
import sys #Used when creating .exe file
import shutil
from math import log

#===================================================================================================
#SETTING UP APP ICON AND USER MANUAL
#===================================================================================================

#We'll use Auto PY to EXE to convert the script into an .exe file so it can run over Windows.
#Now, things is, even thought we do want to include a taskbar icon and an user manual, we do not 
#want to generate a directory with all these depencies inside. Instead, we fancy an standalone 
#executable file. To that end, we make use of the following script,which is covered in more detail 
#over here: https://dev.to/eshleron/how-to-convert-py-to-exe-step-by-step-guide-3cfi

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

icon=resource_path('finalicon.ico')
manual=resource_path('Manual del usuario.pdf')

#===================================================================================================
#SETTING UP NOTIFICATIONS
#===================================================================================================

class notification(object):
    '''
    instantiates a notification class object. 
    Input: any message
    '''
    def __init__(self, done_message):
        self.done_message=done_message
        
    def notif(self): 
        '''
        Notififes the user when task is completed.
        '''
        sg.SystemTray.notify('¡LISTO!', self.done_message,
                                                 display_duration_in_ms=1500,
                                                 fade_in_duration=0.05)

    def notif_time(start,end):
        '''
        Notifies when task is done, adding how long it took to be completed.
        '''
        #If you intend to use it, add the following line at the very beginning:
        # from datetime import datetime as dt
        
        delta=end-start
        if delta.seconds==0:
            timer=delta.microseconds
            timer=round((timer/1000000),5)   
        else:
            timer=delta.seconds
        timer_format=' segundos'

        sg.SystemTray.notify('LISTO!', 'Tiempo de ejecución: '+str(timer)+timer_format,
                                                 display_duration_in_ms=1000,
                                                 fade_in_duration=0.1)
        return None
 
#===================================================================================================
#SETTING UP POPUP WINDOWS
#===================================================================================================
    
class popup(object):
    '''
    instantiates  a popup class object. 
    Input: any message
    '''
    def __init__(self, message):
        self.message=message
    def get_message(self,title):
        '''Initializes a popup window when an element from the Menu is clicked on.'''
        sg.Popup(self.message,font=font_type, title=title,
                              keep_on_top=True,background_color=bc,
                              text_color=tc,modal=True)
        return None
    def get_error_message(self):
        '''Initializes a popup window when an error is raised'''
        sg.Popup(self.message, title='OOPS!',
                             custom_text=('Regresar', None),
                             font=font_type,background_color=bc,text_color=tc)
        return None

#===================================================================================================
#SETTING UP CORE FUNCTIONS
#===================================================================================================
       
class work_with_inputs(object):
    '''
    Initializes a work_with_inputs instance.
    Input: absolute file path returned from "Save as".

    '''
    def __init__(self, SaveAs):
        self.SaveAs=SaveAs
        
    def word_pdf(self):
        '''
        Generates a blank PDF/A document, a key step to convert a PDF into PDF/A.
        
        Takes advantage of MS Word's "Save as PDF's make PDF/A compatible option". 
        A great deal is you don't have an Adobe Acrobat Pro subscription.' 
        '''
        
        from win32com import client
        from docx import Document
        
        DestDir=os.path.abspath(os.path.join(self.SaveAs, '..'))
        
        tempDocx='\\tempword.docx'
        documento = Document()
        documento.save(DestDir+tempDocx)
       
        word = client.DispatchEx('Word.Application')
        worddoc = word.Documents.Open(DestDir+tempDocx,ReadOnly = 1)
            
        worddoc.SaveAs(self.SaveAs,FileFormat = 17)
        
        worddoc.Close(True)
        word.Quit()
        os.remove(DestDir+tempDocx)
        return None
  
 
class work_with_file(work_with_inputs):  
    '''
    Initializes a work_with_inputs's subclass instance.
    Inputs: absolute file paths returned from "Save as" and "Browse".
    '''
    
    def __init__(self, SaveAs,Origin):
        work_with_inputs.__init__(self,SaveAs)
        self.Origin=Origin
            
    def PDF_PDFA(self):
        '''
        Turns a PDF into PDF/A.
        '''
        self.word_pdf()
        
        import fitz
        
        pdf=fitz.open(self.SaveAs)
        opened_file=fitz.open(self.Origin)
        pdf.insertPDF(opened_file)
        pdf.saveIncr()
        opened_file.close()
        pdf.deletePage(0)
        pdf.saveIncr()
        pdf.close()
        os.remove(self.Origin)
        return None
    
    def Req(self,name):
        '''
        Very useful at work.
        Creates a standard zip file required at work alongside a folder with its content. 
        A nice example of using the 're' module.
        '''   
        import re
        import fitz
        import shutil
        
        os.mkdir(self.SaveAs)
        os.mkdir(os.path.join(self.SaveAs,name))
        
        doc = fitz.open(self.Origin)
        page = doc.loadPage(0)              
        text = page.getText()  
        nro_req=re.findall('\s+\d{13}|°+\d{13}|º+\d{13}', text)
        nro_ruc=re.findall('\s+\1d{10}|°+\1d{10}|º+\1d{10}|\s+2\d{10}|°+2\d{10}|º+2\d{10}', text)  
        doc.close()
        nro_ruc = nro_ruc[0].replace("°","").replace("º","").replace(" ","").strip()
        nro_req = nro_req[0].replace("°","").replace("º","").replace(" ","").strip()
        
        text_file_path=self.SaveAs+'\\files.txt'
        text_file = open(text_file_path, 'w')
        text_file.write(nro_ruc+'|'+nro_req)
        text_file.close()
        
        pdf_element=self.SaveAs+'\\'+name+'\\'+nro_ruc+'_'+nro_req+'.pdf'
        text_element=self.SaveAs+'\\'+name+'\\files.txt'
        
        shutil.copy(self.Origin,pdf_element)
        shutil.move(text_file_path,text_element)
   
        shutil.make_archive(self.SaveAs+'\\'+name, 'zip',root_dir=self.SaveAs , base_dir=name)
  
        return None
    
    def PDF_same_size(self):
        '''
        Have you all of your PDF's pages adopt vertical A4 dimensions.
        Best part: gets the job done without messing up its contents.
        Credits go to user "JorjMcKie" from the PyMuPDF team on GitHub.
        '''
        import fitz
        
        src= fitz.open(self.Origin)
        doc = fitz.open()
        for ipage in src:
            #Some pages, even though having a landscape aspect, are not actually 'landscape', 
            #just rotated.
            #These lines of code take care of the many problems that arise when dealing with 
            #different page sizes. 
            if ipage.rotation==90: 
                ipage.setRotation(0)
            fmt = fitz.PaperRect("a4")
            page = doc.newPage(width = fmt.width, height = fmt.height)
            page.showPDFpage(page.rect, src, ipage.number)
        src.close() 
        doc.save(self.Origin)
        
        doc.close()
        shutil.move(self.Origin,self.SaveAs)
        return None
    
class work_with_documents(work_with_inputs):
    '''
    Initializes a work_with_inputs's subclass instance.
    Inputs: absolute file path returned from "Save as" and a list that takes 
    in all absolute file paths returned from "Browse" as elements.
    '''
    def __init__(self,SaveAs,documents):
        work_with_inputs.__init__(self,SaveAs)
        self.documents=documents
    
    def PDF_merger(self):
        '''
        Merges in all PDF files selected by user.
        '''
        import fitz

        pdf=fitz.open()
        for element in self.documents:          
            opened_file=fitz.open(element)        
            pdf.insertPDF(opened_file)
            opened_file.close()
      
        pdf.save(self.SaveAs)
        pdf.close()
       
        return None

#===================================================================================================
#USEFUL DEFINITIONS
#===================================================================================================

default_name='sehr witzig' #default output/folder name

description='Versión 1.0' #Version number
name='LuftMensch' #App title
bottom_description='Una aplicación de productividad' #Short description
note='''Nota: Antes de darle click en "Ejecutar" a cualquiera de las opciones, asegúrate de cerrar los archivos 
PDF con los que vayas a trabajar. La aplicación te notificará una vez que tu archivo se encuentre listo 
para ser visualizado.'''  #Warning-style caption

#Error messages
error1='No se encontró el número de Requerimiento y/o RUC.'
error2='Algo salió mal; asegúrate de que los directorios de origen y destino sean válidos.'
error3=''
error4='Si lo que deseas es unir PDFs, debes seleccionar más de un archivo.'
error5='Ya existe una carpeta con ese nombre. Por favor Ingresa otro.'
error6='Debes ingresar un número.'
error7='Debes ingresar un nombre para tu archivo.'
error8='Debes seleccionar un archivo para trabajar.'
done_message='Ya puedes visualizar tu documento.'

Tip='''
Si deseas unir solo unos cuantos PDFs que se encuentran en la carpeta seleccionada y no todos, haz lo siguiente:
 
Durante la selección, crea una nueva carpeta, copia o arrastra hasta allí los PDFs a unir y listo. 

Adicionalmente, si deseas unirlos en un orden específico, ordénalos dentro de la carpeta creada, selecciona el primer archivo y dale CTRL+E.

Si estás interesado en conocer más sobre esta opción, revisa el manual del usuario.
'''

#===================================================================================================
#OPTIONS
#===================================================================================================

abouts=['','','','',''] #Options
abouts[0]='''
Convierte un archivo PDF en formato PDF/A.

'''
abouts[1]='''Une todos los PDFs que se encuentran en la carpeta seleccionada en un único archivo PDF.

El resultado es un archivo PDF ideal para lectura e impresión.

Usa esta opción si tienes muchos archivos para unir y todos se encuentran en la misma carpeta.

Si quieres unir tus PDFs en un ORDEN específico (por nombre, por ejemplo), 
sigue los siguientes pasos para seleccionar correctamente tus archivos:

  1) Dale click en "Buscar".
        
  2) Asegúrate de que tus PDFs se encuentren en el orden deseado.
    
    2.1) Si no lo están, puedes ordenarlos allí mismo.
    2.2) Ten presente que los PDFs se unirán tal como aparecen en tu pantalla.
    
  3) Selecciona el primer archivo y dale "CTRL + E".
    
    3.1) No importa si en la selección se incluyeron carpetas; no serán tomadas en cuenta.
    3.2) Es muy IMPORTANTE que selecciones el primer archivo de la lista antes de darle "CTRL + E".
    3.3) De no hacerlo, los archivos se unirán en desorden.
    
  4) Dale click en abrir y listo. '''
abouts[2]='''Logra que todas páginas del PDF seleccionado posean las mismas dimensiones (A4 Vertical), 
sin alterar su contenido ni comprometer su correcta visualización. 

Termina por convertir el PDF resultante en PDF/A haciendo uso de la primera opción y 
el archivo que obtengas será ideal para que lo adjuntes al SIEV o lo envíes por SINE, 
ya que elimina el riesgo de que al subirse se recorten algunas páginas. '''
abouts[3]='''Genera el archivo modelo (.zip) de un Requerimiento o Resultado de 
Requerimiento con el contenido apropiado, listo para enviarse por SINE. 

Adicionalmente, se generará una carpeta descomprimida de dicho archivo 
para que puedas visualizar su contenido sin necesidad de descomprimir el fichero.

Solo escoge tu documento PDF, selecciona la carpeta donde quieres que se almacene
el archivo .zip e ingresa un nombre.

Si no ingresas una carpeta de destino, el archivo resultante se irá por defecto
a la carpeta que contenga el documento seleccionado.'''

abouts[4]='''Une todos los PDFs que selecciones en un único archivo PDF, en el orden que escojas.

El archivo PDF resultante es ideal para lectura e impresión.

Usa esta opción si tienes pocos PDFs para unir, se encuentran en distintas 
carpetas o simplemente prefieres seleccionarlos uno por uno.

Si seleccionas un número menor de PDFs al número ingresado (seleccionaste dos, pero en la pantalla 
previa ingresaste seis, por ejemplo), no te preocupes; puedes dejar los demás espacios en blanco.'''
   
choices=['Convertir PDF en PDF/A',
                 'Unir todos los PDFs dentro de una carpeta',
                 'Lograr que todas las páginas de un PDF posean el mismo tamaño',
                 'Crear archivo .zip de Requerimiento',
                 'Unir varios PDFs seleccionándolos uno por uno'] 

#===================================================================================================
#MENU
#===================================================================================================

menu=['Acerca de','Manual del usuario','Visita el repositorio'] 

first='''
Luftmensch es una aplicación de productividad y de código abierto pensada en automatizar ciertas tareas administrativas.

Lima, 2021.
'''  
second='''
Manual del usuario.
'''  

third='''
Si tienes dudas, comentarios o sugerencias, visita el repositorio en la siguiente dirección: 
    
https://github.com/lheredias/Luftmensch     
'''  

descriptions=[first,second,third]

#Menu disctionary with names and descriptions
menu_dict={menu[0]:first,menu[1]:second,menu[2]:third}

#Menu layout for GUI
menu_def = [['Menu',[menu[0],menu[1],menu[2]]]]   

#===================================================================================================
#GRAPHICAL USER INTERFACE: SETTING UP THE BASIC STUFF
#===================================================================================================

font_type=('Helvetica', 12) #Layout font type and size

size1=(72,10) #Option description layout size
size2=(750,500) #GUI main window layout size
first_size=(750,3)
second_size=(750,8)
third_size=(750,6)
fourth_size=(750,11)
fifth_size=(750,9)

#Setting up some colors
bc='#84312E'       
tc='#E7E7E7'
bc2='#FF6F6C'
tc2= '#39569D' 
line_color='#5A5656'

#Setting up the GUI layout theme
sg.LOOK_AND_FEEL_TABLE['TUNED']={'BACKGROUND': '#F5F5F5',
 'TEXT': tc2,
 'INPUT': bc,
 'TEXT_INPUT': tc,
 'SCROLL': '#a5a4a4',
 'BUTTON': (tc,'#223C7C'),
 'PROGRESS': ('#0079d3', '#dae0e6'),
 'BORDER': 1,
 'SLIDER_DEPTH': 0,
 'PROGRESS_DEPTH': 0,
 'ACCENT1': '#ff5414',
 'ACCENT2': '#ff5414',
 'ACCENT3': '#ff5414'}


#===================================================================================================
#FULL STACK
#===================================================================================================
                           
def visual():
    '''
    Initializes the GUI, comprised by one main window with five available options which, 
    when clicked upon, takes the user through five different windows based on choice made.
    The function takes place inside a loop so the user can go back and forth through 
    the many available options. 
    
    ''' 
    done=False
    while not done:

        sg.theme('TUNED')
        #===========================================================================================
        #MAIN WINDOW
        #===========================================================================================
        layout = [[sg.Menu(menu_def,key='MENU', tearoff=False,)],
                  [sg.T("")],[sg.Text(name[:4],font=('Gotham Bold', 45),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 45),text_color=bc),
                              sg.Text(description,font=('Gotham', 12),text_color=line_color)],   
                  [sg.Text(bottom_description,font=('Gotham Bold', 11),text_color=line_color)],
                  [sg.T("")],
                  [sg.Text('='*250,text_color=line_color)],   
                  [sg.Text("Escoge una opción: "),
                   sg.OptionMenu(tooltip='OPCIONES', values=choices,key='choice',size=(90,1))],
                  [sg.Text('='*250,text_color=line_color)],
                  [sg.Submit('CONTINUAR'), sg.Cancel('SALIR')],
                  [sg.Text('='*250,text_color=line_color)],
                  [sg.Text(note,size=first_size,text_color=bc)]]
        
        window = sg.Window(name, layout, size=size2,alpha_channel=.95,disable_minimize=False,
                           font=font_type,resizable=True,icon=icon)
    
        finished=False
        while not finished:
            event, values2 = window.read()
            if event==menu[0]:
                popup(menu_dict[menu[0]]).get_message(menu[0])
            if event==menu[1]:
                import subprocess
                subprocess.Popen([manual],shell=True)          
            if event==menu[2]:
                popup(menu_dict[menu[2]]).get_message(menu[2])
            if event=='SALIR' or event is None:
                finished=True
                done=True
                window.close()
            if event=='CONTINUAR':
                finished=True
                window.close()    
                #===================================================================================
                #OPTIONS NUMBER 1 & 3 WINDOWS
                #===================================================================================
                if values2['choice'] == choices[0] or values2['choice'] == choices[2]:     
                    if values2['choice'] == choices[0]:
                        about=abouts[0]
                        window_size=first_size                   
                    else:
                        about=abouts[2]
                        window_size=third_size
                        
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text('='*250,text_color=line_color)],
                        [sg.Text(about,size=window_size)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Escoge un archivo: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FileBrowse(tooltip='SELECCIONA UN ARCHIVO PDF',
                                             button_text='Buscar', key='Browse',
                                             file_types=(("PDF", "*.pdf"),))],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(tooltip='SOLO PODRÁS ESCOGER CARPETAS', 
                                         button_text='Guardar como',key='SaveAs',
                                         file_types=(("PDF", "*.pdf"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar')]]
                    window = sg.Window(name, layout, size=size2,alpha_channel=.95,
                                       disable_minimize=False,font=font_type,
                                       resizable=True,icon=icon)
                    
                    finished=False
                    while not finished:
                        event, values = window.read()
                        
                        if event=='Limpiar':
                            window.FindElement('InBrowse').Update('')
                            window.FindElement('InSave').Update('')
                        if event=='Atrás':
                            finished=True
                            window.close()
                        if event is None:
                            finished=True
                            done=True
                        if event=='Ejecutar':
                            finished=True
                            window.close()
                            try:
                                if values['SaveAs']!='':
                                    SaveAs=values['SaveAs']                                   
                                    if values2['choice']==choices[0] or values2['choice']==choices[2]: 
                                        SaveAs=os.path.abspath(SaveAs)
                                        if values['Browse']!='':
                                            Origin=values['Browse']
                                            Origin=os.path.abspath(Origin)    
                                            Origin=shutil.copy(Origin,Origin[:-4]+' sehr witzig.pdf')                                           
                                            if values2['choice']==choices[0]:    
                                                result=work_with_file(SaveAs,Origin)
                                                result.PDF_PDFA()
                                                notification(done_message).notif()      
                                            elif values2['choice']==choices[2]:
                                                result=work_with_file(SaveAs,Origin)
                                                result.PDF_same_size() 
                                                notification(done_message).notif()                 
                                        else:
                                            popup(error8).get_error_message()      
                                else:
                                    popup(error7).get_error_message()
                                             
                            except RuntimeError: 
                               popup(error2).get_error_message()
                #===================================================================================
                #OPTION NUMBER 2 WINDOW
                #===================================================================================                      
                elif values2['choice'] == choices[1]:  
                    
                    window_size=second_size
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[1][:400],size=window_size,font=font_type)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[1][400:],size=(750,15),font=font_type,text_color=bc)],
                              [sg.Button('Aquí tienes un TIP',key='Tip')],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Selecciona tus archivos: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FilesBrowse(tooltip='ESCOGE TUS ARCHIVOS', 
                                              button_text='Buscar',key='Browse',
                                              file_types=(("PDF", "*.pdf"),))],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(tooltip='SOLO PODRÁS ESCOGER CARPETAS', 
                                         button_text='Guardar como',key='SaveAs',
                                         file_types=(("PDF", "*.pdf"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar') ]]
                    
                    window = sg.Window(name, layout, size=(750,780),alpha_channel=.95,
                                       disable_minimize=False,font=font_type,
                                       resizable=True,icon=icon)
                    
                    finished=False
                    while not finished:
                        event, values = window.read()
                        
                        if event=='Limpiar':
                            window.FindElement('InBrowse').Update('')
                            window.FindElement('InSave').Update('')
                        if event=='Tip':
                            popup(Tip).get_message('TIP')
                        if event=='Atrás':
                            finished=True
                            window.close()
                        if event is None:
                            finished=True
                            done=True
                        if event=='Ejecutar':
                            finished=True
                            window.close()
                            try:
                                if values['SaveAs']!='':
                                    SaveAs=values['SaveAs'] 
                                    documents=values['Browse'].split(';')
                                    if len(documents)>1:
                                        result=work_with_documents(SaveAs,documents) 
                                        result.PDF_merger()
                                        notification(done_message).notif()                                          
                                    else:
                                        popup(error4).get_error_message()                                        
                                else:
                                    popup(error7).get_error_message()                                
                            except RuntimeError:
                                popup(error2).get_error_message()
                #===================================================================================
                #OPTION NUMBER 4 WINDOW
                #===================================================================================                                                                    
                elif values2['choice'] == choices[3]:
                    
                    window_size=fourth_size
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[3],size=window_size)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Escoge un archivo: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FileBrowse(tooltip='SELECCIONA UN ARCHIVO PDF',
                                             button_text='Buscar',
                                             key='Browse',file_types=(("PDF", "*.pdf"),))],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InFolder',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FolderBrowse(tooltip='SOLO PODRÁS VISUALIZAR CARPETAS',
                                               button_text='Buscar',key='Folder')],
                              [sg.Text("Ingresa un nombre: "),
                               sg.Input(key='SaveAs',default_text=default_name,
                                        tooltip='NOMBRE DEL ARCHIVO')],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar')]]
                    
                    window = sg.Window(name, layout, size=size2,alpha_channel=.95,
                                       disable_minimize=False,font=font_type,
                                       resizable=True,icon=icon)
                    
                    finished=False
                    while not finished:
                        event, values = window.read()
                        
                        if event=='Limpiar':
                            window.FindElement('InBrowse').Update('')
                            window.FindElement('InFolder').Update('')
                            window.FindElement('SaveAs').Update('')
                        if event=='Atrás':
                            finished=True
                            window.close()
                        if event is None:
                            finished=True
                            done=True
                        if event=='Ejecutar':
                            finished=True
                            window.close()
                            try:
                                if values['SaveAs']!='':
                                    SaveAs=values['SaveAs']                                    
                                    if values['Browse']!='':    
                                        Origin=values['Browse']
                                        Origin=os.path.abspath(Origin)                                        
                                        if values['Folder']=='':
                                            Folder=os.path.abspath(os.path.join(Origin, '..'))
                                        else:
                                            Folder=values['Folder']
                                            
                                        Folder=os.path.abspath(Folder)                           
                                        Folder=os.path.abspath(os.path.join(Origin, '..'))
                                        print(SaveAs)
                                        print(Folder)
                                        if values['SaveAs'] not in os.listdir(Folder):  
                                            SaveAsFilename=values['SaveAs']                                           
                                            try:                                               
                                                SaveAsDir=Folder+'\\'+SaveAsFilename   
                                                result=work_with_file(SaveAsDir,Origin)
                                                result.Req(SaveAsFilename) 
                                                notification(done_message).notif()
                                            except IndexError:                                               
                                                shutil.rmtree(Folder+'\\'+SaveAs) 
                                                #removes even non-empty folders
                                                popup(error1).get_error_message()        
                                        else:
                                            popup(error5).get_error_message()
                                    else:
                                        popup(error8).get_error_message()                                        
                                else:
                                    popup(error7).get_error_message()                              
                            except RuntimeError:
                                popup(error2).get_error_message()
                #===================================================================================
                #OPTION NUMBER 5 WINDOW
                #===================================================================================                                               
                elif values2['choice'] == choices[4]:     
  
                    get_text=sg.PopupGetText(title='¿Cuántos archivos PDF deseas unir?',
                                             message='Puedes ingresar hasta 15 archivos.',
                                             font=font_type,size=(55,200))
                    try:      
                        num_files = int(get_text)
                        if num_files>15:
                            num_files=15
                        elif num_files<2:
                            num_files=2
                    except TypeError:
                        break
                    except ValueError:
                        popup(error6).get_error_message()
                        break
                        
                    window_size=fifth_size   
        
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[4],size=window_size,font=font_type)],
                              [sg.Text('='*250,text_color=line_color)],
                             *[[sg.Text("Selecciona un archivo: "),
                               sg.Input(readonly=False,
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FilesBrowse(tooltip='ESCOGE VARIOS ARCHIVOS', 
                                              button_text='Buscar',
                                    file_types=(("PDF", "*.pdf"),))] for x in range(num_files)], 
                             [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(tooltip='GUARDAR COMO', button_text='Guardar como',
                                             key='SaveAs',file_types=(("PDF", "*.pdf"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar')]]
                    
                    window = sg.Window(name, layout, size=(750,int(300*(log(num_files+1)+0.5))),
                                       alpha_channel=.95,disable_minimize=False,
                                       font=font_type,resizable=True,icon=icon)
                    
                    finished=False
                    while not finished:
                        event, values = window.read()
                        if event=='Limpiar':
                            for i in range(num_files):
                                window.FindElement(i).Update('')
                            window.FindElement('InSave').Update('')
                        if event=='Atrás':
                            finished=True
                            window.close()
                        if event is None:
                            finished=True
                            done=True
                        if event=='Ejecutar':
                            finished=True
                            window.close()
                            try:
                                if values['SaveAs']!='':
                                    SaveAs=values['SaveAs']                                    
                                    documents=[]
                                    for i in range(num_files):
                                        documents.append(values[i])
                                    documents = list(filter(None, documents))
                                    if len(documents)>1:                                        
                                        result=work_with_documents(SaveAs,documents) 
                                        result.PDF_merger()                                         
                                        notification(done_message).notif()                                        
                                    else:
                                        popup(error4).get_error_message()                                                                                
                                else:
                                    popup(error7).get_error_message()                                
                            except RuntimeError:
                                popup(error2).get_error_message()
                                            
#===================================================================================================
#END
#===================================================================================================
visual() #Calls function and runs program