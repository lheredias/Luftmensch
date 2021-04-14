
#===================================================================================================
#IMPORTING GENERAL MODULES
#===================================================================================================

import os
import PySimpleGUI as sg
import sys #Used when creating .exe file
from shutil import move,copy,rmtree
from subprocess import Popen 
from math import log

#===================================================================================================
#SETTING UP APP ICON
#===================================================================================================

#We'll use Auto PY to EXE to convert the script into an .exe file so it can run over Windows.
#Now, things is, even thought we do want to include a taskbar icon, we do not 
#want to generate a directory with depencies inside. Instead, we fancy an standalone 
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
                              text_color=tc,icon=icon,modal=True)
        return None
    def get_error_message(self):
        '''Initializes a popup window when an error is raised'''
        sg.Popup(self.message, title='¡ALGO SALIÓ MAL!',
                             custom_text=('Regresar', None),icon=icon,
                             font=font_type,background_color=bc,text_color=tc)
        return None
    
#===================================================================================================
#SETTING UP MOST BASIC USER SETTINGS (to be used someday)
#===================================================================================================
    
# user_path='D:\\Luftmensch user settings.txt'
# def write_user_settings(value):
#     with open(user_path, "w") as text:
#         text.write(str(value))
#     return None  
# def read_user_settings():
#     with open(user_path, "r") as text:
#         return int(text.read())
# if os.path.basename(user_path) not in os.listdir(os.path.dirname(user_path)):
#     write_user_settings(1)

        #---------------------------(to be written in main window)---------------------------
            # check=read_user_settings()    
            # (...)       
            # if values2['CHECK']==True:
            #     write_user_settings(1)
            # elif values2['CHECK']==False:
            #     write_user_settings(0)       
        #---------------------------------------------------------------------------------------

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
        
    def make_txt(self,profile,RUC,doc):
        '''
        Creates a text file with some specifications required at work.
        '''
        
        text='6,'+RUC+','+profile+','+doc+',DESCARGA LE'
        
        with open(self.SaveAs, "w") as text_file:
            text_file.write('')
            text_file.write(text)
        
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
    
    def is_file_opened(self):
        '''
        Helper function to warn the user if output file (a PDF) exist and is currently being used
        by another process. 
        '''
        temp_filename=self.SaveAs[:-4]+' temp.pdf'
        if os.path.exists(self.SaveAs) == True:
            try:
                
                os.rename(self.SaveAs,temp_filename)
                os.rename(temp_filename,self.SaveAs)
               
                return False
            except PermissionError:
                return True
        else:
            return False
       
 
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
        import fitz
        
        
        self.word_pdf()

        pdf=fitz.open(self.SaveAs)
        opened_file=fitz.open(self.Origin)
        pdf.insertPDF(opened_file)
        # pdf.saveIncr()
        opened_file.close()
        pdf.deletePage(0)
        pdf.saveIncr()
        pdf.close()
        os.remove(self.Origin)
        
        return None
    
    def get_text(self):
        '''
        Extracts text from PDF an stores it as .txt file
        '''
        import fitz
        
        #If user decides to replace an existing .txt document, we make sure to leave it blank 
        #before extracting the plain text from the PDF and writing it into the .txt file
        with open(self.SaveAs, "w") as text_file: 
                text_file.write('')
        
        #Extract text in (kind of) natural reading order
        doc = fitz.open(self.Origin)
        for num_page in range(doc.pageCount):
            page = doc.loadPage(num_page)
            blocks = page.getText("blocks")
            blocks.sort(key=lambda block: block[1])  # sort vertically ascending  
            for b in blocks:
                if b[4][0] != '<' and b[4][len(b[4])-1] != '>': 
                    # b[4] is the text part of each block
                    with open(self.SaveAs, "a") as text_file:  
                        text_file.write(b[4])
        doc.close()
        
        return None   
            
    def Req(self,name):
        '''
        Very useful at work.
        Creates a standard zip file required at work alongside a folder with its content. 
        A nice example of using the 're' module.
        '''   
        from re import findall
        import fitz
        from shutil import make_archive
        
        os.mkdir(self.SaveAs)
        os.mkdir(os.path.join(self.SaveAs,name))
        
        doc = fitz.open(self.Origin)
        page = doc.loadPage(0)              
        text = page.getText()  
        nro_req=findall('\s+\d{13}|°+\d{13}|º+\d{13}', text)
        nro_ruc=findall('\s+\\b1\d{10}\\b|°+\\b1\d{10}\\b|º+\\b1\d{10}\\b|\s+\\b2\d{10}\\b|°+\\b2\d{10}\\b|º+\\b2\d{10}\\b', text)  
        doc.close()
        nro_ruc = nro_ruc[0].replace("°","").replace("º","").replace(" ","").strip()
        nro_req = nro_req[0].replace("°","").replace("º","").replace(" ","").strip()
        
        text_file_path=self.SaveAs+'\\files.txt'
        text_file = open(text_file_path, 'w')
        text_file.write(nro_ruc+'|'+nro_req)
        text_file.close()
        
        pdf_element=self.SaveAs+'\\'+name+'\\'+nro_ruc+'_'+nro_req+'.pdf'
        text_element=self.SaveAs+'\\'+name+'\\files.txt'
        
        copy(self.Origin,pdf_element)
        move(text_file_path,text_element)
   
        make_archive(self.SaveAs+'\\'+name, 'zip',root_dir=self.SaveAs , base_dir=name)
  
        return None
    
    def PDF_same_size(self):
        '''
        Have you all of your PDF's pages adopt vertical A4 dimensions.
        Best part: gets the job done without messing up its contents.
        Credits go to user "JorjMcKie" from the PyMuPDF team on GitHub.
        '''
        import fitz
        from time import sleep
        
        src= fitz.open(self.Origin)
        doc = fitz.open()
        for ipage in src:
            #Some pages, even though having a landscape aspect, are not actually 'landscape', 
            #just rotated.
            #These lines of code take care of the many problems that arise when dealing with 
            #different page sizes. 
            if ipage.get_contents() != []:
                if ipage.rotation==90: 
                    ipage.setRotation(0)
                fmt = fitz.PaperRect("a4")
                page = doc.newPage(width = fmt.width, height = fmt.height)
                page.showPDFpage(page.rect, src, ipage.number)
        src.close() 
        
       
        #-----------------------------------------MINI-LOOP----------------------------------------# 
        #If there's an issue with given paths, doc.save() will raise a RuntimeError if run 
        #inmeadiatelly after closing source file. Therefore, a mini-loop is implemented so it can 
        #give the module (fitz) enough time to properly perform the process. Time limit is set 
        #at 5 seconds.
        count=0
        while count<5:
            try:
                print(count)
                doc.save(self.Origin)
                count=5
                print('sucess!')
            except RuntimeError:
                print('Pymupdf permission denied')
                sleep(count)
                count+=0.1
        #-----------------------------------------MINI-LOOP----------------------------------------#        
        doc.close()
        move(self.Origin,self.SaveAs)
        
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
    
    def IMG_to_PDF(self):
        '''
        Image-to-PDF converter. If there is more than one image file, it will create a single PDF 
        with one image per page.
        '''
        import fitz
        from time import sleep
        doc = fitz.open()
        for file in self.documents:
            img=fitz.open(file)
            fmt = fitz.PaperRect("a4")
            pdfbytes = img.convertToPDF()
            img.close()
            imgPDF = fitz.open("pdf", pdfbytes)
            page = doc.newPage(width = fmt.width, height = fmt.height)
            page.showPDFpage(page.rect, imgPDF, 0)   
           
        #-----------------------------------------MINI-LOOP----------------------------------------# 
        #If there's an issue with given paths, doc.save() will raise a RuntimeError if run 
        #inmeadiatelly after closing source file. Therefore, a mini-loop is implemented so it can 
        #give the module (fitz) enough time to properly perform the process. Time limit is set 
        #at 5 seconds.
        count=0
        while count<5:
            try:
                print(count)
                doc.save(self.SaveAs)
                count=5
                print('sucess!')
            except RuntimeError:
                print('Pymupdf permission denied')
                sleep(count)
                count+=0.1
        #-----------------------------------------MINI-LOOP----------------------------------------#        
        doc.close()
        
        return None

#===================================================================================================
#USEFUL DEFINITIONS
#===================================================================================================

default_name='sehr witzig' #default output/folder name

description='Versión 1.2.1' #Version number
name='LuftMensch' #App title
bottom_description='Una aplicación de productividad' #Short description
note=''  #Warning-style caption
note2='Nota: en Microsoft Word, la opción "compatible con PDF/A" debe encontrarse activada.'

#Error messages
error1='No se encontró el número de Requerimiento y/o RUC.'
error2='Algo salió mal; asegúrate de que los directorios de origen y destino sean válidos.'
error3='Ingresa el número de RUC y OF.'
error4='Si lo que deseas es unir PDFs, debes seleccionar más de un archivo.'
error5='Ya existe una carpeta con ese nombre; por favor Ingresa otro.'
error6='Debes ingresar un número.'
error7='Debes ingresar un nombre para tu archivo.'
error8='Debes seleccionar un archivo para trabajar.'
error9='El archivo sobre el cual deseas guardar tu documento se encuentra abierto; ciérralo y vuelve a intentar.'
error10='No se puede guardar el documento encima de uno de los PDFs a unir.'
error11='Debes ingresar un número de RUC y/o documento de sustento válidos.'
done_message='Ya puedes visualizar tu documento.'

Tip1='''
Si deseas unir solo unos cuantos de los PDFs que se encuentran en la carpeta seleccionada, haz lo siguiente:
 
Durante la selección, crea una nueva carpeta y copia o arrastra hacia allí los PDFs a unir. 

Adicionalmente, si deseas unirlos en un orden específico, ordénalos dentro de la carpeta creada, selecciona el primer archivo y dale CTRL+E.
'''
Tip2='''
Si deseas unir solo unas cuantas de las imágenes que se encuentran en la carpeta seleccionada:
 
Durante la selección, crea una nueva carpeta y copia o arrastra hacia allí las imágenes a unir. 

Adicionalmente, si deseas unirlas en un orden específico, ordénalas dentro de la carpeta creada, selecciona el primer archivo y dale CTRL+E.
'''
#===================================================================================================
#OPTIONS
#===================================================================================================

abouts=['','','','','','','',''] #Options
abouts[0]='''Convierte un archivo PDF en formato PDF/A.

Si dejas en blanco "Guardar como", el  resultado se guardará encima del documento seleccionado.
'''
abouts[1]='''Une varios archivos PDF en uno solo.

El resultado es un PDF ideal para lectura e impresión.

Usa esta opción si tienes muchos archivos para unir.

Si quieres unir tus archivos en ORDEN, sigue estos pasos:
    
  - Asegúrate de que tus PDFs se encuentren en el orden deseado en la pantalla.
    
  - Si no lo están, puedes ordenarlos allí mismo.
    
  - Selecciona el primer archivo y dale "CTRL + E".
    
  - No importa si en la selección se incluyeron carpetas; no serán tomadas en cuenta.
    
  - Si deseas retirar archivos de la lista seleccionada, mantén presionado CTRL y dales click.'''
  
abouts[2]='''Logra que todas páginas del PDF seleccionado posean las mismas dimensiones (A4 Vertical), 
sin alterar sus proporciones, ni comprometer su correcta visualización.

Si dejas en blanco "Guardar como", el  resultado se guardará encima del documento seleccionado.'''

abouts[3]='''Genera el archivo modelo (.zip) de un Requerimiento o Resultado de 
Requerimiento con el contenido apropiado, listo para notificar. 

Adicionalmente, se generará una carpeta descomprimida de dicho archivo 
para que puedas visualizar su contenido sin necesidad de descomprimir el fichero.

Solo escoge tu documento PDF, selecciona la carpeta donde quieres que se almacene
el archivo .zip e ingresa un nombre.

Si no ingresas una carpeta de destino, el archivo resultante se irá por defecto
a la carpeta que contenga el documento seleccionado.'''

abouts[4]='''Une varios archivos PDF en uno solo.

El archivo PDF resultante es ideal para lectura e impresión.

Usa esta opción si tienes pocos PDFs para unir, se encuentran en distintas 
carpetas o simplemente prefieres cargarlos uno por uno.

Si seleccionas un número menor de PDFs al número ingresado (seleccionaste dos, pero en la pantalla 
previa ingresaste seis, por ejemplo), no te preocupes; puedes dejar los demás espacios en blanco.'''

abouts[5]='''Extrae el texto plano de un archivo PDF.'''

abouts[6]='''Convierte una o varias imágenes en un solo archivo PDF de dimensiones A4 vertical.

Si quieres unir tus imágenes en orden: ordénalas, selecciona el primer archivo y dale CTRL + E.

Si no deseas unir todas las imágenes dentro de la carpeta: luego de seleccionar el primer archivo 
y darle CTRL + E, presiona CTRL una vez más y, sin soltar, da click en aquellas imágenes que 
deseas quitar de la lista.'''

abouts[7]='''Genera el archivo de texto para solicitar la descarga de libros electrónicos.

Ingresa el número de RUC y de documento de sustento según tu perfil.'''
   
choices=['1. Convertir PDF en PDF/A',
                 '2. Unir varios archivos PDF',
                 '3. Lograr que todas las páginas de un PDF posean el mismo tamaño',
                 '4. Crear archivo .zip de Requerimiento',
                 '5. Unir varios archivos PDF, cargándolos uno por uno',
                 '6. Extraer el texto de un archivo PDF',
                 '7. Convierte una o varias imágenes en un solo archivo PDF',
                 '8. Generar archivo de texto para solicitar descarga de LE'] 

profiles=['F01 - ORDEN DE FISCALIZACIÓN',
       'F02 - ACCIÓN INDUCTIVA - ESQUELA',
       'F03 - PROGRAMA DE FISCALIZACIÓN - ADUANAS',
       'F04 - ACCIÓN INDUCTIVA - CARTA INDUCTIVA'] 

#===================================================================================================
#MENU
#===================================================================================================

menu=['Acerca de esta aplicación','Ir al repositorio'] 

first='''
Luftmensch es una aplicación de productividad y de código abierto pensada en automatizar ciertas tareas administrativas.

Si deseas descargar la última versión, realizar consultas, hacer algún comentario o dejar una sugerencia; visita el repositorio. 
'''
#Menu disctionary with names and descriptions
menu_dict={menu[0]:first}

#Menu layout for GUI
menu_def = [['Menú',[menu[0],menu[1]]]]   


#===================================================================================================
#GRAPHICAL USER INTERFACE: SETTING UP THE BASIC STUFF
#===================================================================================================

font_type=('Helvetica', 12) #Layout font type and size

size1=(72,10) #Option description layout size
size2=(750,550) #GUI main window layout size
size3=(750,500)
size4=(750,420)
size5=(750,600)
first_size=(750,3)
second_size=(750,7)
third_size=(750,4)
fourth_size=(750,11)
fifth_size=(750,9)
sixth_size=(750,1)

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
        layout = [[sg.Menu(menu_def,key='MENU', tearoff=False)],
                  [sg.T("")],[sg.Text(name[:4],font=('Gotham Bold', 45),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 45),text_color=bc),
                              sg.Text(description,font=('Gotham', 12),text_color=line_color)],   
                  [sg.Text(bottom_description,font=('Gotham Bold', 11),text_color=line_color)],
                  [sg.T("")],
                  [sg.Text('='*250,text_color=line_color)],   
                  [sg.Text("Escoge una opción: "),
                   sg.OptionMenu(values=choices,key='choice',size=(90,1))],
                  [sg.Text('='*250,text_color=line_color)],
                  [sg.Submit('CONTINUAR'), sg.Cancel('SALIR')],
                  [sg.Text('='*250,text_color=line_color)],
                  [sg.Text(note,size=first_size,text_color=bc)]]
        
        window = sg.Window(name, layout, size=size4,alpha_channel=.95,disable_minimize=False,
                           font=font_type,icon=icon)
    
        finished=False
        while not finished:
            event, values2 = window.read()
            if event==menu[0]:
                popup(menu_dict[menu[0]]).get_message(menu[0])    
            if event==menu[1]:
                from webbrowser import open as op
                op('https://github.com/lheredias/Luftmensch')
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
                        choice=choices[0]
                        description_size=first_size   
                    else:
                        about=abouts[2]
                        choice=choices[2]
                        description_size=third_size
                         
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text(choice,font=('Gotham Bold', 12),text_color=line_color)],
                              [sg.Text('='*250,text_color=line_color)],
                        [sg.Text(about,size=description_size)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Escoge un archivo: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FileBrowse(button_text='Buscar', key='Browse',
                                             file_types=(("PDF", "*.pdf"),))],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(button_text='Guardar como',key='SaveAs',
                                         file_types=(("PDF", "*.pdf"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar')],
                          [sg.Text('='*250,text_color=line_color)],
                          [sg.Checkbox('ABRIR DE INMEDIATO EL DOCUMENTO GENERADO', 
                               default=False,text_color=bc, key='CHECK')],
                          *[[sg.Checkbox('CONVERTIR DE INMEDIATO A PDF/A', 
                               default=False,text_color=bc, key='PDFA')] 
                            if values2['choice'] == choices[2] else [sg.T(note2)]],
                          [sg.Text('='*250,text_color=line_color)]]
                    window = sg.Window(name, layout, size=size3,alpha_channel=.95,
                                       disable_minimize=False,font=font_type,icon=icon)
                    
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
                                    SaveAs=os.path.abspath(SaveAs)
                                elif values['SaveAs']=='':
                                    SaveAs=values['Browse']
                                    SaveAs=os.path.abspath(SaveAs)
                                if values2['choice']==choices[0] or values2['choice']==choices[2]: 
                                    
                                    if values['Browse']!='':
                                        Origin=values['Browse']
                                        Origin=os.path.abspath(Origin)    
                                        Origin=copy(Origin,Origin[:-4]+' sehr witzig.pdf')
                                        result=work_with_file(SaveAs,Origin)
                                        if values2['choice']==choices[0]:
                                            
                                            if result.is_file_opened()==False:
                                                result.PDF_PDFA()
                                                if values['CHECK']==True:
                                                    Popen([SaveAs],shell=True)
                                                else:
                                                    notification(done_message).notif()
                                            else:
                                                os.remove(Origin)
                                                popup(error9).get_error_message()   
                                        elif values2['choice']==choices[2]:
                                            
                                            if result.is_file_opened()==False:
                                                result.PDF_same_size()
                                                if values['PDFA']==True:
                                                    Origin=SaveAs  
                                                    Origin=copy(Origin,Origin[:-4]+' sehr witzig.pdf')
                                                    result=work_with_file(SaveAs,Origin)
                                                    result.PDF_PDFA()
                                                if values['CHECK']==True:
                                                    Popen([SaveAs],shell=True)
                                                else:
                                                    notification(done_message).notif()
                                            else:
                                                os.remove(Origin)
                                                popup(error9).get_error_message()                                 
                                    else:
                                        popup(error8).get_error_message()      
                             
                                             
                            except RuntimeError: 
                               popup(error2).get_error_message()
                #===================================================================================
                #OPTION NUMBER 2 WINDOW
                #===================================================================================                      
                elif values2['choice'] == choices[1]:  
                    choice=choices[1]
                    description_size=second_size
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text(choice,font=('Gotham Bold', 12),text_color=line_color)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[1][:210],size=second_size,font=font_type)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[1][212:],size=(750,10),font=font_type,text_color=bc)],
                              [sg.Button('Aquí tienes un TIP',key='Tip')],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Selecciona tus archivos: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FilesBrowse(button_text='Buscar',key='Browse',
                                              file_types=(("PDF", "*.pdf"),))],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(button_text='Guardar como',key='SaveAs',
                                         file_types=(("PDF", "*.pdf"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar')],
                          [sg.Text('='*250,text_color=line_color)],
                          [sg.Checkbox('ABRIR DE INMEDIATO EL DOCUMENTO GENERADO', 
                               default=False,text_color=bc, key='CHECK')],
                          [sg.Text('='*250,text_color=line_color)]]
                    
                    window = sg.Window(name, layout, size=(750,820),alpha_channel=.95,
                                       disable_minimize=False,font=font_type,icon=icon)
                    
                    finished=False
                    while not finished:
                        event, values = window.read()
                        
                        if event=='Limpiar':
                            window.FindElement('InBrowse').Update('')
                            window.FindElement('InSave').Update('')
                        if event=='Tip':
                            popup(Tip1).get_message('TIP1')
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
                                    SaveAs=os.path.abspath(SaveAs)
                                    documents=values['Browse'].split(';')
                                    if len(documents)>1:
                                        if values['SaveAs'] in documents:
                                            popup(error10).get_error_message() 
                                        else: 
                                            result=work_with_documents(SaveAs,documents)
                                            if result.is_file_opened()==False:
                                                result.PDF_merger()
                                                if values['CHECK']==True:
                                                    Popen([SaveAs],shell=True)
                                                else:
                                                    notification(done_message).notif() 
                                            else:
                                                popup(error9).get_error_message()  
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
                    choice=choices[3]
                    description_size=fourth_size
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text(choice,font=('Gotham Bold', 12),text_color=line_color)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[3],size=description_size)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Escoge un archivo: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FileBrowse(button_text='Buscar',
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
                                       disable_minimize=False,font=font_type,icon=icon)
                    
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
                                        if values['SaveAs'] not in os.listdir(Folder):  
                                            SaveAsFilename=values['SaveAs']                                           
                                            try:                                               
                                                SaveAsDir=Folder+'\\'+SaveAsFilename   
                                                result=work_with_file(SaveAsDir,Origin)
                                                result.Req(SaveAsFilename) 
                                                notification(done_message).notif()
                                            except IndexError:                                               
                                                rmtree(Folder+'\\'+SaveAs) 
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
                                             icon=icon,
                                             message='Ingresa un número entre 2 y 15.',
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
                        
                    description_size=fifth_size   
                    
                    choice=choices[4]
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text(choice,font=('Gotham Bold', 12),text_color=line_color)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[4],size=description_size,font=font_type)],
                              [sg.Text('='*250,text_color=line_color)],
                             *[[sg.Text("Selecciona un archivo: "),
                               sg.Input(readonly=False,
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FilesBrowse(button_text='Buscar',
                                    file_types=(("PDF", "*.pdf"),))] for x in range(num_files)], 
                             [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(button_text='Guardar como',
                                             key='SaveAs',file_types=(("PDF", "*.pdf"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar'),
                          sg.Checkbox('ABRIR DE INMEDIATO EL DOCUMENTO GENERADO', 
                               default=True,text_color=bc, key='CHECK')],
                          [sg.Text('='*250,text_color=line_color)]]
                    
                    window = sg.Window(name, layout, size=(750,int(300*(log(num_files+1)+0.6))),
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
                                    SaveAs=os.path.abspath(SaveAs) 
                                    documents=[]
                                    for i in range(num_files):
                                        documents.append(values[i])
                                    documents = list(filter(None, documents))
                                    if len(documents)>1:  
                                        if values['SaveAs'] in documents:
                                            popup(error10).get_error_message() 
                                        else: 
                                            result=work_with_documents(SaveAs,documents)
                                            if result.is_file_opened()==False:
                                                result.PDF_merger()
                                                if values['CHECK']==True:
                                                    Popen([SaveAs],shell=True)
                                                else:
                                                    notification(done_message).notif() 
                                            else:
                                                popup(error9).get_error_message()                                       
                                    else:
                                        popup(error4).get_error_message()                                                                                
                                else:
                                    popup(error7).get_error_message()                                
                            except RuntimeError:
                                popup(error2).get_error_message()
                #===================================================================================
                #OPTION NUMBER 6 WINDOW
                #===================================================================================                      
                elif values2['choice'] == choices[5]:     
                    
                    about=abouts[5]
                    choice=choices[5]
                    description_size=sixth_size
                        
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text(choice,font=('Gotham Bold', 12),text_color=line_color)],
                              [sg.Text('='*250,text_color=line_color)],
                        [sg.Text(about,size=description_size)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Escoge un archivo: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FileBrowse(button_text='Buscar', key='Browse',
                                             file_types=(("PDF", "*.pdf"),))],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(button_text='Guardar como',key='SaveAs',
                                         file_types=(("Archivo de texto", "*.txt"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar')],
                          [sg.Text('='*250,text_color=line_color)],
                          [sg.Checkbox('ABRIR DE INMEDIATO EL DOCUMENTO GENERADO', 
                               default=True,text_color=bc, key='CHECK')],
                          [sg.Text('='*250,text_color=line_color)]]
                    
                    window = sg.Window(name, layout, size=size4,alpha_channel=.95,
                                       disable_minimize=False,font=font_type,icon=icon)
                    
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
                                    SaveAs=os.path.abspath(SaveAs)
                                    if values['Browse']!='':
                                        Origin=values['Browse']
                                        Origin=os.path.abspath(Origin)
                                        result=work_with_file(SaveAs,Origin)
                                        result.get_text()
                                        if values['CHECK']==True:
                                            Popen([SaveAs],shell=True)
                                        else:
                                            notification(done_message).notif() 
                                    else:
                                        popup(error8).get_error_message()      
                                else:
                                    popup(error7).get_error_message()
                                             
                            except RuntimeError: 
                               popup(error2).get_error_message()  
                #===================================================================================
                #OPTION NUMBER 7 WINDOW
                #===================================================================================                      
                elif values2['choice'] == choices[6]:  
                    choice=choices[6]
                    description_size=second_size
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text(choice,font=('Gotham Bold', 12),text_color=line_color)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[6][:400],size=description_size,font=font_type)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Button('Aquí tienes un TIP',key='Tip')],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Selecciona tus archivos: "),
                               sg.Input(readonly=True,key='InBrowse',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.FilesBrowse(button_text='Buscar',key='Browse',
                                file_types=(("IMÁGENES", "*.png *.jpg *.jpeg *.jfif *.tiff"),))],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(button_text='Guardar como',key='SaveAs',
                                         file_types=(("PDF", "*.pdf"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar') ],
                          [sg.Text('='*250,text_color=line_color)],
                          [sg.Checkbox('ABRIR DE INMEDIATO EL DOCUMENTO GENERADO', 
                               default=True,text_color=bc, key='CHECK')],
                          [sg.Text('='*250,text_color=line_color)]]
                    
                    window = sg.Window(name, layout, size=size5,alpha_channel=.95,
                                       disable_minimize=False,font=font_type,icon=icon)
                    
                    finished=False
                    while not finished:
                        event, values = window.read()
                        
                        if event=='Limpiar':
                            window.FindElement('InBrowse').Update('')
                            window.FindElement('InSave').Update('')
                        if event=='Tip':
                            popup(Tip2).get_message('TIP2')
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
                                    SaveAs=os.path.abspath(SaveAs)
                                    documents=values['Browse'].split(';')
                                    result=work_with_documents(SaveAs,documents)
                                    if result.is_file_opened()==False:
                                        result.IMG_to_PDF()
                                        if values['CHECK']==True:
                                            Popen([SaveAs],shell=True)
                                        else:
                                            notification(done_message).notif() 
                                    else:
                                        popup(error9).get_error_message()                                           
                                else:
                                    popup(error7).get_error_message()                                
                            except RuntimeError:
                                popup(error2).get_error_message() 
                #===================================================================================
                #OPTION NUMBER 8 WINDOW
                #===================================================================================                      
                elif values2['choice'] == choices[7]:  
                    choice=choices[7]
                    description_size=first_size
                    layout = [[sg.Text(name[:4],font=('Gotham Bold', 25),text_color=tc2),
                              sg.Text(name[4:],font=('Gotham Bold', 25),text_color=bc)],
                              [sg.Text(choice,font=('Gotham Bold', 12),text_color=line_color)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text(abouts[7],size=description_size,font=font_type)],
                              [sg.Text('='*250,text_color=line_color)],
                              [sg.Text("Selecciona tu perfil: "),
                   sg.OptionMenu(values=profiles,key='profile',size=(90,1))],
                              [sg.Text('='*250,text_color=line_color)],
                                [sg.Text('RUC:', size =(18, 1)), sg.InputText(key='RUC')], 
                                [sg.Text('Documento de sustento:', size =(18, 1)), 
                                 sg.InputText(key='doc')],
                              [sg.Text("Escoge el destino: "),
                               sg.Input(readonly=True,key='InSave',
                                        disabled_readonly_background_color=bc,
                                        disabled_readonly_text_color=tc),
                               sg.SaveAs(button_text='Guardar como',key='SaveAs',
                                         file_types=(("Texto", "*.txt"),))],
                              [sg.Text('='*250,text_color=line_color)],
                          [sg.Submit('Ejecutar'), sg.Cancel('Atrás'),sg.Button('Limpiar') ],
                          [sg.Text('='*250,text_color=line_color)],
                          [sg.Checkbox('ABRIR DE INMEDIATO EL DOCUMENTO GENERADO', 
                               default=False,text_color=bc, key='CHECK')],
                          [sg.Text('='*250,text_color=line_color)]]
                    
                    window = sg.Window(name, layout, size=size2,alpha_channel=.95,
                                       disable_minimize=False,font=font_type,icon=icon)
                    
                    finished=False
                    while not finished:
                        event, values = window.read()
                        if event=='Limpiar':
                            window.FindElement('RUC').Update('')
                            window.FindElement('doc').Update('')
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
                            if values['profile']==profiles[0]:
                                profile='F01'
                            elif values['profile']==profiles[1]:
                                profile='F02'
                            elif values['profile']==profiles[2]:
                                profile='F03'
                            elif values['profile']==profiles[3]:
                                profile='F04'
                            try:
                                if values['SaveAs']!='':
                                    SaveAs=values['SaveAs']
                                    SaveAs=os.path.abspath(SaveAs)
                                    if values['doc'] !='' and values['doc']!='':
                                        RUC=values['RUC'].strip()
                                        doc=values['doc'].strip()
                                        if len(RUC)==11 and len(doc)==12 or len(doc)==18:
                                            result=work_with_inputs(SaveAs) 
                                            result.make_txt(profile,RUC,doc)
                                            if values['CHECK']==True:
                                                Popen([SaveAs],shell=True)
                                            else:
                                                notification(done_message).notif() 
                                        else:
                                            popup(error11).get_error_message() 
                                    else:
                                        popup(error3).get_error_message() 
                                else:
                                    popup(error7).get_error_message()                                
                            except RuntimeError:
                                popup(error2).get_error_message()                                                                             
#===================================================================================================
#END
#===================================================================================================
visual() #Calls function and runs program
