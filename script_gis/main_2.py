from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.service import Service as FirefoxService

from selenium.webdriver.common.by import By
import credentials
import time
import os
from datetime import date
import pyautogui
import pandas as pd
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import tkinter as tk



def getTitulo(tit):
    return tit.get()



def save(top,success,errors,nome):
    endereco_final = filedialog.askdirectory(initialdir="C:/")
    data = date.today()
    os.mkdir("%s/%s_%s"%(endereco_final,nome,data))
    success.to_csv("%s/%s_%s/%s_report.csv"%(endereco_final,nome,data,nome), index = False,sep=',')
    lista_status = success['Status'].to_list()
    print(lista_status)
    messagebox.showinfo("SACT","Relatório salvo na pasta: %s\nShipments With Invoice: %s\nShipments Without Invoice: %s"%(endereco_final,str(lista_status.count("SUCCESSFUL")),str(len(lista_status)-lista_status.count("SUCCESSFUL"))))
    top.destroy()

def sobrescrever(top,success,errors,nome):
    endereco_final = filedialog.askdirectory(initialdir="C:/")
    data = date.today()
    success.to_csv("%s/%s_report.csv"%(endereco_final,nome), index = False,sep=',')
    lista_status = success['Status'].to_list()
    print(lista_status)
    messagebox.showinfo("SACT","Relatório salvo na pasta: %s\nShipments With Invoice: %s\nShipments Without Invoice: %s"%(endereco_final,str(lista_status.count("SUCCESSFUL")),str(len(lista_status)-lista_status.count("SUCCESSFUL"))))
    top.destroy()

def arsenal(driver,shipments):
    ships = ""
    for ship in shipments:
        ships = ships + ship + "\n"
    driver.get('https://arsenal-na.amazon.com/shipment/encryptDecrypt')
    WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.ID,"shipmentIds")))
    
    fill_shipments = driver.find_element(By.ID,"shipmentIds")
    fill_shipments.send_keys(ships)

    submit_button = driver.find_element(By.ID,'submitBtn')
    submit_button.click()

    time.sleep(2)

    tables = driver.find_elements(By.CLASS_NAME,"table-result")
    decrypted_table = tables[1]
    elements = decrypted_table.find_elements(By.TAG_NAME,"tr")
    dec_shipments = []
    for element in elements:
        dec_shipments.append(element.text)
    dec_shipments = dec_shipments[1::]
    retorno_shipments = []
    for i in range(len(shipments)):
        if dec_shipments[i] == '0':
            retorno_shipments.append((shipments[i],'Invalid Shipment'))
        else:
            retorno_shipments.append((shipments[i],dec_shipments[i]))
    print(retorno_shipments)
    return(retorno_shipments)    

def rodar(root,top,textos):
    #Save XML folder path
    cwd = str(os.getcwd())
    xml_folder = cwd + "\\xml_folder"

    texto = textos
    shipments = texto.splitlines()
    my_progress2 = ttk.Progressbar(top,orient=HORIZONTAL,
    length=300,mode='determinate')
    my_progress2.pack(pady=10)
    step = 100/len(shipments)
    text = StringVar()
    label1  = Label(top,textvariable=text).pack()
    text.set("Quantidade de Shipments Encontrados: 0/%i"%(len(shipments)))
    top.update_idletasks()
    root.update()
    print(len(shipments))
    with open("retorno.csv","w") as file:
        file.writelines("ClientId Domain,Client Id,Status,GiasId,GiasId Type,Client Id Encrypted\n")



    #Start Web Driver
    options = webdriver.FirefoxOptions()

    prefs = {
        'safebrowsing.enabled' : 'false', 
        'download.default_directory' : xml_folder,
        }
    driver = webdriver.Firefox(options=options,service_log_path="null")

    driver.get('https://gisportal.aka.amazon.com/#/gisv2')
    print("Please Login")

    username = driver.find_element(By.ID,'user_name_field')
    username.send_keys(credentials.gis['username'])

    time.sleep(2)
    submit_button = driver.find_element(By.ID,'user_name_btn')
    submit_button.click()

    time.sleep(2)
    password = driver.find_element(By.ID,'password_field')
    password.send_keys(credentials.gis['password'])

    login_button = driver.find_element(By.ID,'password_btn')
    login_button.click()

    print('Waiting for OTP')
    root.update()

    time.sleep(10)
    print('Login done')

    shipments = arsenal(driver,shipments)

    iteracoes = 1 + len(shipments)//100
    initial = 0
    ships = ""
    if iteracoes < 2:
        end = len(shipments)
    else:
        end = 99
    top.update_idletasks()
    root.update()
    encontrados = 0
    time.sleep(2)
    driver.get('https://gisportal.aka.amazon.com/#/gisv2-batch')
    WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"form-control")))
    time.sleep(10)
    for i in range(iteracoes):
        print("Rodando Batch %i/%i"%(i+1,1+len(shipments)//100))
        forms = driver.find_elements(By.CLASS_NAME,"form-control")
        forms[0].click()
        select = Select(forms[0])
        select.select_by_value("0")
        for ini in range(initial,end):
            if shipments[ini][1] != 'Invalid Shipment':
                ships = ships + shipments[ini][1] + "," 
        forms[1].clear()
        forms[1].send_keys(ships[:len(ships)-1:])
        getStatus = driver.find_element(By.ID,'btn-getStatus')
        getStatus.click()
        time.sleep(3)
        respostas = driver.find_element(By.CLASS_NAME,'table-striped')
        resp = respostas.find_elements(By.CLASS_NAME,"ng-scope")
        encontrados = encontrados + len(resp)
        text.set("Quantidade de Shipments Encontrados: %i/%i"%(encontrados,len(shipments)))
        for r in resp:
            my_progress2['value'] += step
            top.update_idletasks()
            root.update()
            fields = r.find_elements(By.CLASS_NAME,"ng-binding")
            ret = ""
            for field in fields:
                ret = ret +'"'+field.text +'"'+ ","
            ret = ret[:len(ret):]
            with open("retorno.csv","a") as file:
                file.writelines(ret + "\n")
        initial = end
        if(end+99>len(shipments)):
            end = len(shipments)
        else:
            end+=99

        
        ships = ""

    csv = pd.read_csv("retorno.csv",header=0)
    lista_success = csv["Client Id"].values.tolist()
    for item in shipments:
        csv.loc[csv["Client Id"].astype('str').str.contains(item[1]),"Client Id Encrypted"] = item[0] 
    with open("shipments_without_invoice.csv","w") as file:
        for s in shipments:
            print(s)
            if(s[1]) != 'Invalid Shipment':
                if lista_success.count(int(s[1])) < 1:
                    with open("shipments_without_invoice.csv","w") as file:
                            if lista_success.count(int(s[1])) < 1:
                                file.write("%s,%s\n"%(s[0],s[1]))
                                new_row = {
                                    'Client Id Domain': "",
                                    'Client Id': s[1],
                                    'Status': "Not Found",
                                    "GiasId": "",
                                    "GiasId Type": "",
                                    "Client Id Encrypted": s[0]
                                }
                                csv = csv.append(new_row,ignore_index=True)
            else:
                with open("shipments_without_invoice.csv","w") as file:
                   file.write("%s,%s\n"%(s[0],s[1]))
                new_row = {
                    'Client Id Domain': "",
                    'Client Id': s[1],
                    'Status': "Invalid Shipment",
                    "GiasId": "",
                    "GiasId Type": "",
                    "Client Id Encrypted": s[0]
                }
                csv = csv.append(new_row,ignore_index=True)

    erros = pd.read_csv("shipments_without_invoice.csv",names=("Shipments Not Found","Decrypted Shipments Not Found"))
    print(erros)
    driver.close()
    instlot = Label(top,text='Escreva o título do Lote:').pack()
    titlot = Entry(top,width=50)
    titlot.pack()
    butlot = Button(top,text='Nova Pasta',command=lambda:save(top,csv,erros,getTitulo(titlot)))
    butlot.pack()
    butlot2 = Button(top,text='Salvar',command=lambda:sobrescrever(top,csv,erros,getTitulo(titlot)))
    butlot2.pack()
    tup = pyautogui.size()
    wi = (tup[0]/2)-200
    he = (tup[1]/2)-200
    top.geometry('400x550+%i+%i'%(wi,he))

def parse_xml(xml_folder):
    erp_list = []
    for filename in os.listdir(xml_folder):
        xml = os.path.join(xml_folder, filename)
        print(xml)
        with open(xml) as f:
            try:
                verProc = str(f.read()).split("<verProc>")[1].split("</verProc>")[0]
                erp_list.append([filename[:15],verProc])
            except:
                erp_list.append([filename[:15],'parsing_error'])
    return erp_list

def open_interface(root):
    tup = pyautogui.size()
    wi = (tup[0]/2)-200
    he = (tup[1]/2)-200
    top66 = Toplevel(root)
    top66.geometry('400x400+%i+%i'%(wi,he))
    top66.title('SACT')
    top66.iconbitmap("C:/Users/brunnomi/Downloads/script_gis/script_gis/amazon-logo.ico")
    label1 = Label(top66,text='Insira os shipments:').pack()
    textos = Text(top66,width=20,height=10)
    textos.pack() 
    botao = Button(top66,text='Começar!',command=lambda: rodar(root,top66,textos.get("1.0","end-1c"))).pack()

def update():    
    root.after(1000, update)   

def arsenal_tcorp(driver,shipments):
    ships = ""
    for ship in shipments:
        ships = ships + ship[1] + "\n"
    driver.get('https://arsenal-na.amazon.com/shipment/encryptDecrypt')
    WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.ID,"shipmentIds")))

    fill_shipments = driver.find_element(By.ID,"shipmentIds")
    fill_shipments.send_keys(ships)

    submit_button = driver.find_element(By.ID,'submitBtn')
    submit_button.click()


    WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"table-result")))
    tables = driver.find_elements(By.CLASS_NAME,"table-result")
    decrypted_table = tables[1]
    elements = decrypted_table.find_elements(By.TAG_NAME,"tr")
    dec_shipments = []
    for element in elements:
        dec_shipments.append(element.text)
    dec_shipments = dec_shipments[1::]
    retorno_shipments = []
    for i in range(len(shipments)):
        retorno_shipments.append((shipments[i][0],shipments[i][1],dec_shipments[i]))
    print(retorno_shipments)
    return(retorno_shipments)    

def gis_tcorp(driver,top,shipments):
        #Save XML folder path
    cwd = str(os.getcwd())
    xml_folder = cwd + "\\xml_folder"
    my_progress2 = ttk.Progressbar(top,orient=HORIZONTAL,
    length=300,mode='determinate')
    my_progress2.pack(pady=10)
    step = 100/len(shipments)
    text = StringVar()
    label1  = Label(top,textvariable=text).pack()
    text.set("Quantidade de Shipments Encontrados: 0/%i"%(len(shipments)))
    top.update_idletasks()
    root.update()
    print(len(shipments))

    with open("retorno.csv","w") as file:
        file.writelines("Ticket,ClientId Domain,Client Id,Status,GiasId,GiasId Type,Client Id Encrypted\n")



    #Start Web Driver
    options = webdriver.FirefoxOptions()

    prefs = {
        'safebrowsing.enabled' : 'false', 
        'download.default_directory' : xml_folder
        }

    iteracoes = 1 + len(shipments)//100
    initial = 0
    ships = ""
    if iteracoes < 2:
        end = len(shipments)
    else:
        end = 99
    top.update_idletasks()
    root.update()
    encontrados = 0
    driver.get('https://gisportal.aka.amazon.com/#/gisv2-batch')
    WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"form-control")))
    time.sleep(10)
    for i in range(iteracoes):
        print("Rodando Batch %i/%i"%(i+1,1+len(shipments)//100))
        forms = driver.find_elements(By.CLASS_NAME,"form-control")
        forms[0].click()
        select = Select(forms[0])
        select.select_by_value("0")
        for ini in range(initial,end):
            print(shipments[ini][2])
            ships = ships + shipments[ini][2] + ","
        forms[1].clear()
        forms[1].send_keys(ships[:len(ships)-1:])
        getStatus = driver.find_element(By.ID,'btn-getStatus')
        getStatus.click()
        time.sleep(3)
        respostas = driver.find_element(By.CLASS_NAME,'table-striped')
        resp = respostas.find_elements(By.CLASS_NAME,"ng-scope")
        encontrados = encontrados + len(resp)
        text.set("Quantidade de Shipments Encontrados: %i/%i"%(encontrados,len(shipments)))
        for r in resp:
            my_progress2['value'] += step
            top.update_idletasks()
            root.update()
            fields = r.find_elements(By.CLASS_NAME,"ng-binding")
            ret = ""
            for field in fields:
                ret = ret +'"'+field.text +'"'+ ","
            ret = ","+ret[:len(ret):]
            with open("retorno.csv","a") as file:
                file.writelines(ret + "\n")
        initial = end
        if(end+99>len(shipments)):
            end = len(shipments)
        else:
            end+=99

        
        ships = ""

    csv = pd.read_csv("retorno.csv",header=0)
    lista_success = csv["Client Id"].values.tolist()
    for item in shipments:
        csv.loc[csv["Client Id"].astype('str').str.contains(item[2]),"Ticket"] = item[0] 
        csv.loc[csv["Client Id"].astype('str').str.contains(item[2]),"Client Id Encrypted"] = item[1] 
    with open("shipments_without_invoice.csv","w") as file:
        for s in shipments:
            print(s)
            if lista_success.count(int(s[2])) < 1:
                file.write("%s,%s,%s\n"%(s[0],s[2],s[1]))
                new_row = {
                    'Ticket':s[0],
                    'Client Id Domain': "",
                    'Client Id': s[2],
                    'Status': "Not Found",
                    "GiasId": "",
                    "GiasId Type": "",
                    "Client Id Encrypted": s[1]
                }
                csv = csv.append(new_row,ignore_index=True)
    erros = pd.read_csv("shipments_without_invoice.csv",names=("Shipments Not Found","Encrypted Shipments Not Found"))
    print(erros)
    driver.close()
    instlot = Label(top,text='Escreva o título do Lote:').pack()
    titlot = Entry(top,width=50)
    titlot.pack()
    butlot = Button(top,text='Nova Pasta',command=lambda:save(top,csv,erros,getTitulo(titlot)))
    butlot.pack()
    butlot2 = Button(top,text='Salvar',command=lambda:sobrescrever(top,csv,erros,getTitulo(titlot)))
    butlot2.pack()
    tup = pyautogui.size()
    wi = (tup[0]/2)-200
    he = (tup[1]/2)-200
    top.geometry('400x550+%i+%i'%(wi,he))

def run_tcorp(root,top):
    text = StringVar()
    label1  = Label(top,textvariable=text).pack()
    text.set("Realizando Login")
    top.update_idletasks()
    root.update()
    #Start Web Driver
    options = webdriver.FirefoxOptions()
    options.set_preference('profile', "/Users/test/Library/Application Support/Firefox/Profiles/default-esr")
    driver = webdriver.Firefox(options=options)
    driver.get('https://t.corp.amazon.com/issues/?q=extensions.tt.status%3A%28Assigned%20OR%20Researching%20OR%20%22Work%20In%20Progress%22%20OR%20Pending%29%20AND%20extensions.tt.assignedGroup%3A%223P%20Integration%20BR%20Support%22')
    print("Please Login")

    username = driver.find_element(By.ID,'user_name_field')
    username.send_keys(credentials.gis['username'])

    time.sleep(2)
    submit_button = driver.find_element(By.ID,'user_name_btn')
    submit_button.click()

    time.sleep(2)
    password = driver.find_element(By.ID,'password_field')
    password.send_keys(credentials.gis['password'])

    login_button = driver.find_element(By.ID,'password_btn')
    login_button.click()

    text.set('Waiting for OTP')
    top.update_idletasks()
    root.update()

    
    WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"sim-list--shortId")))
    text.set('Login done')
    top.update_idletasks()
    root.update()
    tickets = driver.find_elements(By.CLASS_NAME,"sim-list--shortId")
    lista_tickets = []
    shipments = []
    for ticket in tickets:
        print(ticket.text)
        lista_tickets.append(ticket.text)
    text.set('Tickets Encontrados: %i'%len(lista_tickets))
    top.update_idletasks()
    root.update()
    for ticket in lista_tickets:
        driver.get("https://t.corp.amazon.com/%s"%ticket)
        WebDriverWait(driver,timeout=60,poll_frequency=0.5).until(EC.presence_of_element_located((By.CLASS_NAME,"plain-text-display")))
        texto = driver.find_element(By.CLASS_NAME,"plain-text-display")
        try:
            texto = texto.text.replace(" ","")
            texto = texto.replace("*","")
            aux = texto.split("|Affectedshipments/Id(s)da(s)remessa(s):")
            aux2 = aux[1].split("|Numberofaffectedshipments")
            ships = aux2[0].splitlines()
            for ship in ships:
                if ship.count(",")>0:
                    aux = ship.split(",")
                    for item in aux:
                        if item != "" and len(item) ==9:
                            shipments.append((ticket,item))
                else:
                    if ship != "" and len(ship) ==9:
                        shipments.append((ticket,ship))
    
        except:
            pass
    text.set("Decriptando Shipments")
    top.update_idletasks()
    print(shipments)
    retorno = arsenal_tcorp(driver,shipments)
    text.set("Shipments Decriptados")
    top.update_idletasks()
    time.sleep(2)
    print("Aqui")
    gis_tcorp(driver,top,retorno)

def open_tcorp(root):
    tup = pyautogui.size()
    wi = (tup[0]/2)-200
    he = (tup[1]/2)-200
    top66 = Toplevel(root)
    top66.geometry('400x400+%i+%i'%(wi,he))
    top66.title('SACT')
    top66.iconbitmap("amazon-logo.ico")
    label1 = Label(top66,text='Para Shipments no TCORP\nclique no botão').pack()
    botao = Button(top66,text='Começar!',command=lambda: run_tcorp(root,top66)).pack()

def login(top,login,password):
    credentials.gis['username'] = login
    credentials.gis['password'] = password

    top.destroy()
    messagebox.showinfo("SACT","Credenciais Atualizadas")
    root.deiconify() #Unhides the root window

def open_login(root):
    tup = pyautogui.size()
    wi = (tup[0]/2)-200
    he = (tup[1]/2)-200
    top66 = Toplevel(root)
    top66.geometry('400x400+%i+%i'%(wi,he))
    top66.title('SACT')
    top66.iconbitmap("amazon-logo.ico")
    label1 = Label(top66,text='Login').pack()
    username = StringVar()
    usernameEntry = Entry(top66,textvariable=username)
    usernameEntry.insert(0,credentials.gis['username'])
    usernameEntry.pack()
    label2 = Label(top66,text='Password').pack()
    password = StringVar()
    passwordEntry = Entry(top66,textvariable=password,show="♣")
    passwordEntry.insert(0,credentials.gis['password'])
    passwordEntry.pack()
    botao = Button(top66,text='Submit',command=lambda: login(top66,getTitulo(username),getTitulo(password))).pack()
    
    
root= Tk()
open_login(root)
root.geometry('368x270')
root.title('SACT')
root.iconbitmap("amazon-logo.ico")
btn_correio = Button(root,text='Consultar',command=lambda:open_interface(root),height=4,width=35,pady=5).grid(column=1,row=4,padx=5,pady=5)
#btn_tcorp = Button(root,text='Tcorp',command=lambda:open_tcorp(root),height=4,width=35,pady=5).grid(column=1,row=5,padx=5,pady=5)
btn_credencial = Button(root,text='Atualizar Credenciais',command=lambda:open_login(root),height=4,width=35,pady=5).grid(column=1,row=6,padx=5,pady=5)
root.after(1000, update)
root.withdraw()
root.mainloop()


