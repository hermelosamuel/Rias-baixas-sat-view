import json
import os
import traceback
import zipfile
import pandas as pd
from dateutil import relativedelta
from os import listdir
import xlsxwriter
import matplotlib.pyplot as plt
import sys

from sentinelsat import SentinelAPI, read_geojson, geojson_to_wkt
import cv2
from datetime import date
import numpy as np
from pathlib import Path
import re
from glob import glob
from dms2dec.dms_convert import dms2dec
import subprocess
import netCDF4 as nc


import landsatfetch
import tarfile



username = ''
password = ''

fechas = []
data_muestras = []
coordenadasCTDs = []
data_muestras_superf = []

surVigo = -8.938751220703125
norteVigo = -8.598175048828125
oesteVigo = 42.3047373589234
esteVigo = 42.11401459274903

surArousa = -9.021835327148436
norteArousa = -8.739280700683594
oesteArousa = 42.65794930135694
esteArousa = 42.46399280017058

Vigo = True

def listdirs(rootdir):
    for file in os.listdir(rootdir):
        d = os.path.join(rootdir, file)
        if os.path.isdir(d):
            print(d)
            listdirs(d)

def BuscarDatos():
    if Vigo:
        f = open("historicos_Vigo.txt", "r")
    else:
        f = open("historicos_Arousa.txt", "r")
    contents = f.readlines()
    f.close()
    j=1
    correctas=0
    totales=0
    incorrectas=0
    validos=0
    todovalidos=0
    originales=0
    for content in contents:
        if re.compile("^[A-Z][0-9]\s").search(content) or re.compile("^[A-Z]{2}\s").search(content):
            variables = content.split(" ")
            variables = list(filter(None, variables))
            #print(variables)

            for variable in variables:
                variable.split('\t')

                if re.compile("^[A-Z][0-9]$").search(variable) or re.compile("^[A-Z]{2}$").search(variable):
                    ID =variable

                fecha1 = re.findall("[0-3]?[0-9]/[0-3]?[0-9]/(?:[0-9]{2})?[0-9]{2}", variable)
                if fecha1:

                    fecha1 = fecha1[0].split("/")
                    fecha = date((int)(fecha1[2]),(int)(fecha1[1]),(int)(fecha1[0]))

                    if not(fecha in fechas):

                        fechas.append(fecha)
                variable = variable.split('\t')
                variable = list(filter(None, variable))

                if(len(variable) == 11):
                    totales+=5
                    try:
                        if int(variable[10]) < 3:
                            Profundidad = -float(variable[9].replace(',','.'))
                            if int(variable[10]) == 2:
                                correctas +=1
                            if int(variable[10]) == 1:
                                validos +=1
                            if int(variable[10]) == 0:
                                originales += 1
                        else:
                            Profundidad = np.nan
                            incorrectas += 1
                    except:
                        Profundidad = 0


                    try:
                        if int(variable[2]) < 3:
                            Temperatura = float(variable[1].replace(',','.'))
                            if int(variable[2]) == 2:
                                correctas +=1
                            if int(variable[2]) == 1:
                                validos +=1
                            if int(variable[2]) == 0:
                                originales += 1
                        else:
                            Temperatura = np.nan
                            incorrectas+=1
                    except:
                        Temperatura = 0


                    try:
                        if int(variable[4]) < 3:
                            Turbidez = float(variable[3].replace(',','.'))
                            if int(variable[4]) == 2:
                                correctas +=1
                            if int(variable[4]) == 1:
                                validos +=1
                            if int(variable[4]) == 0:
                                originales += 1
                        else:
                            Turbidez = np.nan
                            incorrectas += 1
                    except:
                        Turbidez = 0


                    try:
                        if int(variable[6]) < 3:
                            FluorescenciaUV = float(variable[5].replace(',','.'))
                            if int(variable[6]) == 2:
                                correctas +=1
                            if int(variable[6]) == 1:
                                validos +=1
                            if int(variable[6]) == 0:
                                originales += 1
                        else:
                            FluorescenciaUV = np.nan
                            incorrectas += 1
                    except:
                        FluorescenciaUV = 0

                    try:
                        if int(variable[8]) < 3:
                            Fluorescencia = float(variable[7].replace(',','.'))
                            if int(variable[8]) == 2:
                                correctas +=1
                            if int(variable[8]) == 1:
                                validos += 1
                            if int(variable[8]) == 0:
                                originales += 1
                        else:
                            Fluorescencia = np.nan
                            incorrectas += 1
                    except:
                        Fluorescencia = 0

                    if int(variable[2])  == 1 and int(variable[4])  == 1 and int(variable[6])  == 1 and int(variable[8])  == 1 and int(variable[10])  == 1:
                        todovalidos += 1
                    try:
                        if ID == data_muestras[len(data_muestras)-1][0] and fecha == data_muestras[len(data_muestras)-1][1]:
                            if Profundidad > -2 and data_muestras[len(data_muestras)-1][2] > -2:
                                j+=1
                                deltaP += (float(data_muestras[len(data_muestras)-1][2]) - Profundidad)
                                deltaT += -(float(data_muestras[len(data_muestras)-1][3]) - Temperatura)

                                deltaF += -(float(data_muestras[len(data_muestras)-1][4]) - Fluorescencia)
                                deltaFUV += -(float(data_muestras[len(data_muestras)-1][5]) - FluorescenciaUV)
                                deltaTurb += -(float(data_muestras[len(data_muestras)-1][6]) - Turbidez)
                                mediaT += Temperatura

                                mediaF += Fluorescencia
                                mediaFUV += FluorescenciaUV
                                mediaTurb += Turbidez

                                if float(data_muestras[len(data_muestras)-1][2]) > minP:
                                    minP = float(data_muestras[len(data_muestras)-1][2])
                                    minPT = float(data_muestras[len(data_muestras)-1][3])
                                    minPF = float(data_muestras[len(data_muestras)-1][4])
                                    minPFUV = float(data_muestras[len(data_muestras)-1][5])
                                    minTurb = float(data_muestras[len(data_muestras) - 1][6])
                        else:
                            if j!=1:

                                deltaP /= j-1
                                deltaT /= j-1
                                deltaF /= j-1
                                deltaFUV /= j-1
                                deltaTurb /= j-1
                                mediaT /= j
                                mediaF /= j
                                mediaFUV /= j
                                mediaTurb /= j

                                j=1
                                    # print(str(deltaP) + ","+str(deltaT)+","+str(deltaF)+","+str(deltaFUV)+","+str(minP))
                                prof_aprox = minP
                                temp_aprox = minPT
                                fluor_aprox = minPF
                                fluorUV_aprox = minPFUV
                                turb_aprox = minTurb
                                while prof_aprox < 0:
                                    prof_aprox += deltaP
                                    if prof_aprox >0:
                                        prof_aprox=0

                                    if minPT < 0 and deltaT > 0:
                                        temp_aprox += deltaT
                                    else:
                                        temp_aprox -= deltaT

                                    if minPF < 0 and deltaF > 0:
                                        fluor_aprox += deltaF
                                    else:
                                        fluor_aprox -= deltaF

                                    if minPFUV < 0 and deltaFUV > 0:
                                        fluorUV_aprox += deltaFUV
                                    else:
                                        fluor_aprox -= deltaF

                                    if minPT < 0 and deltaT > 0:
                                        turb_aprox += deltaTurb
                                    else:
                                        turb_aprox -= deltaTurb

                                    data_muestras.append(
                                        [ID, data_muestras[len(data_muestras)-1][1], prof_aprox, temp_aprox, fluor_aprox, fluorUV_aprox, turb_aprox])

                                turb_aprox=100 if turb_aprox > 100 else turb_aprox
                                turb_aprox = 0 if turb_aprox < 0 else turb_aprox

                                mediaTurb=100 if mediaTurb > 100 else mediaTurb
                                mediaTurb = 0 if mediaTurb < 0 else mediaTurb

                                fluor_aprox = 0 if fluor_aprox < 0 else fluor_aprox
                                mediaF = 0 if mediaF < 0 else mediaF

                                fluorUV_aprox = 0 if fluorUV_aprox < 0 else fluorUV_aprox
                                mediaFUV = 0 if mediaFUV < 0 else mediaFUV

                                data_muestras_superf.append(
                                    [ID, data_muestras[len(data_muestras)-1][1], prof_aprox, temp_aprox, fluor_aprox, fluorUV_aprox, turb_aprox, mediaT, mediaF, mediaFUV, mediaTurb])
                                print(ID, data_muestras[len(data_muestras)-1][1], prof_aprox, temp_aprox, fluor_aprox, fluorUV_aprox, turb_aprox, mediaT, mediaF, mediaFUV, mediaTurb)

                                print("ID: " + ID)

                                print("Date: " + data_muestras[len(data_muestras)-1][1])

                                print("Temperature: ", temp_aprox, mediaT)

                                print("Fluorescence: ", fluor_aprox, mediaF)

                                print("FluorescenceUV: ", fluorUV_aprox, mediaFUV)

                                print('***************************************************')

                                minP = -500
                                deltaP = 0
                                deltaT = 0
                                deltaF = 0
                                mediaT = 0
                                mediaF = 0
                                mediaFUV = 0
                                mediaTurb = 0
                                deltaFUV = 0
                                deltaTurb = 0

                    except:

                        minP = -500
                        deltaP = 0
                        deltaT = 0
                        deltaF = 0
                        mediaT = Temperatura
                        mediaF = Fluorescencia
                        mediaFUV = FluorescenciaUV
                        mediaTurb = Turbidez
                        deltaFUV = 0
                        deltaTurb = 0
                    data_muestras.append([ID,fecha,Profundidad,Temperatura,Fluorescencia,FluorescenciaUV,Turbidez])
    fechas.sort()
    data_muestras.sort()
    print("Numero de datos con todos las variables validas = " + str(todovalidos) + " un " + str(todovalidos / (totales/4) * 100) + "%")
    print("Numero de datos validos = " + str(validos) + " un " + str(validos / totales * 100) + "%")
    print("Numero de datos correctos = "+ str(correctas) +" un " + str(correctas/totales*100) + "%")
    print("Numero de datos incorrectos = " + str(incorrectas) + " un " + str(incorrectas / totales * 100) + "%")
    print("Numero de datos originales = " + str(originales) + " un " + str(originales / totales * 100) + "%")
    print(totales)

def BuscarCoord():
    if Vigo:
        df = pd.read_excel('coordenadas_de_las_estaciones_oceanográficas.xlsx',usecols="B,I:J",skiprows=[0],nrows=8)
    else:
        df = pd.read_excel('coordenadas_de_las_estaciones_oceanográficas.xlsx', usecols="B,I:J", skiprows=25, nrows=9)
    for d in df.values:
        nombre=d[0]
        dd0=dms2dec(d[1])
        dd1=dms2dec(d[2])
        dd1*=-1
        coordenadasCTDs.append([nombre,dd0,dd1])
        #print([nombre,dd0,dd1])

def LandsatFetch(username,password):

    f = open("scenes.txt", "w")
    f.write('landsat_ot_c2_l1|displayId'+"\n")
    f.close()
    f = open("scenes.txt", "a")
    for fecha in fechas:
        diasiguiente = fecha+relativedelta.relativedelta(days=1)
        searchrequest = "landsatxplore search --dataset landsat_ot_c2_l1 --location 42.27 -8.80 --start "+str(fecha.year)+"-"+ str(fecha.month)+"-"+str(fecha.day)+" --end "+str(diasiguiente.year)+"-"+ str(diasiguiente.month)+"-"+str(diasiguiente.day)+" --username "+username+" --password "+password
        scenes = subprocess.Popen(searchrequest, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

        for line in scenes.stdout.readlines():

            if str(line) != 'b\'\\x1b[0m\'':
                print(line)
                f.writelines(str(line).replace('b\'','').replace('\\r','').replace('\\n','').replace('\'','')+"\n")
    f.close()
    if Vigo:
        landsatfetch.solicitardescarga(username,password,'bundle','/Vigo/')
    else:
        landsatfetch.solicitardescarga(username, password, 'bundle', '/Arousa/')

def LandsatExtract():
    if Vigo:
        zips = glob(os.getcwd() + "/Vigo/*.tar")
        Imagenes = glob(os.getcwd() + "/Vigo/*.tif")
        Imagenes.extend(os.getcwd() + "/Vigo/*.jpg")
    else:
        zips = glob(os.getcwd() + "/Arousa/*.tar")
        Imagenes = glob(os.getcwd() + "/Arousa/*.tif")
        Imagenes.extend(os.getcwd() + "/Arousa/*.jpg")
    for zip in zips:
        try:
            print("Extracting: "+zip)
            if not os.path.exists(zip.split('.')[0]):
                os.mkdir(zip.split('.')[0])
            for Imagen in Imagenes:
                nombreImagen = Imagen.split("\\")[len(Imagen.split("\\"))-1]
                #print(nombreImagen)
                if zip.split('.')[0] in Imagen:
                    os.rename(Imagen,zip.split(".")[0]+"/"+nombreImagen)
            try:
                with tarfile.open(zip) as zip_ref:
                    zip_ref.extractall(zip.split('.')[0])
                    zip_ref.close()
            except:
                os.rename(Imagen, zip.split(".")[0] + "/" + nombreImagen)
                #os.remove(Imagen)
                zip_ref.close()
        except:
            print("Archivo corrupto")

def normalize8(I):
    I[np.isnan(I)]=0
    if len(I[I != 0])>0:
        mn = I[I!=0].min()
        mx = I.max()
    else:
        mn = I.min()
        mx = I.max()
    #print("Chl-a range: "+ str(mx)+','+str(mn))
    mx -= mn

    I = ((I - mn)/mx) * 255
    return I.astype(np.uint8)

def AcoliteProcess():
    if Vigo:
        rootdir = str(os.getcwd())+'/Vigo'
    else:
        rootdir = str(os.getcwd())+'/Arousa'
    for it in os.scandir(rootdir):
        if it.is_dir():
            f = open("settings.txt", "w")
            f.write("inputfile="+it.path+'\n')
            if Vigo:
                out = os.getcwd() + '/Imagenes/Vigo/'+it.path.split('\\')[len(it.path.split('\\'))-1]
            else:
                out = os.getcwd() + '/Imagenes/Arousa/' + it.path.split('\\')[len(it.path.split('\\')) - 1]
            if not os.path.isdir(out):
                os.mkdir(out)
            f.write("output="+out+'\n')
            f.write('EARTHDATA_u=hermelosamuel\n')
            f.write('EARTHDATA_p=Riasatview1234\n')
            if(Vigo):
                f.write('polygon='+str(os.getcwd())+'/mapVigo.geojson'+'\n')
            else:
                f.write('polygon=' + str(os.getcwd()) + '/mapArousa.geojson' + '\n')
            f.write("l2w_parameters=chl_oc3,tur_nechad2016,chl_re_bramich,bt10,bt11,chl_re_mishra,chl_re_moses3b,Rrs_*"+'\n')
            f.write("rgb_rhot=True"+'\n')
            f.write("rgb_rhos=True"+'\n')
            f.write("map_l2w=True"+'\n')
            f.close()
            print("|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||")
            p = subprocess.Popen(['acolite_py_win\dist\\acolite\\acolite.exe', '--cli', '--settings=settings.txt', ], stdout=subprocess.PIPE, bufsize=1)
            for line in iter(p.stdout.readline, b''):
                print(line)
            p.stdout.close()
            p.wait()

def snapc2(mostrarImagenes,guardarImagenes):
    bands=[]
    comp_oc3=0
    comp_mishra=0
    comp_bramich=0
    comp_remoses=0
    comp_tur=0
    comp_bt10=0
    comp_bt11=0
    if Vigo:
        workbook = xlsxwriter.Workbook('ComparisonVigo.xlsx')
    else:
        workbook = xlsxwriter.Workbook('ComparisonArousa.xlsx')
    worksheet = workbook.add_worksheet()
    rows = 1
    if Vigo:
        pathroot=os.getcwd() + "/Imagenes/Vigo/"
    else:
        pathroot = os.getcwd() + "/Imagenes/Arousa/"
    for path in Path(pathroot).rglob("*"):
        fns = glob(str(path)+"/*L2W.nc")
        for fn in fns:
            #if 'L2W' in fn:
            ds = nc.Dataset(fn)
            lat = ds.variables['lat'][:]
            lon = ds.variables['lon'][:]
            derecha = np.amin(lon)
            abajo = np.amin(lat)
            izquierda = np.amax(lon)
            arriba = np.amax(lat)
            fname = os.path.basename(fn)
            fecha_sat = date(int(fname.split('_')[2]), int(fname.split('_')[3]), int(fname.split('_')[4]))
            chl_oc3 = ds.variables['chl_oc3'][:]
            chl_oc3_image = chl_oc3.copy()
            chl_oc3_image = np.multiply(chl_oc3_image, 255.0)
            chl_oc3_image = chl_oc3_image.astype(np.uint8)


            # RGB_gt = np.divide(RGB_gt, max_pixel_value)
            chl_oc3_image = cv2.applyColorMap(chl_oc3_image, cv2.COLORMAP_WINTER)

            chl_oc3_image[np.isnan(chl_oc3)] = 127
            try:
                chl_re_moses3b = ds.variables['chl_re_moses3b'][:]
                chl_re_moses3b_image = chl_re_moses3b.copy()
                chl_re_moses3b_image = cv2.applyColorMap(np.uint8(chl_re_moses3b_image), cv2.COLORMAP_WINTER)
                chl_re_moses3b_image[np.isnan(chl_re_moses3b)] = 127
            except:
                chl_re_moses3b = np.nan

            try:
                chl_re_mishra = ds.variables['chl_re_mishra'][:]
                chl_re_mishra_image = chl_re_mishra.copy()
                chl_re_mishra_image = cv2.applyColorMap(np.uint8(chl_re_mishra_image), cv2.COLORMAP_WINTER)
                chl_re_mishra_image[np.isnan(chl_re_mishra)] = 127
            except:
                chl_re_mishra = np.nan

            try:
                chl_re_bramich = ds.variables['chl_re_bramich'][:]
                chl_re_bramich_image = chl_re_bramich.copy()
                chl_re_bramich_image = cv2.applyColorMap(np.uint8(chl_re_bramich_image), cv2.COLORMAP_WINTER)
                chl_re_bramich_image[np.isnan(chl_re_bramich)] = 127
            except:
                chl_re_bramich = np.nan

            try:
                bt10 = ds.variables['bt10'][:]
                bt10_image = bt10.copy()
                bt10_image = cv2.applyColorMap(np.uint8(bt10_image), cv2.COLORMAP_WINTER)
                bt10_image[np.isnan(bt10)] = 127
            except:
                bt10 = np.nan
            try:
                bt11 = ds.variables['bt11'][:]
                bt11_image = bt11.copy()
                bt11_image = cv2.applyColorMap(np.uint8(bt11_image), cv2.COLORMAP_WINTER)
                bt11_image[np.isnan(bt11)] = 127
            except:
                bt11 = np.nan

            Rrs = []
            try:
                Rrs.append(ds.variables['Rrs_443'][:])
                if not '443' in bands:
                    bands.append('443')
            except:
                Rrs.append('-')
            try:
                Rrs.append(ds.variables['Rrs_483'][:])
                if not '483' in bands:
                    bands.append('483')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_482'][:])
                if not '482' in bands:
                    bands.append('482')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_492'][:])
                if not '492' in bands:
                    bands.append('492')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_561'][:])
                if not '561' in bands:
                    bands.append('561')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_560'][:])
                if not '560' in bands:
                    bands.append('560')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_559'][:])
                if not '559' in bands:
                    bands.append('559')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_655'][:])
                if not '655' in bands:
                    bands.append('655')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_654'][:])
                if not '654' in bands:
                    bands.append('654')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_665'][:])
                if not '665' in bands:
                    bands.append('665')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_865'][:])
                if not '865' in bands:
                    bands.append('865')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_704'][:])
                if not '704' in bands:
                    bands.append('704')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_740'][:])
                if not '740' in bands:
                    bands.append('740')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_1609'][:])
                if not '1609' in bands:
                    bands.append('1609')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_1608'][:])
                if not '1608' in bands:
                    bands.append('1608')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_1614'][:])
                if not '1614' in bands:
                    bands.append('1614')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_2201'][:])
                if not '2201' in bands:
                    bands.append('2201')
            except:
                Rrs.append('-')
            try:
                Rrs.append(ds.variables['Rrs_2202'][:])
                if not '2202' in bands:
                    bands.append('2202')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_2186'][:])
                if not '2186' in bands:
                    bands.append('2186')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_592'][:])
                if not '592' in bands:
                    bands.append('592')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_594'][:])
                if not '594' in bands:
                    bands.append('594')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_783'][:])
                if not '783' in bands:
                    bands.append('783')
            except:
                Rrs.append('-')

            try:
                Rrs.append(ds.variables['Rrs_613'][:])
                if not '613' in bands:
                    bands.append('613')
            except:
                Rrs.append('-')
            try:
                Rrs.append(ds.variables['Rrs_833'][:])
                if not '833' in bands:
                    bands.append('833')
            except:
                Rrs.append('-')


            chl_image = chl_oc3.copy()
            chl_image = normalize8(chl_image)
            #chl_image = cv2.applyColorMap(chl_image, cv2.COLORMAP_WINTER)
            #chl_image[chl == float('nan')] = 127
            try:
                turb = ds.variables['TUR_Nechad2016_655'][:]
            except:
                try:
                    turb = ds.variables['TUR_Nechad2016_654'][:]
                except:
                    turb = ds.variables['TUR_Nechad2016_665'][:]
            turb_image = turb.copy()

            turb_image = cv2.applyColorMap(np.uint8(turb_image), cv2.COLORMAP_WINTER)
            cv2.applyColorMap(np.uint8(turb_image), cv2.COLORMAP_WINTER)
            turb_image[np.isnan(turb)] = 127

                #turb = np.empty((chl.shape[0],chl.shape[0],))
                #turb[:] = np.nan
            dx = lon.shape[1] / (lon[0][0] - lon[0][len(lon[0])-1])
            dy = lat.shape[0] / (abajo-arriba)

            if Vigo:
                ria_topleft = [int(chl_oc3.shape[1] + (norteVigo - izquierda) * dx),
                               int((oesteVigo - arriba) * dy)]
                ria_bottomleft = [int(chl_oc3.shape[1] + (surVigo - izquierda) * dx),
                                  int((esteVigo - arriba) * dy)]

                ria_topright = [int(chl_oc3.shape[1] + (norteVigo - izquierda) * dx),
                                int((oesteVigo - arriba) * dy)]
                ria_bottomright = [int(chl_oc3.shape[1] + (surVigo - izquierda) * dx),
                                   int((esteVigo - arriba) * dy)]
            else:
                ria_topleft = [int(chl_oc3.shape[1] + (norteArousa - izquierda) * dx),
                               int((oesteArousa - arriba) * dy)]
                ria_bottomleft = [int(chl_oc3.shape[1] + (surArousa - izquierda) * dx),
                                  int((esteArousa - arriba) * dy)]

                ria_topright = [int(chl_oc3.shape[1] + (norteArousa - izquierda) * dx),
                                int((oesteArousa - arriba) * dy)]
                ria_bottomright = [int(chl_oc3.shape[1] + (surArousa - izquierda) * dx),
                                   int((esteArousa - arriba) * dy)]
            if "L8" in fn:
                print('Landsat 8 at '+ str(fecha_sat))
            if "L9" in fn:
                print('Landsat 9 at '+ str(fecha_sat))
            if "S2A" in fn or 'S2B' in fn:
                print('Sentinel 2 at '+ str(fecha_sat))

            for coorCTD in coordenadasCTDs:
                '''
                cv2.circle(chl_image,
                           [int(chl.shape[1] + (coorCTD[2] - derecha) * dx),
                            int((coorCTD[1] - arriba) * dy)], 10, 100, 4)
                cv2.putText(chl_image, coorCTD[0],
                            [int(chl.shape[1] + (coorCTD[2] - derecha) * dx),
                             int((coorCTD[1] - arriba) * dy)],
                            cv2.FONT_HERSHEY_SIMPLEX, 3, 100, 4)
                '''
                #print(int(chl_oc3.shape[1] - (coorCTD[2] - izquierda) * dx),',',int((coorCTD[1] - arriba) * dy))
                try:
                    satChl_oc3= chl_oc3[int((coorCTD[1] - arriba) * dy),int(chl_oc3.shape[1] - (coorCTD[2] - izquierda) * dx)]

                    satChl_oc3 = np.nan if satChl_oc3 > 100 else satChl_oc3
                    cv2.circle(chl_oc3_image, [int(chl_oc3_image.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)]
                               , 20, (0, 255, 255), 10)
                    cv2.putText(chl_oc3_image, coorCTD[0],
                                [int(chl_oc3_image.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)],
                                cv2.FONT_HERSHEY_SIMPLEX, 1+int(chl_oc3_image.shape[1]/1000*1.5), (0, 255, 255), 4)
                except:
                    satChl_oc3=np.nan

                try:
                    satChl_re_bramich= chl_re_bramich[int((coorCTD[1] - arriba) * dy),int(chl_re_bramich.shape[1] - (coorCTD[2] - izquierda) * dx)]
                    satChl_re_bramich = np.nan if satChl_re_bramich > 100 else satChl_re_bramich
                    cv2.circle(chl_re_bramich_image,[int(chl_re_bramich_image.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)]
                               , 10, (0, 255, 255), 4)
                    cv2.putText(chl_re_bramich_image, coorCTD[0],
                                [int(chl_re_bramich_image.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)],
                                cv2.FONT_HERSHEY_SIMPLEX, 1+int(chl_re_bramich_image.shape[1]/1000*1.5), (0, 255, 255), 4)
                except:
                    satChl_re_bramich=np.nan

                try:
                    satChl_re_moses3b= chl_re_moses3b[int((coorCTD[1] - arriba) * dy),int(chl_re_moses3b.shape[1] - (coorCTD[2] - izquierda) * dx)]
                    satChl_re_moses3b = np.nan if satChl_re_moses3b > 100 else satChl_re_moses3b
                    cv2.circle(chl_re_moses3b_image,[int(chl_re_moses3b.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)]
                               , 10, (0, 255, 255), 4)
                    cv2.putText(chl_re_moses3b_image, coorCTD[0],
                                [int(chl_re_moses3b.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)],
                                cv2.FONT_HERSHEY_SIMPLEX, 1+int(chl_re_moses3b_image.shape[1]/1000*1.5), (0, 0, 255), 4)
                except:
                    satChl_re_moses3b=np.nan

                try:
                    satChl_re_mishra= chl_re_mishra[int((coorCTD[1] - arriba) * dy),int(chl_re_mishra.shape[1] - (coorCTD[2] - izquierda) * dx)]
                    satChl_re_mishra = np.nan if satChl_re_mishra > 100 else satChl_re_mishra
                    cv2.circle(chl_re_mishra_image,[int(chl_re_mishra.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)]
                               , 10, (0, 255, 255), 4)
                    cv2.putText(chl_re_mishra_image, coorCTD[0],
                                [int(chl_re_mishra.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)],
                                cv2.FONT_HERSHEY_SIMPLEX, 1+int(chl_re_mishra_image.shape[1]/1000*1.5), (0, 0, 255), 4)
                except:
                    satChl_re_mishra=np.nan


                try:
                    satbt10= bt10[int((coorCTD[1] - arriba) * dy),int(bt10.shape[1] - (coorCTD[2] - izquierda) * dx)]
                    cv2.circle(bt10_image,[int(bt10.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)], 10, (0, 255, 255), 4)
                    cv2.putText(bt10_image, coorCTD[0],
                                [int(bt10.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)],
                                cv2.FONT_HERSHEY_SIMPLEX, 1+int(bt10_image.shape[1]/1000*1.5), (0, 255, 255), 4)
                    if np.isnan(satChl_oc3):
                        satbt10 = np.nan

                except:
                    satbt10=np.nan

                try:
                    satbt11 = bt11[int((coorCTD[1] - arriba) * dy),int(bt11.shape[1] - (coorCTD[2] - izquierda) * dx)]
                    cv2.circle(bt11_image,[int(bt11.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)], 10, (0, 255, 255), 4)
                    cv2.putText(bt11_image, coorCTD[0],
                                [int(bt11.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)],
                                cv2.FONT_HERSHEY_SIMPLEX, 1+int(bt11_image.shape[1]/1000*1.5), (0, 255, 255), 4)
                    if np.isnan(satChl_oc3):
                        satbt11 = np.nan
                except:
                    satbt11 = np.nan

                try:
                    satsst = satbt10 + (2.946*(satbt10 - satbt11)) - 0.038
                except:
                    satsst = np.nan
                try:
                    saturb = turb[int((coorCTD[1] - arriba) * dy),int(turb.shape[1] - (coorCTD[2] - izquierda) * dx)]

                    cv2.circle(turb_image,[int(turb.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)], 10, (0, 255, 255), 4)
                    cv2.putText(turb_image, coorCTD[0],
                                [int(turb.shape[1] - (coorCTD[2] - izquierda) * dx),int((coorCTD[1] - arriba) * dy)],
                                cv2.FONT_HERSHEY_SIMPLEX, int(turb_image.shape[1]/1000)*2, (0, 255, 255), 4)
                except:
                    saturb=np.nan
                if mostrarImagenes:
                    cv2.imshow('chl_oc3',chl_oc3_image)
                    try:
                        cv2.imshow('chl_re_bramich',chl_re_bramich_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imshow('chl_re_mishra',chl_re_mishra_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imshow('chl_re_moses3b',chl_re_moses3b_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imshow('bt10',bt10_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imshow('bt11',bt11_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imshow('turb',turb_image)
                    except:
                        print('',end='')
                    cv2.waitKey(0)
                if guardarImagenes:
                    filename = str(path).split('\\')[len(str(path).split('\\'))-1]
                    cv2.imwrite(str(path) + '\\'+filename+'chl_oc3.jpg', chl_oc3_image)
                    try:
                        cv2.imwrite(str(path) + '\\'+filename+'chl_re_bramich.jpg', chl_re_bramich_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imwrite(str(path) + '\\'+filename+'chl_re_mishra.jpg', chl_re_mishra_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imwrite(str(path) + '\\'+filename+'chl_re_moses3b.jpg', chl_re_moses3b_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imwrite(str(path) + '\\'+filename+ 'bt10.jpg', bt10_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imwrite(str(path) + '\\'+filename+'bt11.jpg', bt11_image)
                    except:
                        print('',end='')
                    try:
                        cv2.imwrite(str(path) + '\\'+filename+'turb.jpg', turb_image)
                    except:
                        print('',end='')

                worksheet.write('A1', 'Date')
                worksheet.write('B1', 'Station')
                worksheet.write('C1', 'Latitude')
                worksheet.write('D1', 'Longitude')
                worksheet.write('E1', 'oc3')
                worksheet.write('F1', 're_mishra')
                worksheet.write('G1', 're_bramich')
                worksheet.write('H1', 're_moses3b')
                worksheet.write('I1', 'in-situ chloro')
                worksheet.write('J1', 'in-situ chloro average')
                worksheet.write('K1', 'tur_nechad2016')
                worksheet.write('L1', 'in-situ turbidity')
                worksheet.write('M1', 'in-situ turbidity average')
                worksheet.write('N1', 'bt10')
                worksheet.write('O1', 'bt11')
                worksheet.write('P1', 'SST')
                worksheet.write('Q1', 'in-situ temperature')
                worksheet.write('R1', 'in-situ temperature average')



                i=18
                for muestra in data_muestras_superf:
                    columns=0
                    if muestra[0] == coorCTD[0] and muestra[1] == fecha_sat:
                        worksheet.write(rows, columns, str(fecha_sat))
                        print(str(fecha_sat), end='\t')
                        columns+=1
                        worksheet.write(rows, columns, str(coorCTD[0]))
                        print((coorCTD[0]),end='\t')
                        columns+=1
                        worksheet.write(rows,columns,str(coorCTD[1]))
                        print("{:.4f}".format(coorCTD[1]),end='\t')
                        columns+=1
                        worksheet.write(rows,columns,str(coorCTD[2]))
                        print("{:.4f}".format(coorCTD[2]),end='\t')
                        columns+=1

                        if not np.isnan(satChl_oc3):
                            worksheet.write(rows, columns, (satChl_oc3))
                            print('\t\t' + "{:.4f}".format(satChl_oc3), end='')
                            comp_oc3 += 1
                            worksheet.write(0, i, 'OC_3 Comparison Deviation')
                            worksheet.write(rows, i, (satChl_oc3-muestra[4]))
                            i+=1
                            worksheet.write(0, i, 'OC_3 Comparison Average')
                            worksheet.write(rows, i, (satChl_oc3-muestra[8]))
                            i+=1
                        else:
                            worksheet.write(rows,columns,'-')
                            print('\t\t\t-', end='')
                            i += 2
                        columns+=1

                        if not np.isnan(satChl_re_mishra):
                            worksheet.write(rows,columns,(satChl_re_mishra))
                            print('\t\t' + "{:.4f}".format(satChl_re_mishra), end='')
                            comp_mishra+=1
                            worksheet.write(0, i, 'Mishra Comparison Deviation')
                            worksheet.write(rows, i, (satChl_re_mishra-muestra[4]))
                            i+=1
                            worksheet.write(0, i, 'Mishra Comparison Average')
                            worksheet.write(rows, i, (satChl_re_mishra-muestra[8]))
                            i+=1

                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-',end='')
                            i += 2
                        columns+=1

                        if not np.isnan(satChl_re_bramich):
                            worksheet.write(rows, columns, (satChl_re_bramich))
                            print('\t\t' + "{:.4f}".format(satChl_re_bramich), end='')
                            comp_bramich+=1
                            worksheet.write(0, i, 'Bramich Comparison Deviation')
                            worksheet.write(rows, i, (satChl_re_bramich-muestra[4]))
                            i+=1
                            worksheet.write(0, i, 'Bramich Comparison Average')
                            worksheet.write(rows, i, (satChl_re_bramich-muestra[8]))
                            i+=1

                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')
                            i += 2
                        columns+=1

                        if not np.isnan(satChl_re_moses3b):
                            worksheet.write(rows, columns, (satChl_re_moses3b))
                            print('\t\t' + "{:.4f}".format(satChl_re_moses3b), end='')
                            comp_remoses+=1
                            worksheet.write(0, i, 'Moses3B Comparison Deviation')
                            worksheet.write(rows, i, (satChl_re_moses3b-muestra[4]))
                            i+=1
                            worksheet.write(0, i, 'Moses3B Comparison Average')
                            worksheet.write(rows, i, (satChl_re_moses3b-muestra[8]))
                            i+=1
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')
                            i += 2
                        columns+=1

                        if not np.isnan(muestra[4]):
                            worksheet.write(rows, columns, (muestra[4]))
                            print('\t\t' + "{:.4f}".format(muestra[4]), end='')
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')

                        columns+=1

                        if not np.isnan(muestra[8]):
                            worksheet.write(rows, columns, (muestra[8]))
                            print('\t\t' + "{:.4f}".format(muestra[8]), end='')
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')

                        columns+=1

                        if not np.isnan(saturb):
                            worksheet.write(rows, columns, (saturb))
                            print('\t\t' + "{:.4f}".format(saturb),end='')
                            comp_tur+=1
                            worksheet.write(0, i, 'Turb Comparison Deviation')
                            worksheet.write(rows, i, (saturb-(100-muestra[6])))
                            i+=1
                            worksheet.write(0, i, 'Turb Comparison Average')
                            worksheet.write(rows, i, (saturb-(100-muestra[10])))
                            i+=1
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-',end='')
                            i += 2
                        columns+=1

                        if not np.isnan(muestra[6]):
                            worksheet.write(rows, columns, (100-muestra[6]))
                            print('\t\t' + "{:.4f}".format(100-muestra[6]), end='')
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')

                        columns+=1


                        if not np.isnan(muestra[10]):
                            worksheet.write(rows, columns, (100-muestra[10]))
                            print('\t\t' + "{:.4f}".format(100-muestra[10]), end='')
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')

                        columns+=1

                        if not np.isnan(satbt10) and not np.isnan(muestra[3]) and not np.isnan(muestra[7]):
                            worksheet.write(rows, columns, (satbt10-273.15))
                            print('\t\t' + "{:.4f}".format(satbt10-273.15),end='')
                            comp_bt10+=1
                            worksheet.write(0, i, 'bt10 Comparison Deviation')
                            worksheet.write(rows, i, (satbt10-273.15-muestra[3]))
                            i+=1
                            worksheet.write(0, i, 'bt10 Comparison Average')
                            worksheet.write(rows, i, (satbt10-273.15-muestra[7]))
                            i+=1
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-',end='')
                            i += 2
                        columns+=1

                        if not np.isnan(satbt11) and not np.isnan(muestra[3]) and not np.isnan(muestra[7]):
                            worksheet.write(rows, columns, (satbt11-273.15))
                            print('\t\t' + "{:.4f}".format(satbt11-273.15), end='')
                            comp_bt11+=1
                            worksheet.write(0, i, 'bt11 Comparison Deviation')
                            worksheet.write(rows, i, (satbt11-273.15-muestra[3]))
                            i+=1
                            worksheet.write(0, i, 'bt11 Comparison Average')
                            worksheet.write(rows, i, (satbt11-273.15-muestra[7]))
                            i+=1

                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')
                            i += 2
                        columns+=1
                        if not np.isnan(satsst) and not np.isnan(muestra[3]) and not np.isnan(muestra[7]):
                            worksheet.write(rows, columns, (satsst-273.15))
                            print('\t\t' + "{:.4f}".format(satsst-273.15), end='')
                            comp_bt11+=1
                            worksheet.write(0, i, 'sst Comparison Deviation')
                            worksheet.write(rows, i, (satsst-273.15-muestra[3]))
                            i+=1
                            worksheet.write(0, i, 'sst Comparison Average')
                            worksheet.write(rows, i, (satsst-273.15-muestra[7]))
                            i+=1

                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')
                            i += 2
                        columns+=1
                        if not np.isnan(muestra[3]):
                            worksheet.write(rows, columns, (muestra[3]))
                            print('\t\t' + "{:.4f}".format(muestra[4]), end='')
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')
                        columns+=1

                        if not np.isnan(muestra[7]):
                            worksheet.write(rows, columns, (muestra[7]))
                            print('\t\t' + "{:.4f}".format(muestra[7]), end='')
                        else:
                            worksheet.write(rows, columns, '-')
                            print('\t\t\t-', end='')
                        columns=i

                        for band in bands:
                            worksheet.write(0, i, 'Rrs_' + str(band))
                            i += 1
                        for R in Rrs:
                            try:
                                satband = R[int(R.shape[1] - (coorCTD[2] - izquierda) * dx),
                                              int((coorCTD[1] - arriba) * dy)]
                            except:
                                satband = np.nan
                            if not np.isnan(satband):
                                worksheet.write(rows, columns, "{:.4f}".format(satband))
                                print('\t\t' + "{:.4f}".format(satband),end='')

                            else:
                                worksheet.write(rows, columns, '-')
                                print('\t\t\t-',end='')

                            columns+=1
                        print('')
                        rows+=1

                            #print('Chl: ' + str(muestra[4]) + " - " + str(satChl_oc3))
            ria_topleft[0] = ria_topleft[0] if ria_topleft[0] > 0 else 0
            ria_topleft[1] = ria_topleft[1] if ria_topleft[1] > 0 else 0
            ria_topright[0] = ria_topright[0] if ria_topright[0] > 0 else 0
            ria_topright[1] = ria_topright[1] if ria_topright[1] < chl_image.shape[1] else chl_image.shape[1]
            ria_bottomleft[0] = ria_bottomleft[0] if ria_bottomleft[0] < chl_image.shape[0] else chl_image.shape[0]
            ria_bottomleft[1] = ria_topleft[1] if ria_topleft[1] > 0 else 0
            ria_bottomright[0] = ria_bottomright[0] if ria_bottomright[0] < chl_image.shape[0] else chl_image.shape[0]
            ria_bottomright[1] = ria_bottomright[1] if ria_bottomright[1] < chl_image.shape[0] else chl_image.shape[0]
            cv2.rectangle(chl_image, ria_topleft, ria_bottomright, 200, thickness=5)

    try:
        worksheet2 = workbook.add_worksheet('Chlorophile comparisons')
    except:
        worksheet2 = workbook.get_worksheet_by_name("Chlorophile comparisons")


    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$S$2:$S$'+str(rows),
                      'categories': '=Sheet1!$S$1',
                      'name': 'Average'})
    chart.set_style(11)
    #chart.add_series({'values': '=Sheet1!$S$2:$S$' + str(rows)})
    chart.add_series({'values': '=Sheet1!$T$2:$T$'+str(rows),
                      'categories': '=Sheet1!$T$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet2.insert_chart('A1', chart)

    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$U$2:$U$'+str(rows),
                      'categories': '=Sheet1!$U$1',
                      'name': 'Average'})
    chart.set_style(11)
    #chart.add_series({'values': '=Sheet1!$S$2:$S$' + str(rows)})
    chart.add_series({'values': '=Sheet1!$V$2:$V$'+str(rows),
                      'categories': '=Sheet1!$V$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet2.insert_chart('O1', chart)

    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$W$2:$W$'+str(rows),
                      'categories': '=Sheet1!$W$1',
                      'name': 'Average'})
    chart.set_style(11)
    #chart.add_series({'values': '=Sheet1!$S$2:$S$' + str(rows)})
    chart.add_series({'values': '=Sheet1!$X$2:$X$'+str(rows),
                      'categories': '=Sheet1!$X$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet2.insert_chart('A27', chart)

    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$Y$2:$Y$'+str(rows),
                      'categories': '=Sheet1!$Y$1',
                      'name': 'Average'})
    chart.set_style(11)
    #chart.add_series({'values': '=Sheet1!$S$2:$S$' + str(rows)})
    chart.add_series({'values': '=Sheet1!$Z$2:$Z$'+str(rows),
                      'categories': '=Sheet1!$Z$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet2.insert_chart('O27', chart)

    try:
        worksheet3 = workbook.add_worksheet('Other comparisons')
    except:
        worksheet3 = workbook.get_worksheet_by_name("Other comparisons")


    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$AA$2:$AA$'+str(rows),
                      'categories': '=Sheet1!$AA$1',
                      'name': 'Average'})
    chart.set_style(11)
    chart.add_series({'values': '=Sheet1!$AB$2:$AB$'+str(rows),
                      'categories': '=Sheet1!$AB$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet3.insert_chart('A1', chart)

    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$AC$2:$AC$'+str(rows),
                      'categories': '=Sheet1!$AC$1',
                      'name': 'Average'})
    chart.set_style(11)
    #chart.add_series({'values': '=Sheet1!$S$2:$S$' + str(rows)})
    chart.add_series({'values': '=Sheet1!$AD$2:$AD$'+str(rows),
                      'categories': '=Sheet1!$AD$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet3.insert_chart('O1', chart)

    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$AE$2:$AE$'+str(rows),
                      'categories': '=Sheet1!$AE$1',
                      'name': 'Average'})
    chart.set_style(11)
    #chart.add_series({'values': '=Sheet1!$S$2:$S$' + str(rows)})
    chart.add_series({'values': '=Sheet1!$AF$2:$AF$'+str(rows),
                      'categories': '=Sheet1!$AF$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet3.insert_chart('A27', chart)

    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({'values': '=Sheet1!$AG$2:$AG$'+str(rows),
                      'categories': '=Sheet1!$AG$1',
                      'name': 'Average'})
    chart.set_style(11)
    #chart.add_series({'values': '=Sheet1!$S$2:$S$' + str(rows)})
    chart.add_series({'values': '=Sheet1!$AH$2:$AH$'+str(rows),
                      'categories': '=Sheet1!$AH$1',
                      'name': 'Deviation'})
    chart.set_style(12)
    worksheet3.insert_chart('O27', chart)
    print('---------------------------------------------------')
    plt.show()
    workbook.close()
    print("chl_oc3 comparisons: ",comp_oc3)
    print("re_moses3b comparisons: ", comp_remoses)
    print("chl_mishra comparisons: ",comp_mishra)
    print("bramich comparisons: ", comp_bramich)
    print("turbidity comparisons: ",comp_tur)
    print("bt10 comparisons: ", comp_bt10)
    print("bt11 comparisons: ", comp_bt11)
    print()
            # cv2.imshow('image', cv2.resize(RGB_gt, (int(img.shape[1] * 12 / 100),int(img.shape[0] * 12 / 100)), interpolation = cv2.INTER_AREA))
            # cv2.waitKey(0)

            #cv2.imshow('Chl-a', cv2.resize(chl_image, (int(chl.shape[1] * 12 / 100),int(chl.shape[0] * 12 / 100)), interpolation = cv2.INTER_AREA))
            #cv2.imshow('Turb', cv2.resize(turb, (int(turb.shape[1] * 12 / 100),int(turb.shape[0] * 12 / 100)), interpolation = cv2.INTER_AREA))
            #cv2.waitKey(0)

###################### SENTINEL FETCH #########################
def SentinelFetch():
    if Vigo:
        footprint = geojson_to_wkt(read_geojson('mapVigo.geojson'))
    else:
        footprint = geojson_to_wkt(read_geojson('mapArousa.geojson'))
    sentinelapi.lta_timeout = 300
    #print(fechas)
    for fecha in fechas:
        try:
            print(fecha)
            products = sentinelapi.query(footprint,
                             date=(fecha, fecha+relativedelta.relativedelta(days=1)),
                             platformname='Sentinel-2',
                             cloudcoverpercentage=(0, 30),
                             producttype='S2MSI1C',)
            if Vigo:
                sentinelapi.download_all(products, os.getcwd() +"/Vigo")
            else:
                sentinelapi.download_all(products, os.getcwd() + "/Arousa")
        except:
            print(traceback.print_exc())

    ##Cuando acabe de descargar
    if(Vigo):
        zips = listdir(os.getcwd() + "/Vigo/")
    else:
        zips = listdir(os.getcwd() + "/Arousa/")
    print(zips)
    for zip in zips:
        if '.zip' in zip:
            if Vigo:
                with zipfile.ZipFile(os.getcwd()+"/Vigo/"+zip, 'r') as zip_ref:
                    zip_ref.extractall(os.getcwd()+"/Vigo/")
            else:
                with zipfile.ZipFile(os.getcwd()+"/Arousa/"+zip, 'r') as zip_ref:
                    zip_ref.extractall(os.getcwd()+"/Arousa/")

def menu():
    global sentinelapi
    global Vigo
    print("0-Select which Ria do you want to get data from (default Vigo)")
    print("1-Process in-situ data")
    print("2-Download satellite data")
    print("3-Process satellite images")
    print("4-Generate Spreadsheet")
    print("5-Exit")
    seleccion = input("Option: ")
    print()
    if seleccion == "0":
        print("0-Vigo")
        print("1-Arousa")
        print("Anything else-exit")
        seleccionRia = input("Option: ")
        if seleccionRia == "0":
            Vigo = True
            print("Changed to Ria de Vigo")
            menu()
        elif seleccionRia == "1":
            Vigo = False
            print("Changed to Ria de Arousa")
            menu()
        else:
            menu()
    if seleccion == "1":
        BuscarCoord()
        BuscarDatos()
        menu()
    elif seleccion == "2":
        print("0-Sentinel")
        print("1-Landsat")
        print("2-Both")
        print("Anything else-exit")
        seleccionDownload = input("Option: ")
        if seleccionDownload == "0":
            username=input("Type Sentinel API username: ")
            password=input("Type Sentinel API password: ")
            sentinelapi = SentinelAPI(username, password, 'https://scihub.copernicus.eu/dhus')
            SentinelFetch()
            menu()
        elif seleccionDownload == "1":
            username=input("Type Landsat API username: ")
            password=input("Type Landsat API password: ")
            LandsatFetch(username,password)
            LandsatExtract()
            menu()
        elif seleccionDownload == "2":
            username=input("Type Sentinel API username: ")
            password=input("Type Sentinel API password: ")
            sentinelapi = SentinelAPI(username, password, 'https://scihub.copernicus.eu/dhus')
            SentinelFetch()
            username=input("Type Landsat API username: ")
            password=input("Type Landsat API password: ")
            LandsatFetch(username,password)
            LandsatExtract()
            menu()
    elif seleccion == "3":
        AcoliteProcess()
        menu()
    elif seleccion == "4":
        print("0-Neither show nor save")
        print("1-Save images")
        print("2-Show images")
        print("3-Both")
        print("Anything else-exit")
        seleccionProcess = input("Option: ")
        if seleccionProcess == "0":
            snapc2(False,False)
            menu()
        elif seleccionProcess == "1":
            snapc2(False,True)
            menu()
        elif seleccionProcess == "2":
            snapc2(True,False)
            menu()
        elif seleccionProcess == "3":
            snapc2(True,True)
            menu()
        else:
            menu()
    elif seleccion == "5":
        sys.exit()
    else:
        print("Option not found")
        menu()


if __name__ == '__main__':
    menu()



