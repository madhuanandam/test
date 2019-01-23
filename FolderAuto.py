import os, glob
import xml.dom.minidom

from numpy import save
from pandas import DataFrame
from openpyxl import load_workbook


def searchFile(serviceName, file_name):

    list = os.chdir("C:\\Temp\\SOA\\" + serviceName + "\\trunk\\SOA\\" + serviceName + "")
    list_config, list_datasource, list_wsdl, list_service = [], [], [], []

    for file in glob.glob(file_name):

        #print(file)
        doc = xml.dom.minidom.parse(file)
        for property_wsdl in doc.getElementsByTagName('binding.ws'):

            if property_wsdl:
                conf_file_wsdl = property_wsdl.getAttribute('location')
                if conf_file_wsdl:
                    list_wsdl.append(conf_file_wsdl)

                else:
                    conf_file_wsdl='NA'
                    list_wsdl.append(conf_file_wsdl)

            else:
                conf_file_wsdl='NA'
                list_wsdl.append(conf_file_wsdl)


        for property in doc.getElementsByTagName('binding.jca'):

            if property:
                conf_file = property.getAttribute('config')
                # print('-------------Start-------------')
                # print('Values for the service name ', serviceName, 'are as below ')
                # print('------')
                # print('Configuration JCA is ',conf_file)
                list_config.append(conf_file)
                list_service.append(serviceName)

            else:
                conf_file='NA'
                list_config.append(conf_file)
                list_service.append(serviceName)

            doc2 = xml.dom.minidom.parse(conf_file)

            for property2 in doc2.getElementsByTagName('connection-factory'):

                if property2:
                    conf_file_loc = property2.getAttribute('location')
                    # print('Respective data source is ',conf_file_loc)
                    list_datasource.append(conf_file_loc)
                    # list_service.append(serviceName)
                    # print('------------End--------------')

                else:
                    conf_file_loc='NA'
                    list_datasource.append(conf_file_loc)

    val_config=len(list_config)
    val_datasource=len(list_datasource)
    val_wsdl=len(list_wsdl)
    val_service=len(list_service)

    #print(list_config, list_datasource, list_wsdl, list_service)
    #print(val_config, val_datasource , val_wsdl , val_service)
    max_value=max(val_config, val_datasource , val_wsdl , val_service)
    #print(max_value)


    if len(list_config)==max_value:
        req1=max_value-len(list_datasource)
        for i in range(0,req1):
            list_datasource.append('NA')
        req2 = max_value - len(list_wsdl)
        for i in range(0, req2):
            list_wsdl.append('NA')
        req3 = max_value - len(list_service)
        for i in range(0, req3):
            list_service.append(serviceName)

    elif len(list_datasource)==max_value:
        req1 = max_value - len(list_config)
        for i in range(0, req1):
            list_config.append('NA')
        req2 = max_value - len(list_wsdl)
        for i in range(0, req2):
            list_wsdl.append('NA')
        req3 = max_value - len(list_service)
        for i in range(0, req3):
            list_service.append(serviceName)

    elif len(list_wsdl)==max_value:
        req1 = max_value - len(list_config)
        for i in range(0, req1):
            list_config.append('NA')
        req2 = max_value - len(list_datasource)
        for i in range(0, req2):
            list_datasource.append('NA')
        req3 = max_value - len(list_service)
        for i in range(0, req3):
            list_service.append(serviceName)

    elif len(list_service)==max_value:
        req1 = max_value - len(list_config)
        for i in range(0, req1):
            list_config.append('NA')
        req2 = max_value - len(list_wsdl)
        for i in range(0, req1):
            list_wsdl.append('NA')
        req3 = max_value - len(list_datasource)
        for i in range(0, req1):
            list_datasource.append('NA')

    #print(list_config, list_datasource, list_wsdl, list_service)




    os.chdir("C:\\Temp\\SOA\\")

    return list_config, list_datasource, list_wsdl, list_service

path=os.chdir("C:\\Temp\\SOA\\")
list=os.listdir(path)
#print(list)
newlist=[]
for dir_list in list:
    #print(dir_list)
    try:
        newdir = os.chdir("C:\\Temp\\SOA\\" + dir_list + "\\trunk\\SOA\\" + dir_list + "")
        newlist.append(dir_list)
    except:
        print('No SOA folder for '+dir_list)


#print(newlist)
file_name='composite.xml'
final_list_config, final_list_datasource, final_list_wsdl, final_list_service, f_final_service, f_final_config, f_final_datasource, f_final_wsdl=[],[],[],[],[],[],[],[]

for dir_list in newlist:
    list_config, list_datasource, list_wsdl, list_service = searchFile(dir_list, file_name)


    final_list_config.append(list_config)
    final_list_datasource.append(list_datasource)
    final_list_wsdl.append(list_wsdl)
    final_list_service.append(list_service)

#print(final_list_config)
#print(final_list_datasource)
#print(final_list_wsdl)
#print(final_list_service)

for sep_list in final_list_service:
    for inn_list in sep_list:
        f_final_service.append(inn_list)

for sep_list in final_list_wsdl:
    for inn_list in sep_list:
        f_final_wsdl.append(inn_list)

for sep_list in final_list_config:
    for inn_list in sep_list:
        f_final_config.append(inn_list)

for sep_list in final_list_datasource:
    for inn_list in sep_list:
        f_final_datasource.append(inn_list)

os.chdir("C:\\Temp\\excel\\")
dataframe = DataFrame({'Service Name':f_final_service, 'Dependent Service':f_final_wsdl,'JCA':f_final_config, 'DataSource':f_final_datasource})
#dataframe = DataFrame([final_list_service])
#print(dataframe)
#print(len(final_list_service))
#print(final_list_service[1][1])
dataframe.to_excel('test.xlsx', sheet_name='sheet1', index=False)














