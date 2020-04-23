from __future__ import print_function
from os.path import join, dirname, abspath
import xlrd

# Set variables
vdom_status = '0'
while vdom_status != 'N' and vdom_status != 'S':
    print ('Su configuracion emplea VDOMs (S/N):')
    vdom_status = raw_input()
    if vdom_status != 'N' and vdom_status != 'S':
    	print ('Por favor introduzca una opcion valida (S/N)')
print ('Introduce el nombre del VDOM a configurar:')
vdom_name = raw_input ()
print ('Introduce el nombre del Portal SSL:')
heading = raw_input()
print ('Introduce el nombre del Servidor LDAP:')
ldap = raw_input()

# Open the workbook
xl_workbook = xlrd.open_workbook('/Users/tonimartinez/Desktop/Scripts/ssl.xlsx')
sheet_names = xl_workbook.sheet_names()
xl_sheet = xl_workbook.sheet_by_index(0)

row = xl_sheet.row(0)  # 1st row
num_rows = xl_sheet.nrows

from xlrd.sheet import ctype_text   
file = open("ssl_vpn_auto.txt", "a")
# Si se trabaja con VDOMs esta linea es necesaria, si no no debe usarse
if vdom_status == 'S':
	file.write("config vdom"+"\n"+"edit "+vdom_name+"\n")
	
#Creacion usuarios
file.write("config user local"+"\n")
for row_idx in range(1, xl_sheet.nrows):    # Iterate through rows
	user = xl_sheet.cell(row_idx,0)
	file.write("edit " + user.value + "\n")
	file.write("set type ldap"+"\n")
	file.write("set ldap-server " + ldap +"\n")
	file.write("next"+"\n")
file.write("end"+"\n")

#Creacion portales SSL

for row_idx in range(1, xl_sheet.nrows):    # Iterate through rows
	user = xl_sheet.cell(row_idx,0)
	IP = xl_sheet.cell(row_idx,1)
	file.write("config vpn ssl web portal"+"\n")
	file.write("edit " + user.value +"\n")
	file.write("set tunnel-mode disable"+"\n")
	file.write("set web-mode enable"+"\n")
	file.write("config bookmark-group"+"\n")
	file.write("edit " + user.value + "_bookmarks"+"\n")
	file.write("config bookmarks"+"\n")
	file.write("edit " + user.value + "_RDP"+"\n")
	file.write("set apptype rdp"+"\n")
	file.write("set host " + IP.value +"\n")
	file.write("set description " + IP.value + "_RDP"+"\n")
	file.write("set server-layout es-es-qwerty"+"\n")
	file.write("set security any"+"\n")
	file.write("set preconnection-id 1"+"\n")
	file.write("set port 3389"+"\n")
	file.write("set sso auto"+"\n")
	file.write("set sso-credential sslvpn-login"+"\n")
	file.write("set sso-credential-sent-once enable"+"\n")
	file.write("next"+"\n")
	file.write("end"+"\n")
	file.write("next"+"\n")
	file.write("end"+"\n")
	file.write("set display-connection-tools disable"+"\n")
	file.write("set display-history disable"+"\n")
	file.write("set display-status disable"+"\n")
	file.write("set heading "+heading+"\n")
	file.write("set theme blue"+"\n")
	file.write("end"+"\n")
print ("Se han creado ", xl_sheet.nrows, "portales SSL y bookmarks personalizados con exito")

#Asociacion usuario - portal

file.write("config vpn ssl settings"+"\n")
file.write("config authentication-rule"+"\n")
for row_idx in range(1, xl_sheet.nrows):    # Iterate through rows
	user = xl_sheet.cell (row_idx,0)
	IP = xl_sheet.cell (row_idx,1)
	ID = str(row_idx+999)
	#ID = str(ID_temp.value)[:-2]
	file.write("edit " + ID +"\n")
	file.write("set users " + user.value +"\n")
	file.write("set portal "+ user.value + "\n")
	file.write("next"+"\n")
file.write("end"+"\n")
print ("Se han asociado ", xl_sheet.nrows, "portales SSL y bookmarks personalizados con exito")
file.close