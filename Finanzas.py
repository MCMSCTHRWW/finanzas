import openpyxl, datetime, requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

class Coin:#Crear la clase Coin
	def __init__(self, nombre):
		self.nombre = nombre
		self.url = 'https://www.binance.com/es/price/'+self.nombre
		self.page = requests.get(self.url)
		self.soup = BeautifulSoup(self.page.content, 'html.parser')
		self.precio = self.soup.find('div', class_='css-12ujz79')
		self.session = requests.Session()
		
#Crear valor del euro
euro = Coin('euro')
valor_euro = euro.precio.text.replace('$ ','').replace(',','')#Quitar el inicio del valor

#Cargar el documento Excel de Finanzas
workbook = load_workbook('finanzas.xlsx')
#Seleccionar la hoja a modificar del excel
movimientos = workbook['Movimientos']
cartera = workbook['Cartera']
historicoprecios = workbook['HistoricoPrecios']
	
accion_realizada = input('Deseas COMPRAR, VENDER, o ACTUALIZAR los datos?: ')
if accion_realizada == 'COMPRAR':
	accion = 'comprado'
if accion_realizada == 'VENDER':
	accion = 'vendido'

##Volver a recopilar las monedas que ya hay con su cantidad quitando el FIAT
def recopilar():
	
	global longitud, listado_ordenado
	listado_ordenado = []
	j = 3
	while cartera['A'+str(j)].value is not None:
		datos_moneda = [cartera['A'+str(j)].value, None]
		listado_ordenado.append(datos_moneda)
		j += 1
	longitud = len(listado_ordenado)
	return(listado_ordenado)

def obtener_precios(nombre):
	global precio_actual
	coin = Coin(nombre)
	coin.precio = coin.precio.text.replace('$ ','').replace(',','')
	coin.precio = float(coin.precio) / float(valor_euro)
	precio_actual = coin.precio
	return coin.precio

#### SI queremos actualizar los datos de nuestra cartera(ACTUALIZAR)
def actualizar():
	nombre_monedas = []
	recopilar()
	#Crear nuevo listado
	actualizado = []
	vacio = False
	for k in range(longitud):
		nombre = cartera['A'+str(k+3)].value
		nombre_monedas.append(nombre)
		datos_actualizados = [cartera['A'+str(k+3)].value, cartera['B'+str(k+3)].value, obtener_precios(listado_ordenado[k][0]), cartera['B'+str(k+3)].value*obtener_precios(listado_ordenado[k][0])]
		actualizado.append(datos_actualizados)
		#Checkeamos si hay alguna moneda que hemos vendido su totalidad
		if cartera['B'+str(k+3)].value == 0:
			vacio = True
	#Ordenar la cartera por valor en orden descendiente	
	actualizado.sort(key = lambda x: x[3], reverse = True)	
	##Imprimir la lista de nuevo en las columnas de la hoja Cartera
	for q in range(longitud):
			cartera['A'+str(q+3)] = actualizado[q][0]
			cartera['B'+str(q+3)] = actualizado[q][1]
			cartera['C'+str(q+3)] = actualizado[q][2]
	if vacio == True:
		cartera['A'+str(longitud+2)] = None
		cartera['B'+str(longitud+2)] = None
		cartera['C'+str(longitud+2)] = None
		del actualizado[-1]
		workbook.save('finanzas.xlsx')
	###Histórico precios		
	##Imprimir los nuevos datos en el histórico
	pos = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P',
	'Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI'
	'AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB'
	'BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU']
	t = 1
	while historicoprecios['A'+str(t)].value is not None:
		t += 1
	historicoprecios['A'+str(t)].value = datetime.datetime.now()#Poner la fecha de hoy
	historicoprecios['A'+str(t)].number_format = 'dd-mm-yy hh:mm'#Poner el formato de fechas en (DIA-MES-AÑO)
	historicoprecios['C'+str(t)] = float(cartera['B2'].value)#Actualizar el valor FIAT
	for q in range(len(actualizado)):
		for moneda in nombre_monedas:
			r = 0
			while moneda != historicoprecios[pos[r]+'1'].value:
				r += 1
			if moneda == actualizado[q][0]:
				historicoprecios[pos[r]+str(t)] = cartera['B'+str(q+3)].value*obtener_precios(actualizado[q][0])
	
####Si hemos echo algún movimiento(COMPRAR)(VENDER)
def movimiento(accion_realizada):
	recopilar()
	###Fechas [A]
	i=1
	while movimientos['A'+str(i)].value is not None:#Comprueba si la celda esta vacía. Si está ocupada, pasa a la siguiente
		i += 1
	movimientos['A'+str(i)].value = datetime.datetime.now()#Poner la fecha de hoy
	movimientos['A'+str(i)].number_format = 'dd-mm-yy hh:mm'#Poner el formato de fechas en (DIA-MES-AÑO)
	
	###Moneda [B]
	moneda = input('Que moneda has '+accion+'?: ')#Escribir la moneda que has comprado
	movimientos['B'+str(i)] = moneda
	
	###Cantidad en euros [E]
	cantidad = input('Cuántos euros has '+accion+' de '+moneda+'?: ')
	cantidad = float(cantidad.replace(',','.'))
	if accion_realizada == 'VENDER':
		cantidad = cantidad * (-1)
	movimientos['E'+str(i)] = cantidad
		
	###Cantidad de moneda [C]
	cantidad_moneda = input('Cuánta cantidad de '+moneda+' has '+accion+'?: ')
	cantidad_moneda = float(cantidad_moneda.replace(',','.'))
	if accion_realizada == 'VENDER':
		cantidad_moneda = cantidad_moneda *(-1)
	movimientos['C'+str(i)] = float(cantidad_moneda)
	
	###Precio compra/venta [D]
	precio = float(cantidad) / float(cantidad_moneda)
	print('Has '+accion+' un total de '+str(abs(cantidad_moneda))+' '+moneda+' a '+str(precio)+' €/'+moneda+' por valor de '+str(abs(float(cantidad)))+' euros')
	movimientos['D'+str(i)] = float(precio)
	
	###Recopilar las monedas que ya hay
	listado_monedas = []
	j = 3
	while cartera['A'+str(j)].value is not None:
		datos_moneda = (cartera['A'+str(j)].value)
		listado_monedas.append(datos_moneda)
		j += 1
	
	###Dinero en cartera [F]. Dos opciones, operar con la cantidad anterior de la moneda, o sumar nueva moneda		
	##Añadir la nueva moneda a listado_monedas si es nueva
	if moneda not in listado_monedas:
		j = 2
		while cartera['A'+str(j)].value is not None:
			j += 1
		cartera['A'+str(j)] = moneda
		cartera['B'+str(j)] = float(cantidad_moneda)
		cartera['B2'] = float(cartera['B2'].value) - float(cantidad)#Actualizar valor de FIAT
		###Añadir al histórico
		pos = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P',
		'Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI'
		'AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ','BA','BB'
		'BC','BD','BE','BF','BG','BH','BI','BJ','BK','BL','BM','BN','BO','BP','BQ','BR','BS','BT','BU']
		p = 0
		while historicoprecios[pos[p]+'1'].value is not None:
			p += 1
		historicoprecios[pos[p]+'1'] = moneda
			
	##Si es una moneda de la lista: O compras o vendes
	if moneda in listado_monedas:
		if moneda == 'FIAT':#Si has sacado o metido EUROS a la cartera
			cartera['B2'] = float(cartera['B2'].value) + float(cantidad_moneda)
		else:		
			j = 2
			while moneda != cartera['A'+str(j)].value:
				j += 1
			cartera['B'+str(j)] = float(cartera['B'+str(j)].value) + float(cantidad_moneda)#Actualizar valor de la moneda
			cartera['B2'] = float(cartera['B2'].value) - float(cantidad)#Actualizar valor de FIAT
	workbook.save('finanzas.xlsx')
	actualizar()

####Preguntar si has comprado o vendido o solo actualizar los datos [COMPRAR,VENDER, ACTUALIZAR]
def pregunta(accion_realizada):
	if (accion_realizada == 'COMPRAR') or (accion_realizada == 'VENDER'):
		movimiento(accion_realizada)
	if accion_realizada == 'ACTUALIZAR':
		actualizar()
pregunta(accion_realizada)

#Guardar los cambios en el documento Excel
workbook.save('finanzas.xlsx')
print('Cambios realizados satisfactoriamente en el documento Excel')
