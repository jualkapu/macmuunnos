from xlrd import open_workbook
import sys
import csv
import re
import string

# Otetaan argumenttina saatu muokkaamaton .xlsx file MAC osoitteista
inputFile = sys.argv[1]

# Avataan syötetty exeli ja valitaan eka taulukko
book = open_workbook(inputFile)
sheet = book.sheet_by_index(0)

# Lukee arvon joka riviltä ja lisää listaan 
# Range = 0 - taulukon rivien määrä
# Tää palauttaa jokasen solun sisällön listana. Values = lista listoja. Miksi?
values = [sheet.row_values(i) for i in range(0, sheet.nrows)]

# Regex MAC-osoitteelle
pattern = re.compile("^([0-9A-Fa-f]{2}[:]){5}[0-9A-Fa-f]{2}$")

# Löytyykö listasta virheellisiä MAC-osoitteita
invalids = False

# Jos löytyy lisätään tänne
virheelliset = []

# Etsitään virheelliset MAC-osoitteet ja poistetaan ne
# Tunkataan virheelliset omaan .csv filuun
for x in range(len(values) - 1,  0, -1):
	
	st = values[x][0]
	values[x][0] = re.sub(r"\s+", "", st) # Poistetaan trailing whitespacet
	
	p = pattern.match(values[x][0]) # Verrataan taulukon MAC:iä Regexiin
	
	if not p:
		virheelliset.append(values[x])
		values.remove(values[x])
		invalids = True

#muutetaan sisältö oikeaan muotoon
for x in range(0, len(values)):
	str1 = ''.join(values[x]).replace(':',' ') # Poistetaan kaksoispisteet
	str1 = str1 + ',"OU=Inet-Media,OU=Devices,DC=pos2,DC=veikkaus,DC=fi"' # Lisataan tarvittava loppu
	values[x] = re.sub(r"\s+", "", str1)

# Muutetaan eka rivi syottokelpoiseen muotoon
values[0] = 'username,uo'
	
# loopataan listan läpi ja kirjoitetaan muutetut tiedot uuteen filuun
with open('output.csv','w') as file: 
	for x in range(0, len(values)):
		file.write(values[x])
		file.write('\n')

# Jos virheellisä on löytynyt kirjoitetaan ne omaan filuunsa
if invalids:
	with open('virheelliset.csv','w') as file: 
		for x in range(0, len(virheelliset)):
			file.write(virheelliset[x][0])
			file.write('\n')
	print("Virheellisia osoitteita loytyi " + str(len(virheelliset)) + " kappaletta")
	print("Virheelliset MAC:it loytyvat tiedostosta: Virheelliset.csv")
	
print("Oikeassa muodossa olevat MAC:it loytyvat tiedostosta: output.csv")