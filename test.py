import os
directory = 'C:\Py\MD_LS'

f = open('listofprograms.txt','w')

for filename in os.listdir(directory):
	f.write('{}\n'.format(filename))