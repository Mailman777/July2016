import openpyxl

#Should run through a list of programs and retrive a series of information adding each
# piece of information to a new excel sheet. Counts Spotwelds and Patterns of welds

testfile = 'listofprograms.txt'
with open(testfile,"r") as f: #opens a txt file with all the filenames to cycle through
	for line in f: # runs the whole program for each filename
		line = line.strip('\n') #had to strip the \N from the imported filename
		line2 = line
		f = open(line2)
		
		if len(line) >31: #Excel sheets limited to 31 Chars
			line = line[0:30]
		call_list = {}
		looper = 0
		loop = 2
		file_name = line
		print(file_name)
		print(looper)
		#f = open(file_name)	
		text = f.readlines()
		wb = openpyxl.load_workbook('programs.xlsx')
		#sheet = ws.active()

		ws2 = wb.create_sheet(title=file_name)
		wb.save('programs.xlsx')

		#sheet = wb.get_sheet_by_name('Sheet1')

		counter = 1
		print( counter, 'Counter')
		
		#text2 = text[looper-4]
		ws2['C1'] = file_name
		
		def check_cnt(N): #function designed to look for patterns in code
						# each robot code is made different and the identifier
						# can be in any line preceding the CALL line.
			for i in range(N-5,N):
				if 'CNT]=' in text[i]: 
					new = text[i].split('=')
					new[1] = new[1].strip('    ;\n')
					print( text[i])
					return int(new[1])
				else:
					return 1
			
			
		def check(X): #Finds lines with CALL function and adds them to dict
					
			if 'CALL' in X:
				global loop
				global counter
				broken = X.split(" ")
				lefty = X[X.find('CALL ',)+5:]
				ws2['C{}'.format(loop)] = lefty
				loop += 1
				#text2 = text[looper-counter].find('CNT]=')
				#text3 = text[looper-counter2].find('=')
				#print(text2)
				for i in broken: # Adds 
					if 'F_' in i and i not in call_list:
						call_list[i] = check_cnt(looper)
						
					elif 'F_' in i and i in call_list: #if already in list Adds to count
						call_list[i] += check_cnt(looper)
						
					else:
						continue
				return True
			else:
				return False
			


		with open(line2) as f: # calling each line in the program to run check()
			for line in f:
				looper += 1
				if check(line) == True:
					continue
					#print(text[looper-4][text[looper-4].find('=') + 1:text[looper-4].find('=')+3] )
				else:
					continue
				
		print(call_list)
		Z = 1
		for i in call_list: #prints to excel sheet the entire dict
			ws2['I{}'.format(Z)] = i
			ws2['J{}'.format(Z)] = call_list[i]
			Z+=1
		f = open(line2)	
				
		with open(line2) as text_file:
			contents = text_file.read()


		print(contents.count('SPOT[')) #counts single spots in program
		ws2['K1'] = contents.count('SPOT[')		
		for i in range(0,35): #counts the number of each type of schedule for the spots
								#Sum of all schedules should = total # of spots
			#print( "Schedule   {} = {}".format(i,contents.count('S={},'.format(i)) ))
			ws2['G{}'.format(i+1)] = "Schedule   {} =".format(i)
			ws2['H{}'.format(i+1)] = contents.count('S={},'.format(i))
		wb.save('programs.xlsx')	
		#line[line.find('CALL ') + 5:]