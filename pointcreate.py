import datetime
import openpyxl

wb = openpyxl.load_workbook('FencePanelDesignTable.xlsx', data_only = True)
sheet = wb.get_sheet_by_name('Sheet1')


railson = 4
rail1 = sheet['B3'].value
rail2 = sheet['C3'].value
rail3 = sheet['D3'].value
rail4 = sheet['E3'].value
length = sheet['F22'].value
num_pickets = sheet['V3'].value
picket_width = sheet['O3'].value
picket_depth =sheet['P3'].value
weld_size = 6.350
first_space = sheet['N3'].value
post_size = sheet['G3'].value
spacing = sheet['C22'].value
beginningX = post_size + first_space - 1
beginningY = rail1
program_line = 1

today = datetime.date.today()
print(today)
time = datetime.datetime.now().time()
print(time)
#creates filename same as CAD name
prog_title = sheet['A1'].value.split(':') 
filename = open('{0}.ls'.format(prog_title[1]),'w')



def point_create(X,Y,Z,track,N): #creates points accoding to XYZ as well as track
	#also creates the point number as N
	return """P[{0}] {{
   GP1:
	UF: 0, UT : 1,	CONFIG : 'N U T, 0, 0, 0',
	X = {1:.3f} mm, 	Y = {2:.3f} mm, 	Z = {3:.3f} mm,
	W = 0.000 deg, 	P = 45.000 deg,	R = 90.000 deg,
	E1 = {4:.3f} mm
}};
""".format(N, X, Y, Z, track)

def program_start_info(): #first few lines of the robot code
	intro = """/MN
   1:  !WeldPRO Auto-Generated TPP ;
   2:  !Part outer_mast-1531658, Feature ;
   3:   ;
   4:  UFRAME_NUM=0 ;
   5:  UTOOL_NUM=1 ;
   6:  !Feature Approach ;
   7:J P[1] 100% FINE    ;\n"""
	return intro

	
def crap(): #intro stuff in program, lot to be changed still
	return"""/PROG  {}
/ATTR
OWNER		= MNEDITOR;
COMMENT		= "WeldPRO Auto-Gen";
PROG_SIZE	= 943;
CREATE		= DATE 16-06-28  TIME 10:42:34;
MODIFIED	= DATE 16-06-28  TIME 10:42:34;
FILE_NAME	= ;
VERSION		= 0;
LINE_COUNT	= 13;
MEMORY_SIZE	= 1263;
PROTECT		= READ_WRITE;
TCD:  STACK_SIZE	= 0,
      TASK_PRIORITY	= 50,
      TIME_SLICE	= 0,
      BUSY_LAMP_OFF	= 0,
      ABORT_REQUEST	= 0,
      PAUSE_REQUEST	= 0;
DEFAULT_GROUP	= 1,*,*,*,*;
CONTROL_CODE	= 00000000 00000000;
/APPL
  ARC : TRUE ; 
  ARC Welding Equipment : 1,*,*,*,*;\n""".format(prog_title[1])
	
def code_sections(N): #some of the deliminators in the program
	if N == 0:
		return "/APPL"
	elif N == 1:
		return "/MN"
	elif N == 2:
		return "/POS\n"
	elif N == 3:
		return "/END\n"
	else:
		return "error"


	
def begin_robot_line(N): #This is how each line in the program list starts
	build_robot_num = (" " * 3 + "{0}:".format(N))
	return build_robot_num

def line_writer(N,weldnum,speed): #This writes the start of a weld
	return """{0}L P[{1}] {2} FINE  
    : Weld Start E1[1,1]   ;\r""".format(begin_robot_line(N),weldnum,speed)
	
def line_writer2(N,weldnum,speed): #Writes the end of a weld
	return """{0}L P[{1}] WELD_SPEED FINE 
	: Weld End E1[1,1]   ;\r""".format(begin_robot_line(N+1),weldnum +1)
	
def program_write_weld(N,step): #labels the weld segment in the weld program
	build_seg_line = (begin_robot_line(N) 
		+ '!Segment{0}'.format(step))
	return build_seg_line
		
		
def robo_prog_name(L,H,picket_number,series,LP,style): #creates a weld program
	#based off of information dependant on entry of program parameters
	name = "{0}-{1}-{2}-{3}-{4}-{5}".format(L,H,picket_number,series,LP,style)
	return name
		

	
def create_weld_horz(X,Y,Z,track,N,picket_w): #Creates a weld start and end point
	pstart = point_create(X,Y,Z,track,N)
	pend = point_create(X+picket_w,Y,Z,track,N+1)
	return pstart + pend

	
	
#writing lines beyon this
filename.write(crap() + program_start_info())	
ind = 1

#This writes the weld sequences step by step	
for i in range(1, num_pickets *2 * railson,2):
	speed = '2000mm/sec'
	line_number = i + 7
	line1 = line_writer(line_number,i,speed)
	line2 = line_writer2(line_number,i,'Weld_Speed')
	filename.write(line1 + line2)
	
filename.write(code_sections(2))


#this writes the weld points and numbers them for each rail
while ind < railson * num_pickets *2:

	filename.write(create_weld_horz(beginningX,rail1,100,200,ind,picket_width))
	ind += 2
	filename.write(create_weld_horz(beginningX,rail2,100,200,ind,picket_width))
	ind += 2
	filename.write(create_weld_horz(beginningX,rail3,100,200,ind,picket_width))
	ind += 2
	filename.write(create_weld_horz(beginningX,rail4,100,200,ind,picket_width))
	ind += 2
	beginningX += spacing
	

	
filename.write(code_sections(3))
	
