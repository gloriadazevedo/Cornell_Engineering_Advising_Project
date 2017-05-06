#!/usr/bin/env python3

#All majors are: AE,BE,BME,CHEME,CE,CS,ECE,EP,EnvirE,ISST,OR,MSE,ME,Pre-health,SES,Undecided

#Script that processes student, advisor, and course data file into an AMPL .dat file

#Now student data has points instead of just 1st/2nd/3rd
#Each major has its own column that can't be hardcoded--must be scraped from the CSV

#The files are now hardcoded since this will be called from VBA and we know exactly 
#what those files are called

#Use PointAllocation.run which calls PointAllocation.mod and the respective .dat file 

#Input uses Kim's version of the course input instead of the Google Form

#Adds similar majors to all advisors based on their department major
#In this iteration, we use "directional" similar majors
#I.e. CS advisors can take OR, ISST, CS, and ECE
#However, ECE advisors can take CS but not OR, or ISST
#Includes breaking minorities into two different tiers of minority
#Tier 1: Hispanic only
#Tier 2: All other URM

printing="on"
#option_times should be "preference" or "conflict"
#if it's "conflict" then we want to return the complement set of all possible times
#if it's "preference" then we want to match the advisors with the times that they indicated
option_times = "preference"

import sys
import argparse
import csv
import copy
import os
import tkinter as tk
from tkinter import filedialog,constants,messagebox

#Only for choice model 
excluded_majors=['Undecided']

#Hard code the .dat file name so that the run file doesn't need to be changed
full_data_file='Output_dat_file.dat'

#Have a new file for the error printing
error_print_file = "ErrorPrint.txt"

#Other hardcoded files
student_file_points="New_Full_Student_Data.csv"
advisor_file="Advisor_Preference_Data.csv"
course_conflict_list="Course_Conflict_Data_Sheet.csv"
ampl_file="PointAllocationRun.run"

#Define column indices
#advisor_csv
advisor_id_col=0
advisor_dep_pairing_col=1
advisor_pairing_str_col=2
advisor_dept_col=3
advisor_majors_col=4
advisor_times_col=11
num_advisor_cols=12

#course_conflict_csv
course_code_col = 0
course_name_col = 1
course_lec_col = 2
course_days_col = 3
course_start_col = 4
course_end_col = 5
course_majors_col = 6


class TkFileDialogExample(tk.Frame):
		
	def __init__(self,root):
		tk.Frame.__init__(self, root)
		self.pack()

		button_opt = {'fill': constants.BOTH, 'padx': 5, 'pady': 5}
		
		#Get file paths and filenames for the ones we need to write on
		#student data with points
		tk.Button(self, text='Open student data file', command=self.student_askopenfilename).pack(**button_opt)
		
		#advisor preference form
		tk.Button(self, text='Open advisor data file', command=self.advisor_askopenfilename).pack(**button_opt)
		
		#course times and their respective major conflicts
		tk.Button(self, text='Open time conflicts data file', command=self.course_askopenfilename).pack(**button_opt)
		
		#ampl.exe filepath
		#tk.Button(self,text='File path to ampl.exe',command=self.amplexe_askopenfilename).pack(**button_opt)
		
		#ampl run file filepath
		tk.Button(self,text='File path to ampl run file',command=self.amplfile_askopenfilename).pack(**button_opt)
		
		#Submit button 
		self.submit=tk.Button(self,text="Submit",command=self.ask_quit)
		self.submit.pack(**button_opt)
		
		self.file_opt = options = {}
		options['defaultextension'] = '.csv'
		options['filetypes'] = [('all files', '.*'), ('comma separated', '.csv')]
		options['initialdir'] = './'
		
	#function for user to select the files that they want to open
	def student_askopenfilename(self):
		# get filename; does work when printed out
		self.student_file_points = tk.filedialog.askopenfilename(**self.file_opt)
		
	def advisor_askopenfilename(self):
		# get filename; does work when printed out
		self.advisor_file = tk.filedialog.askopenfilename(**self.file_opt)
		
	def course_askopenfilename(self):
		# get filename; does work when printed out
		self.course_conflict_list = tk.filedialog.askopenfilename(**self.file_opt)	
		
	# def amplexe_askopenfilename(self):
		# # get filename; does work when printed out
		# self.ampl_exe = tk.filedialog.askopenfilename(**self.file_opt)	
	
	def amplfile_askopenfilename(self):
		# get filename; does work when printed out
		self.ampl_file = tk.filedialog.askopenfilename(**self.file_opt)	
	
	#function for user to type in a filename that the pairing will save into
	def asksaveasfilename(self):
	# get filename
		self.ampl_file_result = tk.filedialog.asksaveasfilename(**self.file_opt)

	def ask_dat_filename(self):
		self.ampl_file_save=tk.filedialog.asksaveasfilename(**self.file_opt)

	def ask_quit(self):
		if tk.messagebox.askokcancel("Submit", "The scheduling algorithm will run. Proceed?"):
			if self.student_file_points!=None and self.advisor_file!=None and self.course_conflict_list!=None:	
				#Hardcoded the dat filename so that the run file doesn't need to be changed.
				#self.ask_dat_filename()
				main(self.student_file_points,self.advisor_file,self.course_conflict_list,full_data_file,self.ampl_file)
				#putting root.destroy after the call erases all tk boxes
				root.destroy()
				
			elif self.student_file_points==None:
				tk.messagebox.showinfo("Error", "Error: File with Students points is not specified.")
			elif self.advisor_file==None:
				tk.messagebox.showinfo("Error", "Error: File with advisor preferences is not specified.")
			elif self.course_conflict_list==None:
				tk.messagebox.showinfo("Error", "Error: File with course conflicts is not specified.")
			#elif self.ampl_exe==None:
			#	tk.messagebox.showinfo("Error", "Error: ampl.exe file is not specified.")
			elif self.ampl_file==None:
				tk.messagebox.showinfo("Error", "Error: AMPL run file is not specified.")

#Need to define this so the messagebox works later
#leaving all the I/O code in here, just in case
root = tk.Tk()
#This command hides the main tk window but the messagebox still operates correctly
root.withdraw()



#Mini helper function to convert strings from the form into consistent codes
def convert_string_to_code(long_major):
	major_code=""
	if long_major=="Aerospace Engineering" or long_major == "AE":
		major_code="AE"
	elif long_major=="Biological Engineering" or long_major == "BE":
		major_code="BE"
	elif long_major=="Biomedical Engineering" or long_major == "BME":
		major_code="BME"
	elif long_major=="Chemical Engineering" or long_major == "CHEME":
		major_code="CHEME"
	elif long_major=="Civil Engineering" or long_major == "CE":
		major_code="CE"
	elif long_major=="Computer Science" or long_major == "CS":
		major_code="CS"
	elif long_major=="Electrical and Computer Engineering" or long_major == "ECE":
		major_code="ECE"
	elif long_major=="Engineering Physics" or long_major == "EP":
		major_code="EP"
	elif long_major=="Environmental Engineering" or long_major =="EnvirE":
		major_code="EnvirE"
	elif long_major=="Information Science Systems and Technology" or long_major == "ISST":
		major_code="ISST"
	elif long_major=="Operations Research" or long_major=="OR":
		major_code="OR"
	elif long_major=="Materials Science Engineering" or long_major =="MSE":
		major_code="MSE"
	elif long_major=="Mechanical Engineering" or long_major=="ME":
		major_code="ME"
	elif long_major=="Pre-Health (Includes Pre-Vet and Pre-Dental)" or long_major=="Pre-health":
		major_code="Pre-health"
	elif long_major=="Science of Earth Systems" or long_major=="SES":
		major_code="SES"
	elif long_major=="Undecided":
		major_code="Undecided"
	else:
		major_code="Other/New Major"
	
	return major_code
		
#Mini helper function to convert department strings into the same major codes
def convert_department_to_code(long_dept):
	dept_code=""
	if long_dept=="Applied and Engineering Physics" or long_dept=="Engineering Physics" or long_dept =="EP":
		dept_code="EP"
	elif long_dept=="Biological and Environmental Engineering" or long_dept=="BE":
		dept_code="BE"
	elif long_dept=="Biomedical Engineering" or long_dept=="BME":
		dept_code="BME"
	elif long_dept=="Chemical and Biomolecular Engineering" or long_dept =="CHEME":
		dept_code="CHEME"
	elif long_dept=="Civil Engineering" or long_dept=="CE":
		dept_code="CE"
	elif long_dept=="Computer Science" or long_dept=="CS":
		dept_code="CS"
	elif long_dept=="Earth and Atmospheric Sciences" or long_dept=="SES":
		dept_code="SES"
	elif long_dept=="Electrical and Computer Engineering" or long_dept=="ECE":
		dept_code="ECE"
	elif long_dept=="Environmental Engineering" or long_dept=="EnvirE":
		dept_code="EnvirE"
	elif long_dept=="Information Science Systems and Technology" or long_dept=="ISST":
		dept_code="ISST"
	elif long_dept=="Materials Science and Engineering" or long_dept=="MSE":
		dept_code="MSE"
	elif long_dept=="Mechanical and Aerospace Engineering" or long_dept=="ME":
		dept_code="ME"
	elif long_dept=="Operations Research and Information Engineering" or long_dept=="OR":
		dept_code="OR"
		
	return dept_code
		
def full_permutation_time():
	days=["M","T","W","R","F"]
	times=["905","1010","1115","1220","125","230","335"]
	full_list=[]
	for i in days:
		for j in times:
			full_string=i+j
			full_list.append(full_string)
	return full_list
		
def check_similar_majors(major):
	#Using a dictionary to have direction of which advisors take which majors
	#Incded the dictionary key inside as well	
	major_dict={'ME':['ME','CE','EnvirE','SES','EP','Undecided'],\
				'CE':['CE','ME','EnvirE','SES','EP','Undecided'],\
				'EnvirE':['ME','CE','EnvirE','Undecided'],\
				'OR':['OR','ISST','CS','Undecided'],\
				'ISST':['ISST','OR','CS','Undecided'],\
				'CS':['CS','OR','ISST','ECE','Undecided'],\
				'ECE':['ECE','CS','Undecided'],\
				'CHEME':['CHEME','MSE','Undecided'],\
				'BE':['BE','BME','Pre-health','Undecided'],\
				'BME':['BME','BE','Pre-health','Undecided'],\
				'Pre-health':['Pre-health','BE','BME','Undecided'],\
				'EP':['EP','ME','CE','EnvirE','SES','Undecided'],\
				'SES':['SES','ME','CE','EnvirE','EP','Undecided'],\
				'MSE':['MSE','CHEME','Undecided'],\
				'Undecided':['Undecided','ME','CE','EnvirE','OR','ISST','CS','ECE','CHEME','BE','BME','Pre-health','EP','SES','MSE']}

	major_list=major_dict[major]
		
	return major_list

#Need a function to convert the dates and times such as "Wed 9:05" or "Wednesday 9:05" to W905
def convert_time(x):
	#Get the date and the time separately
	x_split=x.split(" ")
	#print(x_split)
	result=""
	time=x_split[1].split(":")
	if x_split[0]=="Mon" or x_split[0]=="Monday" or x_split[0]=="M":
		result="M"+ str(time[0])+str(time[1])
	elif x_split[0]=="Tues" or x_split[0]=="Tuesday" or x_split[0]=="T":
		result="T"+str(time[0])+str(time[1])
	elif x_split[0]=="Wed" or x_split[0]=="Wednesday" or x_split[0]=="W":
		result="W"+str(time[0])+str(time[1])
	elif x_split[0]=="Thur" or x_split[0]=="Thursday" or x_split[0]=="R":
		result="R" +str(time[0])+str(time[1])
	elif x_split[0]=="Fri" or x_split[0]=="Friday" or x_split[0]=="F":
		result="F" +str(time[0])+str(time[1])
	else:
		print("Error: No valid date for ")
		print(x)
	
	return result

#Function that converts the number part of the time to 2400 time for actual comparisons
def convert_time_24hr(x):
	if [x][0] in ("M","T","W","R","F"):
		x=x[1:]
	if int(x)<=800:
		x=int(x)+1200
	return int(x)

#Function that takes in an error code and asks the user 
#whether or not they want to continue
def msg_ask_continue(error):
	#If error is 0, then nothing is wrong and we should proceed
	if error==0:
		return True
	
	#If error is 1 then there are some suboptimal solutions
	#ask the user if they would like to continue and get a suboptimal solution
	#Maybe it's not bad enough that they can do by hand
	#one example is that the advisors aren't paired correctly
	elif error==1:
		if tk.messagebox.showinfo("Suboptimal", "Pre-check has found some errors in input data.  See ErrorPrint.txt or the Dashboard sheet in Excel for instructions on how to correct them."):
			return True
		else:
			return False
	#These will throw infeasible issues in AMPL so alert the user
	#and don't run the algorithm
	# elif error==2:
		# tk.messagebox.showerror("Infeasible", "Pre-check has found some errors in input data that will yield no solution from the model.  See ErrorPrint.txt for details.  The run will now terminate.")
		# return False
	return
	
#Function that takes into a list of times and returns the complement of 
#that set from the total list of sets
#use if we want to ask advisors for what times they aren't available
def find_available_times(unavailable_times):
	#have the full list of times
	all_times=full_permutation_time()
	#Check for each time if it's in the unavailable_times then remove it
	for time in unavailable_times:
		all_times.remove(time)
	#Return the list of available times
	return all_times
	
	
def main(student_file_points,advisor_file,course_conflict_list,full_data_file,ampl_file):
	
	#initialize data matrices
	advisor_csv=[]
	course_conflict_csv=[]
		
	#Open and read the advisor preference file
	#See header pre functions for column indices
	with open(advisor_file, encoding='utf-8') as f:
		#Have to process by hand since lists are getting split up
		#f.readline() gives one giant string
		#read and skip the header
		header=f.readline() 
		
		#first line
		reader=f.readline()
		while reader!="":
			iterator=0
			temp_string=""
			check_quote=0
			temp_index=[]
			temp_list=[]
			while iterator<len(reader):
				if reader[iterator]=='"' and check_quote==0:
					check_quote=1
					#Don't append the quote
					iterator=iterator+1
				elif reader[iterator]=='"' and check_quote==1:
					check_quote=0
					#Don't append the quotes
					#the string is complete so append it to the index
					temp_list.append(temp_string)
					temp_string=""
					temp_index.append(copy.deepcopy(temp_list))
					temp_list=[]
					#Need to skip the comma after the end quote as well
					iterator=iterator+2
				elif reader[iterator]=="," and check_quote==0:
					#Don't append the comma
					#Add the string to the index
					temp_index.append(temp_string)
					temp_string=""
					iterator=iterator+1
				elif reader[iterator]=="," and check_quote==1:
					#Append the comma to the string
					#Don't clear temp_string yet
					temp_list.append(temp_string)
					temp_string=""
					#When there is a comma in the quotes, it DOES NOT include an extra space after them so skip just the comma in this version
					iterator=iterator+1
				#If we're at the end of the line
				elif reader[iterator]=='\n' and temp_string!='':
					temp_index.append(temp_string)
					iterator=iterator+1
				else: #Regular letter
					temp_string=temp_string+reader[iterator]
					iterator=iterator+1				
			advisor_csv.append(temp_index)
			#Read the next line
			reader=f.readline()
			
			
	#Open and import the course information from file
	#See the header before the functions for the column definitions
	with open(course_conflict_list, encoding='utf-8') as f:
		#Have to process by hand since lists are getting split up
		#f.readline() gives one giant string
		#read and skip the header
		reader=f.readline() 
		
		#first line
		reader=f.readline()
		while reader!="":
			iterator=0
			temp_string=""
			check_quote=0
			temp_index=[]
			temp_list=[]
			while iterator<len(reader):
				if reader[iterator]=='"' and check_quote==0:
					check_quote=1
					#Don't append the quote
					iterator=iterator+1
				elif reader[iterator]=='"' and check_quote==1:
					check_quote=0
					#Don't append the quotes
					#the string is complete so append it to the index
					temp_list.append(temp_string)
					temp_string=""
					temp_index.append(temp_list)
					temp_list=[]
					#Need to skip the comma after the end quote as well
					iterator=iterator+2
				elif reader[iterator]=="," and check_quote==0:
					#Don't append the comma
					#Add the string to the index
					temp_index.append(temp_string)
					temp_string=""
					iterator=iterator+1
				elif reader[iterator]=="," and check_quote==1:
					#Append the comma to the string
					#temp_string=emp_string+reader[iterator]
					#Don't clear temp_string yet
					temp_list.append(temp_string)
					temp_string=""
					#When there is a comma in the quotes, it DOES NOT include an extra space after them so skip that comma only in this 2nd version
					iterator=iterator+1
				else: #Regular letter
					temp_string=temp_string+reader[iterator]
					iterator=iterator+1
			#At this point, if there's only a string at the end, then we should append it
			if temp_string!="" and temp_string!="\n":
				temp_index.append(temp_string.split("\n")[0])
			course_conflict_csv.append(temp_index)
			#Read the next line
			reader=f.readline()
	
	#Manually need to write the error file
	with open(error_print_file, 'w') as output_file:
		#Write out heading
		output_file.write("Errors \n")
		
		#Write out the invalid pairings as comments in the dat file at the top
		#Add the advisor pairings
		#First need to check if the advisors had requested each other
		#Then need to check if their times overlap
		#Initialize a dictionary just for the pairings to each other
		pairing_dict=dict()
		
		#Initialize this error parameter as a check to see if the dat file should stop
		error=0
		
		for i in range(0,len(advisor_csv)):
			#Add all the advisors and their pairing if any
			pairing_dict[advisor_csv[i][advisor_id_col].lower()]=advisor_csv[i][advisor_pairing_str_col].lower()
		
		#Remove empty key in pairing_dict
		if '' in pairing_dict.keys():
			del pairing_dict['']
		
		#Go through the keys to see if the advisors are in each other's values
		for key in pairing_dict.keys():
			if not isinstance(pairing_dict[key],str):
				for value in pairing_dict[key]:
					if key not in pairing_dict[value.lower()]:
						output_file.write("Advisor "+key+" not in "+value.lower()+" pairing request list.\n")
						error=1
			elif pairing_dict[key]!='':
				if key != pairing_dict[pairing_dict[key].lower()].lower():
					output_file.write("Advisor "+key+" not in "+pairing_dict[key]+" pairing request list.\n")
					error=1
				
		#Initialize a dictionary for the pairing times
		pairing_time_dict=dict()
		for i in range(0,len(advisor_csv)):
			advisor = advisor_csv[i][advisor_id_col].lower()
			#Add all the advisors and their times
			#Check if they even put times; if they didn't then give them all the times
			if len(advisor_csv[i])>=num_advisor_cols:
				if option_times = "conflict":
					pairing_time_dict[advisor]=find_available_times(advisor_csv[i][advisor_times_col])
				elif option_times = "preference"
					pairing_time_dict[advisor]=advisor_csv[i][advisor_times_col]
			else:
				pairing_time_dict[advisor]=full_permutation_time()
				output_file.write("Advisor "+advisor+" did not write any time preferences and they will be assigned all possible times when computing the model.\n")
				error = 1
				
		#Delete the empty key
		if '' in pairing_time_dict.keys():
			del pairing_time_dict['']
			
		#Need to check that the valid pairings have overlapping times
		#Assume that the user is aware that the pairings don't match already
		for key in pairing_time_dict.keys():
			#check that they're in each other's pairings using pairing_dict
			#checks if there's multiple pairings
			if not isinstance(pairing_dict[key],str):
				for value in pairing_dict[key]:
					#doesn't matter if they're a valid pair
					#Check that their times intersect
					if len(set(pairing_time_dict[key]).intersection(pairing_time_dict[value]))==0:
						output_file.write("Error: Times for "+key+" and "+value+" do not intersect. "+value+" will be removed from the pairing request of "+key+" when computing the model.\n")
						pairing_dict[key].remove(value)
						error=1
			#There is only one pairing and need to make sure it's not empty
			#doesn't matter if they requested each other
			elif pairing_dict[key]!="":
				value = pairing_dict[key]
				if len(set(pairing_time_dict[key]).intersection(pairing_time_dict[pairing_dict[key]]))==0:
					output_file.write("Error: Times for "+key+" and "+value+" do not intersect. "+value+" will be removed from the pairing request of "+key+" when computing the model.\n")
					pairing_dict[key]=""
					error=1
						
		#Create a dictionary that has the time slots as the 
		#keys and the values are either 
		#empty if there are no conflicts in that time or the 
		#majors that CAN'T be in that time slot
		#Initialize dictionary
		schedule_dict=dict()
		#For each time slot, make it a key
		all_times=full_permutation_time()
		for i in range(0,len(all_times)):
			schedule_dict[all_times[i]]=[]
		
		#Will always have a blank key--need to delete
		if '' in schedule_dict:
			del schedule_dict['']
		
		#Add values to the keys from the conflicts
		for i in range(0,len(course_conflict_csv)):
			#Convert to only pull the times--split by space first
			temp_start_str=course_conflict_csv[i][course_start_col].split(" ")
			temp_end_str=course_conflict_csv[i][course_end_col].split(" ")
			
			#Now split by ":"
			#print(temp_start_str)
			start_time=convert_time_24hr(temp_start_str[0].split(":")[0]+temp_start_str[0].split(":")[1])
			end_time=convert_time_24hr(temp_end_str[0].split(":")[0]+temp_end_str[0].split(":")[1])
			days=course_conflict_csv[i][course_days_col]
			
			#Need to get a list of major codes that conflict with the time
			conflict_codes=course_conflict_csv[i][course_majors_col]
			
			#For each time in dictionary, add the majors to the list in the dictionary
			#if the dictionary key falls in between the start and end times 
			#(converted to an integer)
			for j in schedule_dict.keys():
				#The dictionary has the day ahead of it
				j_day_convert=j[0]
				j_time_convert=convert_time_24hr(j[1:])
				for jj in range(0,len(days)):
					#Check if days match and if dictionary key time is between start and end
					if days[jj]==j_day_convert and j_time_convert>=start_time and j_time_convert<= end_time:
						#Add all the majors in the list into the dictionary if they're not already there
						#First need to check if there are multiple in there; check if it's a string
						if not isinstance(conflict_codes,str):
							for jjj in conflict_codes:
								if jjj not in schedule_dict[j]:
									schedule_dict[j].append(jjj)			
						else:
							if conflict_codes not in schedule_dict[j]:
								schedule_dict[j].append(conflict_codes)
	
		#Need to make a dictionary of times for advisor preferences 
		#then use it to compare and see if there are conflicts
		advisor_schedule_dict=dict()
		for i in range(0,len(advisor_csv)):
			temp_list=[]
			#Need to check if there are any advisor times
			if len(advisor_csv[i])>=num_advisor_cols:
				advisor_times=pairing_time_dict[advisor_csv[i][advisor_id_col]]
				#check if it's a list
				if not isinstance(advisor_times,str):
					for j in range(0,len(advisor_times)):
						temp_list.append(convert_time(advisor_times[j]))
				else: #it's a string
					temp_list.append(convert_time(advisor_times))
			#If they don't have preferences, then write all the times for them
			else:
				#Redefine all_times for clarity
				temp_list=full_permutation_time()
			advisor_schedule_dict[advisor_csv[i][advisor_id_col].lower()]=copy.deepcopy(temp_list)
			
		#Delete the blank key
		if '' in advisor_schedule_dict.keys():
			del advisor_schedule_dict['']
		
		#Make a dictionary for the majors that the advisor takes		
		advisor_major_dict=dict()
		for i in range(0, len(advisor_csv)):
			temp_list=[]
			#New format, have a list in quotes
			#In theory, by splitting with respect to comma, the quotes should go away
			pref_list=advisor_csv[i][advisors_majors_col]
			if not isinstance(pref_list,str):
				for j in range(0,len(pref_list)):
					temp_list.append(convert_string_to_code(pref_list[j]))
			else:
				temp_list=convert_string_to_code(pref_list)
			advisor_major_dict[advisor_csv[i][advisor_id_col].lower()]=copy.deepcopy(temp_list)
		
		#Delete the blank key in advisor_major_dict
		if '' in advisor_major_dict.keys():
			del advisor_major_dict['']
		
		#Need to check if the advisor department conflicts with a time slot
		for i in range(0,len(advisor_csv)):
			dept = convert_department_to_code(advisor_csv[i][advisor_dept_col])
			advisor = advisor_csv[i][advisor_id_col].lower()
			advisor_times = copy.deepcopy(advisor_schedule_dict[advisor])
			#for each time, see if the department conflicts with it
			#create a list to remove later
			remove_list = []
			for j in range(0,len(advisor_times)):
				if dept in schedule_dict[advisor_times[j]]:
					remove_list.append(advisor_times[j])
			#remove the values in the remove_list
			for j in range(0,len(remove_list)):
				advisor_times.remove(remove_list[j])
					
			#check if there are no remaining times; if there are none, 
			#then give the advisor all the times
			if len(advisor_times)==0:
				output_file.write("Error: Times for "+advisor+" conflict with their department courses.  All times will be added for "+advisor+" when computing the model.\n")
				#give the advisor all the times
				advisor_schedule_dict[advisor]=full_permutation_time()
				error = 0
				
		
		#Need to make a list of all the advisors in each department
		#Initialize dictionary
		#Key is department, value is a list of advisor's lowercase NetIDs
		#Also don't add those advisors that have a pairing
		department_advisor_dict=dict()
		for i in range(0,len(advisor_csv)):
			if advisor_csv[i][advisor_dept_col] not in department_advisor_dict.keys() and advisor_csv[i][advisor_pairing_str_col]=="":
				#add the key if it's not in there
				department_advisor_dict[convert_department_to_code(advisor_csv[i][advisor_dept_col])]=[]
			department_advisor_dict[convert_department_to_code(advisor_csv[i][advisor_dept_col])].append(advisor_csv[i][advisor_id_col].lower())
			
		#Check to see if the times for this advisor overlaps with some times from other
		#advisors in their department if the advisor pairing is "Any in Department"
		#for each advisor we want to know if they want a department pairing
		for i in range(0,len(advisor_csv)):
			advisor_id = advisor_csv[i][advisor_id_col].lower()
			#Index 1 is 0/1  and equal to 1 if they want a department pairing
			if advisor_csv[i][advisor_dep_pairing_col]==1:
				#If they are the only advisor in their department then 
				#assign them to 0 and print an error saying such
				if len(department_advisor_dict[convert_department_to_code(advisor_csv[i][advisor_dept_col])])==1:
					output_file.write("Error: " + advisor_id + " is the only one in their department and can't be paired with anyone else.  The preference will be removed in computing the model.\n")
					advisor_csv[i][advisor_dep_pairing_col]=0
					error = 2
				else:
					#Figure out which are in their department 
					dept = convert_department_to_code(advisor_csv[i][advisor_dept_col])
					other_advisor_ids = copy.deepcopy(department_advisor_dict[dept])
					other_advisor_times = []
					for j in range(0,len(other_advisor_ids)):
						#Need to add the times for the other advisors in the department
						if other_advisor_ids[j]!=advisor_id:
							#check if it's a single value or a list
							if not isinstance(advisor_schedule_dict[other_advisor_ids[j]],str):
								list_to_check=advisor_schedule_dict[other_advisor_ids[j]]
								for k in range(0,len(list_to_check)):
									if list_to_check[k] not in other_advisor_times:
										other_advisor_times.append(list_to_check[k])
							else:
								if list_to_check not in other_advisor_times:
									other_advisor_times.append(list_to_check)
									
				#check if the given advisor has intersecting times with this set 
				#of times from other advisors
				if len(set(other_advisor_times).intersection(advisor_schedule_dict(advisor_id)))==0:
					output_file.write("Error: Times given by "+ advisor_id + "do not intersect with times of others in their department.  The preference will be removed in computing the model.\n")
					advisor_csv[i][advisor_dep_pairing_col]=0
					error = 2
				#else there are some times that this advisor and the others in
				#their department overlap and then we need to make sure
				#that they don't overlap with all the times that the relevant courses
				#for their department are scheduled.
				else:
					#for each of the times, check if the department is in them
					#if so, then remove it from the list
					for j in range(0,len(other_advisor_times)):
						if dept in schedule_dict[other_advisor_times[j]]:
							other_advisor_times.remove(other_advisor_times[j])
					#now check the length of the list
					#if it's 0 then print an error
					if len(other_advisor_times)==0:
						output_file.write("Error: Feasible times for " +advisor_id + " and other members in the department conflict with "+dept+ " courses.  The preference will be removed in computing the model.\n")
						advisor_csv[i][advisor_dep_pairing_col]=0
						error = 2
							
				
		#ERROR checking--Alert and ask if the user would like to continue
		check_continue=msg_ask_continue(error)
		if check_continue==False:
			return None
			
		#ERROR checking--don't want to print the .dat file if there are errors
		#exit and don't finish
		if printing=="off":
			if error!=0:
				return None
		
	return 0
		
		
if __name__ == '__main__':
	main(student_file_points,advisor_file,course_conflict_list,full_data_file,ampl_file)