# Cornell_Engineering_Advising_Project
Code repository for data processing steps for the Cornell Engineering Advising Capstone project (MEng Class of 2017)

The main tool and data processing is done in Microsoft Excel (2016) and Visual Basic for Applications.  Then, the Python routine(s) are called to process the data and convert it to a .dat file.  This .dat file is used in conjunction with a .mod and .run file in AMPL, an optimzation tool (http://ampl.com/) to run the model.  

This model returns 3 files: 
1. File matching the students to advisors
2. File matching the advisors to time slots (predefined set accoring to Cornell scheduling practices)
3. File containing demographic information, the student matching, advisor matching, and the majors that they matched on, among other metrics

Then there is more VBA executed to import the data back into Excel for easier visualization and aggregation.
