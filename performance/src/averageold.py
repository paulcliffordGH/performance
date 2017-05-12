from pathlib import Path
import dateutil.parser
import xlrd
import datetime

def main():
    
    def add_to_unique_job_id_list():
        def add_to_unique_job_id_list():
            #if job not found in list add it after the last highest match
            global last_pos_in_unique_job_id_list
            for i, job_id in enumerate(unique_job_id_list):
                if job_id == job_data.job_id:
                    if last_pos_in_unique_job_id_list < i:
                        last_pos_in_unique_job_id_list = i
                    return    # already in list
                
            unique_job_id_list.insert(last_pos_in_unique_job_id_list, job_data.job_id) 
            last_pos_in_unique_job_id_list += 1   
                
        global last_pos_in_unique_job_id_list
        last_pos_in_unique_job_id_list = 0    #position to add job into list if it doesn't exist
            
        for job_data in input_data.job_list:
            add_to_unique_job_id_list()  
    
    set_ignore_dict()
        
    global old_data, new_data, sum_elapsed, num_elasped
    old_data = InputOldData()
    new_data = InputNewData()
        
#list of jobs in process order
    unique_job_id_list = []  
#seed list with most possible values
    _, longest_input_data = list(old_data.date_data_dict.items())[0]  #pick one to start
    for _, input_data in old_data.date_data_dict.items():
        if len(input_data.job_list) > len(longest_input_data.job_list):
            longest_input_data = input_data
    input_data = longest_input_data
    add_to_unique_job_id_list()
    
    for _, input_data in old_data.date_data_dict.items():
        add_to_unique_job_id_list()
    
    for _, input_data in new_data.date_data_dict.items():
        add_to_unique_job_id_list()
    
    outFile = OutputFile('C:\\Users\\paulc\\Documents\\xsolcorp\\empire\\performance\\output\\overall\\allold.csv')
        
    for job_id in unique_job_id_list:
        outFile.write_value(job_id)
        sum_elapsed = datetime.timedelta(0)
        num_elasped = 0  
        for file_date in sorted(old_data.date_data_dict):
            outFile.write_data(job_id, file_date, old_data)
        if num_elasped > 0: 
            old_average = sum_elapsed/num_elasped
        else:        
            old_average = datetime.timedelta(0)
            
        sum_elapsed = datetime.timedelta(0)
        num_elasped = 0  
            
        for file_date in sorted(new_data.date_data_dict):
            outFile.write_data(job_id, file_date, new_data)
        if num_elasped > 0: 
            new_average = sum_elapsed/num_elasped
        else:        
            new_average = datetime.timedelta(0)
            
        outFile.write_comma_separatedValue(str(old_average)[:7])
        outFile.write_comma_separatedValue(str(new_average)[:7])

        if old_average > datetime.timedelta(0): 
            diff_percent = new_average/old_average
        else:        
            diff_percent = 0

        outFile.write_comma_separatedValue(diff_percent)
        
        if old_average > datetime.timedelta(minutes=10) or new_average > datetime.timedelta(minutes=10):  
            outFile.write_comma_separatedValue("Y")
                
        outFile.write_end_line()
        
                
    outFile.close()
        
class OutputFile():        
    def __init__(self, path_name):
        self.filepathobj = Path(path_name)
        self.file_obj= self.filepathobj.open("w", encoding="UTF-8")
        self.output_heading_1()
        self.output_heading_2()
        
    def close(self):
        self.file_obj.close()
        
    def write_comma_separatedValue(self, input_value):
        self.write_value(',')
        self.write_value(str(input_value))
        
    def write_value(self, value):
        self.file_obj.write(value)
        
    def write_end_line(self):
        self.write_value('\n')
        
    def write_data(self, job_id, file_date, input_data):
        
        global sum_elapsed, num_elasped
        if file_date in include_dates_dict:
            input_file = input_data.date_data_dict[file_date]
            if job_id in input_file.job_dict:
                input_data = input_file.job_dict[job_id]
                self.write_comma_separated(input_data.start_time)
                self.write_comma_separated(input_data.end_time)
                self.write_comma_separated(input_data.elapsed)
                self.write_comma_separated(input_data.duplicate)
                sum_elapsed += input_data.elapsed
                num_elasped +=1
            else:
                self.write_comma_separated('###')
                self.write_comma_separated('')
                self.write_comma_separated('')
                self.write_comma_separated('')
    
    def write_comma_separated(self, input_value):
        self.write_value(',')
        self.write_value(str(input_value))
        
    def output_heading_1(self):
        
        def write_heading():
            if file_date in include_dates_dict:
                self.write_comma_separated(file_date)
                self.write_comma_separated(file_date)
                self.write_comma_separated(file_date)
                self.write_comma_separated(file_date)
            
        self.write_value('')

        for file_date in sorted(old_data.date_data_dict):
            write_heading()

        for file_date in sorted(new_data.date_data_dict):
            write_heading()

        self.write_comma_separated('Old')
        self.write_comma_separated('New')

        self.write_end_line()

    def output_heading_2(self):
        
        def write_heading():
            if file_date in include_dates_dict:
                self.write_comma_separated('Start')
                self.write_comma_separated('End')
                self.write_comma_separated('Elapsed')
                self.write_comma_separated('Dupl')
            
        self.write_value('JOB')

        for file_date in sorted(old_data.date_data_dict):
            write_heading()

        for file_date in sorted(new_data.date_data_dict):
            write_heading()
            
        self.write_comma_separated('Average')
        self.write_comma_separated('Average')
        self.write_comma_separated('% Diff')
        self.write_comma_separated('More than 10 mins')

        self.write_end_line()


class InputData():

    def __init__(self):
        self.date_data_dict = {}
        
        
class InputOldData(InputData):

    def __init__(self):
        super().__init__()

        p = Path('C:\\Users\\paulc\\Documents\\xsolcorp\\empire\\performance\\from empire\\old logs')
        
        for file_obj in p.iterdir():
            input_data = InputOldFile(file_obj)
            self.date_data_dict[input_data.process_date] =  input_data

class InputNewData(InputData):

    def __init__(self):
        
        super().__init__()

        xlsBook = xlrd.open_workbook('C:\\Users\\paulc\\Documents\\xsolcorp\\empire\\performance\\from empire\\IPR1ADD statistics - Combined.xls')
        self.worksheet = xlsBook.sheet_by_name("Sheet1")
        
        self.current_row = 1

        while(self.current_row < self.worksheet.nrows - 1):
            input_data = InputNewFile(self)
            if input_data.job_list:  #nothing to write from data at beginning and end of file
                self.date_data_dict[input_data.job_list[0].start_date] =  input_data


class InputFile():

    global ignore_dict
    def __init__(self):
        self.job_list = []     #preserve order added
        self.job_dict = {}

    def process_job_data(self, job_data):
        if job_data.job_id not in ignore_dict:
            if job_data.job_id in self.job_dict:
                existing_job_data = self.job_dict[job_data.job_id]
                existing_job_data.duplicate = True
                if existing_job_data.elapsed < job_data.elapsed:
                    existing_job_data.start_date = job_data.start_date
                    existing_job_data.start_time = job_data.start_time
                    existing_job_data.end_time = job_data.end_time
                    existing_job_data.elapsed = job_data.elapsed
            else:
                self.job_list.append(job_data)
                self.job_dict[job_data.job_id] = job_data
        

class InputOldFile(InputFile):

    def __init__(self, filePathObj):
        super().__init__()

        file_obj = filePathObj.open("r", encoding="UTF-8")
        
        for fileLine in file_obj:
            if fileLine[31:32] == '.' and fileLine[34:35] == '.' :
                self.process_job_data(OldJobData(fileLine))
                    
        file_obj.close()

        date = self.job_list[0].start_date[3:]
        month = date[2:5]
        day =  date[:2]
        year = '20' +  date[5:]
        try_date = day + ' ' + month + ' ' + year + ' at 0:00:00AM'
        
        date_object = dateutil.parser.parse(try_date)
        self.process_date = date_object.isoformat()[:10]


class InputNewFile(InputFile):

    def __init__(self, input_new_data):
        
        super().__init__()
        
        for input_new_data.current_row in range (input_new_data.current_row,  input_new_data.worksheet.nrows):
            job_data = NewJobData(input_new_data)
            if job_data.job_id == 'PR_PRDBBING1' and self.job_list:
                return
            self.process_job_data(job_data)
        

class JobData():
    def __init__(self):
        
        self.duplicate = ''
    
class OldJobData(JobData):
    
    def __init__(self, fileLine):
        
        super().__init__()
        
        self.job_id = fileLine[:8].rstrip()
        
        if self.job_id == 'PRPD0108':
            self.job_id = 'PRPD1108###'
        
        #self.job_no =  fileLine[9:15]
        self.start_date =  fileLine[16:26]
        self.start_time =  fileLine[29:37]
        self.end_time =  fileLine[38:46]
        self.exec_time =  fileLine[47:52]
        #self.cpu_time = fileLine[53:61]
        #self.comp_code =  fileLine[63:70]
        #self.disk_excp =  fileLine[71:77]
        #self.print_lines =  fileLine[78:83]
        #self.service_units =  fileLine[84:91]
        #self.appl_sys =  fileLine[92:98]
        #self.tw =  fileLine[99:103]
        #self.srb_time =  fileLine[104:112]
        #self.tcb_time =  fileLine[113:121]
        
        if self.exec_time[2:3] == ' ':              #seconds
            if self.exec_time[3:4] == ' ':
                seconds = '0' + self.exec_time[4:]
            else:
                seconds = self.exec_time[3:]
            convTime = '00:00:' + seconds
        else:
            if self.exec_time[0:1] == ' ':
                self.exec_time =  '0' + self.exec_time[1:]

            if self.exec_time[2:3] == 'M':
                convTime = '00:' + self.exec_time[0:2] + ':' + self.exec_time[3:]
            elif self.exec_time[2:3] == 'H':
                convTime = self.exec_time[0:2] + ':' + self.exec_time[3:] + ':00'
            else:
                convTime = '00:00:' + self.exec_time[3:] + ':00'
            
        self.elapsed = dateutil.parser.parse('2017-06-05T' + convTime) - dateutil.parser.parse('2017-06-05T00:00:00')
        
        
class NewJobData(JobData):
    
    def __init__(self, input_new_data):
        
        def getTime(colx):
            CellValue = input_new_data.worksheet.cell_value(input_new_data.current_row, colx)
            if len(CellValue) == 4:
                firstDigitOfHour = '0'
            else:
                firstDigitOfHour = ''
            return firstDigitOfHour + CellValue + ':00'
                    
        super().__init__()
        
        self.job_id = input_new_data.worksheet.cell_value(input_new_data.current_row, 0).rstrip()

        if self.job_id == 'PRPD1108':
            self.job_id = 'PRPD1108###'
        
        self.start_date =  input_new_data.worksheet.cell_value(input_new_data.current_row, 1)
        self.start_time =  getTime(2)
        self.end_date =  input_new_data.worksheet.cell_value(input_new_data.current_row, 3)
        self.end_time =  getTime(4)
        
        self.elapsed = dateutil.parser.parse(self.end_date + 'T' + self.end_time)    \
                     - dateutil.parser.parse(self.start_date + 'T' + self.start_time) 
                     
        
def set_ignore_dict():
    global ignore_dict
    global include_dates_dict
    ignore_dict = {'FILEC449' : True,
                   'FILEDFVV' : True,
                   'FILED170' : True,
                   'FILED211' : True,
                   'FILEIF3A' : True,
                   'FILEIF51' : True,
                   'PRPPQRYC' : True,
                   'FILEXFER' : True,
                   'PRPA9012' : True,
                   'PRPA9112' : True,
                   'PRPD9341' : True,
                   'PRPA9013' : True,
                   'PRFDXFER' : True,
                   'PRP19117' : True,
                   'FILEE449' : True,
                   'CHEQUES_READY' : True,
                   'ING_DISB_EXE' : True,
                   'COMPLETE_STATUS_CHECK' : True,
                   'LINK_REQUESTJOB' : True}

    include_dates_dict = { '2016-06-30' : True,
                           '2016-07-29' : True,
                           '2016-08-28' : True,
                           '2016-09-30' : True,
                           '2016-10-31' : True,
                           '2016-12-01' : True,
                           '2016-12-30' : True,
                           '2017-01-31' : True,
                           '2017-02-28' : True,
                           '2017-04-28' : True}   

        
if __name__ == '__main__':

    main()