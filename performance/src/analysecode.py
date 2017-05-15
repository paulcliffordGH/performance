from pathlib import Path
import csv

def main():
    
    include_tag_dict = { '022325' : True,
                         '023514' : True}   
    
    p = Path('C:\\Users\\paulc\\git\\performance\\performance\\65\\srce')
        
    for file_obj in p.iterdir():
        InputFile(file_obj, include_tag_dict)
        
class InputFile():

    def __init__(self, filePathObj, include_tag_dict):
        
        
        file_obj = filePathObj.open("r", encoding="UTF-8")
        
        last_line_written = 0
        start_para_line = ''
        lines_written_in_para = False
        out_file = OutputFile(filePathObj.stem)
        
        for line_number, fileLine in enumerate(file_obj):
            tag = fileLine[:6]
            tag_found = tag in include_tag_dict
            
            if fileLine[7:11].isdigit():
                if start_para_line == '':
                    lines_written_in_para = False
                    start_para_line = fileLine
                else:
                    if lines_written_in_para:
                        out_file.write_para_end(fileLine)
                    start_para_line = ''
                        
            elif tag_found:
                if not lines_written_in_para:
                    lines_written_in_para = True
                    out_file.write_para_start(start_para_line)
                    
                if last_line_written + 1 < line_number:
                    out_file.write_separator_line()
                    
                out_file.write_line(tag, fileLine)
                
                last_line_written = line_number
                    
        file_obj.close()
        out_file.close()


class OutputFile():        
    def __init__(self, filename):
        self.file_obj = open('C:\\Users\\paulc\\Documents\\xsolcorp\\empire\\performance\\code changes\\code analaysed\\' + filename + '.csv.', 'w', newline='')
        self.writer = csv.writer(self.file_obj)
        
    def _write_row(self, csvlist):
        self.writer.writerow(csvlist)
        
    def close(self):
        self.file_obj.close()
        
    def write_separator_line(self):
        self._write_row(['','#####################################################################'])

    def write_para_start(self, fileLine):
        self.write_separator_line()
        self._write_row(['para-start',fileLine])

    def write_para_end(self, fileLine):
        self.write_separator_line()
        self._write_row(['para-end',fileLine])

    def write_line(self, tag, fileLine):
        self._write_row([tag,fileLine])

        
if __name__ == '__main__':

    main()