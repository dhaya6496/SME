import tabula
import os
import logging
from logging.handlers import RotatingFileHandler

directory = os.path.dirname(os.path.realpath(__file__))
input_directory = os.path.join(directory,'Input')
log_directory = os.path.join(directory,'Logs')
output_directory = os.path.join(directory,'Output')


for pdf in os.listdir(input_directory):
    logger = logging.getLogger(f'{log_directory}\{pdf}.log')
    hdler = RotatingFileHandler(f'{log_directory}\{pdf}.log', maxBytes=1000)
    logger.addHandler(hdler)
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s  %(message)s","%Y-%m-%d %H:%M")
    hdler.setFormatter(formatter)
    logger.info(f'Opened : {pdf}')
    if 'Ushanti' in pdf:
        finance = tabula.read_pdf(input_directory+'/'+pdf,stream=True,pages=189,silent=True)
        management = tabula.read_pdf(input_directory+'/'+pdf,stream=True,pages=162,silent=True)
    
    if 'Vaxtex' in pdf:
        finance = tabula.read_pdf(input_directory+'/'+pdf,stream=True,pages=154,silent=True)
        management = tabula.read_pdf(input_directory+'/'+pdf,stream=True,pages=130,silent=True)
        
    file_name = pdf.replace('.pdf','').replace('.PDF','')
    logger.info(f'Closed : {pdf}')
    #output
    finance[0].to_excel(os.path.join(output_directory,file_name+'_finance.xlsx'),index=False)
    management[0].to_excel(os.path.join(output_directory,file_name+'_management.xlsx'),index=False)
