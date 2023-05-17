import fitz
from PyPDF4 import PdfFileReader, PdfFileWriter
from PyPDF4.generic import NameObject, IndirectObject, BooleanObject, NumberObject
import re
import dash_bootstrap_components as dbc
import dash
from dash import html, dcc
from dash.dependencies import Input, Output, State
import io
import time
import base64
external_stylesheets = [dbc.themes.SLATE]

app = dash.Dash(__name__,external_stylesheets=external_stylesheets,title='SURF Formatter')
server= app.server
app.layout = html.Div([
    html.H4("Fill out Court-Board Member Data Sheets with a member's Surf",className="header-title"),
    html.Div([html.P(className='text-info',children="To use this application, upload a Surf. A fillable PDF file with all of the available information will download automatically.",),
    dcc.Upload(children=html.Div(['Drag and Drop or ',html.A(' Select a File')]),
    id="uploader",
    className='upload-box'),
    dcc.Download(id="download-fillable"),],className='body'),
    html.Footer(className='footer',children=[html.P('This web-app was designed by A1C Brandon Freeman, SrA Andrew Hardy at SparkX Cell and SSgt Casey Gound at Cheyenne Mountain SFS. To find out more information about this project reach out to A1C Brandon Freeman ')])
    ]
)
@app.callback(Output("download-fillable", "data"),
              Input('uploader', 'contents'),
              State('uploader', 'filename'),
              prevent_initial_call=True)
def handle_upload(contents, filename):
    content_string =  base64.b64decode(contents.split(',')[1])
    Member_Data_Extractor(content_string)
    time.sleep(2)
    return dcc.send_file('./assets/output.pdf',filename=f'{filename}' )


#
#
#
# BELOW IS THE LOGIC FOR FILLING OUT THE PDF
#
#
#

class Member_Data_Extractor():
    def __init__(self, surf):
        surf=io.BytesIO(surf)
        surf = fitz.open(stream=surf, filetype='pdf')
        
        data = self.extract_text(surf)
        # print(data)
        tables = []
        member_info= self.get_member_data(data, tables)
        # print(member_info)
        self.fill_out_pdf(member_info)
        self.set_list_box_fields(member_info)
        pass
    
    def get_rid_of_spaces(self, fat_string) -> str:
        if isinstance(fat_string, str):
            fat_string=[fat_string]
        if len(fat_string) > 0:
            fat_string=fat_string[0]
            while re.search('^ ', fat_string):
                fat_string=fat_string[1:]
            while re.search(' $', fat_string):
                fat_string=fat_string[:-1]
        else: 
            fat_string=''
        return fat_string
    
    def extract_text(self, surf)-> list:
        data=[]
        for page_num in range(len(surf)):
            data.extend(surf[page_num].get_textpage().extractDICT()['blocks'])
        return data
    def extract_tables(self, surf):
        pass
    def set_list_box_fields(self, data):
        doc = fitz.open('./assets/output.pdf')
        
        for i in doc[0].widgets():
            print(i.field_name, i.field_value, 'here')

    def get_member_data(self, data, tables) -> dict:
        self.still_in_AF_flag=False
        data_dict = {
            "mbrrank":'',
            "Last Name":'',
            "First": '',
            "MI":'',
            "suf":'',
            "Gender":'',
            "Race":'',
            "Date of Birth":'',
            "Date of Rank":'',
            "Date of Separation":'',
            "TAFMS Date":'',
            "Clearance":'',
            "marital":'',
            "BaseRow1":'',
            "UnitOffice SymbolRow1":'',
            "Dropdown8":'',
            "Duty TitleRow1":'',
            "BaseRow2":'',
            "UnitOffice SymbolRow2":'',
            "Duty TitleRow2":'',
            "BaseRow3":'',
            "UnitOffice SymbolRow3":'',
            "Duty TitleRow3":'',
            "BaseRow4":'',
            "UnitOffice SymbolRow4":'',
            "Duty TitleRow4":''
            }

        # params = ['Name','SSAN','Gr/DOR','EX/RACE/ETH-GR','TAFMSD','DOB','DOS','Security Clearance','Duty title']
        #expand on filtering through array instead of listing all individually
        for block in data:
            if 'lines' in block:
                for line in block['lines']:
                    for span in line['spans']:
                        if 'Name:' in span['text']:
                                member_name = line['spans'][1]['text']
                                data_dict['MI'] = member_name.split(' ')[2][0]
                                data_dict['Last Name'] = member_name.split(' ')[0]
                                data_dict['First'] = member_name.split(' ')[1]
                        elif 'SSAN' in span['text']:
                            if 'Spouse SSAN' not in span['text']:
                                data_dict['SSAN']=line['spans'][1]['text']
                            pass
                        elif 'Gr/DOR' in span['text']:
                                data_dict['mbrrank']=line['spans'][1]['text'].split('/')[0]
                                data_dict['Date of Rank']=line['spans'][1]['text'].split('/')[1]
                        elif 'EX/RACE/ETH-GR:' in span['text']:
                                data_dict['Gender']=line['spans'][1]['text'].split('/')[0]
                                data_dict['Race']=line['spans'][1]['text'].split('/')[1]
                        elif 'TAFMSD' in span['text']:
                                data_dict['TAFMS Date']=line['spans'][1]['text']
                        elif 'DOB' in span['text']:
                                data_dict['Date of Birth']=line['spans'][1]['text']
                        elif 'DOS' in span['text']:
                                data_dict['Date of Separation']=line['spans'][1]['text']
                        elif 'SEC CLNC' in span['text']:
                                data_dict['Clearance']=line['spans'][1]['text']
                        elif 'Command:' in span['text']:
                                data_dict['Dropdown8']=line['spans'][1]['text']
                        elif 'Duty Title' in span['text']:
                                data_dict["Duty TitleRow1"]=line['spans'][1]['text']
                                self.still_in_AF_flag=True
                        elif 'Marital' in span['text']:
                                print(span['text'])
                        elif 'DEGREE' in span['text']:
                                print(line['spans'][1]['text'])


                    #print([spans for spans in line['spans'] if [x for x in spans if 'DUTY TITLE' in spans['text']]])
                    if [spans for spans in line['spans'] if [x for x in spans if 'DUTY TITLE' in spans['text']]]:
                        wait = True
                        row =0
                        after_duty_title = data[data.index(block):]
                        for block in after_duty_title:
                            this_line = ''
                            for line in block['lines']:
                                for span in line['spans']:
                                    this_line=f'{this_line} {span["text"]}'
                            if re.search('DUTY EFF DATE', this_line):
                                        wait= False 
                                        row +=1       
                            elif not wait and re.search('-[A-Z][a-z]{2}-[0-9]{4}$', self.get_rid_of_spaces(this_line)):
                                data_dict[f'Duty TitleRow{row}']= self.get_rid_of_spaces(this_line.split('  ')[1])
                                if len(this_line) > 5:
                                    data_dict[f'UnitOffice SymbolRow{row}'] = self.get_rid_of_spaces(this_line.split('  ')[2])
                                else:
                                    data_dict[f'UnitOffice SymbolRow{row}'] = self.get_rid_of_spaces(this_line.split('  ')[2].split[0])
                                data_dict[f'BaseRow{row}']= self.get_rid_of_spaces(this_line.split('  ')[-2])
                                data_dict[f'FromRow{row}']= self.get_rid_of_spaces(this_line.split('  ')[-1])
                                row +=1
                                if row > 4:
                                    wait= True
                 
        return data_dict
    def set_need_appearances_writer(self, writer: PdfFileWriter):
        # See 12.7.2 and 7.7.2 for more information: http://www.adobe.com/content/dam/acom/en/devnet/acrobat/pdfs/PDF32000_2008.pdf
        try:
            catalog = writer._root_object
            
            # get the AcroForm tree
            if "/AcroForm" not in catalog:
                writer._root_object.update({
                    NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)
                })

            need_appearances = NameObject("/NeedAppearances")
            writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
            # del writer._root_object["/AcroForm"]['NeedAppearances']
            return writer

        except Exception as e:
            print('set_need_appearances_writer() catch : ', repr(e))
            return writer
    
    def fill_out_pdf(self, data_dict) -> None:
        myfile = PdfFileReader("./assets/AFDW Court-Board Member Data Sheet.pdf")

        page= myfile.getPage(0)
        writer = PdfFileWriter()

        annotations = page['/Annots']

        for annotation in annotations:
            print(annotation.getObject())

        self.set_need_appearances_writer(writer)
        writer.updatePageFormFieldValues(page, fields=data_dict)
        writer.addPage(page)
        with open('./assets/output.pdf', 'wb') as f:
            writer.write(f)
        return


#TODO can't find marital, just number of dependants - marital is not listed
#TODO can't find work phone  - not in surf
#TODO can't find personal email - not in surf
#TODO get list of medals??

#TODO DOR and Grade are different between my surf and the example. Have written it for the example
#TODO Can G-series orders be found on this?
#TODO Is there ever anything below Duty history in a SURF?
if __name__ == '__main__':
    app.run_server(debug=False)