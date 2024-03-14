from docx import Document as d
#from assitent_info import *
def blank_auto(cvector, cinfo, developer):
    """ Автозаполнение словаря"""
    i_for_blank = {}
    i_for_blank['vector'] = cvector
    i_for_blank['heat_exchanger'] = cinfo[0][1]
    i_for_blank['temperature'] = cinfo[1][1]
    i_for_blank['consumption'] = cinfo[2][1]
    i_for_blank['consumption_from_blank'] = cinfo[-1][1]
    i_for_blank['glycol'] = cinfo[3][1]
    i_for_blank['orderer'] = cinfo[4][1]
    i_for_blank['object'] = cinfo[5][1] if len(cinfo[5][1])<40 else '\n'.join([' '.join(cinfo[5][1].split()[:3]), ' '.join(cinfo[5][1].split()[3:])]) if len(cinfo[5][1])<80 else '\n'.join([' '.join(cinfo[5][1].split()[:3]), ' '.join(cinfo[5][1].split()[3:6]), ' '.join(cinfo[5][1].split()[6:])])
    i_for_blank['manager'] = cinfo[6][1]
    i_for_blank['order form'] = cinfo[7][1]
    i_for_blank['system'] = cinfo[8][1]
    i_for_blank['developer'] = developer
    # i_for_blank['side'] = cinfo[9][1]
    i_for_blank['intermediate coolant'] = cinfo[10][1]
    i_for_blank['glycol type'] = cinfo[11][1]
    return i_for_blank
   
