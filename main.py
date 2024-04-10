import streamlit as st
from modules.info_search_vector import info #Вытаскиваем информацию по вектору
from modules.info_search_vector import search_glic_ro
from modules.vector_selection import selection #Формируем полное наименование Вектора
from modules.auto_valve_scheme import valve, scheme  # для формирования имени
from modules.blank import blank_auto #Формируем Вектор, который выводится на странице
from modules.order_form import fulfil_temp # По факту, самое главное - формирует сам бланк
import io
from zipfile import ZipFile 
from modules.vector_verosa import many_bl, table_costs
import pandas as pd
from modules.import_google_table import table_vector
from modules.vector_canal import many_bl_kanal

st.set_page_config(layout="wide") 
st.session_state['developer'] = st.text_input('Введите имя разработчика') #Вводим имя инженера
main_tab = st.tabs(['По бланкам ВЕРОСА','По бланкам канального оборудования','Ручной режим']) #Возможность выбрать режим

if 'frame_google' not in st.session_state:
    #Data_frame_google = table_vector()
    Data_frame_google = pd.read_excel('\\Scheme\\Table_vector.xlsx')
    st.session_state['frame_google'] = Data_frame_google

# Номера бланка | Блок | Схема | Клапан | Сторона | Теплообменник | Расход жидкости | Доля жидкости | Резерв
# vector_scheme, vector_valve, vector_side,rezerve = vector_form_foo(ind,col,vector_scheme)


def automate_foo(ind,vector_podbor = bool):
    """ Общая основная функция для автоматического и полуавтоматического режима\n
    Единственные два отличия: вектор и резерв подбираются самостоятельно"""
    f = st.file_uploader("Перетяните бланки сюда", accept_multiple_files=True,key=ind) #Виджет ввода бланков

    #Пустая таблица-словарь для будущего общего фрейма
    Data_table = {'Название бланка':[],'Схема':[],'Клапан':[],'Сторона':[],'Теплообменник':[],'Расход жидкости':[],'Доля гликоля':[],'Резерв':[],} #Пустные столбцы

    #Попытаемся оптимизировать вычисление некоторых значений в виде сохранения данных в виде списков
    List_cblank = []
    List_BZ = []
    key_cost = []
    cvector_list = []
    rezerve_list = []
    type_scheme_list = [] #[vector_scheme,vector_valve,type_size] 
    nn = 0
    #Блок формирования Общей таблицы-фрейма



    #Исключим все бланки, где нет теплообменников
    count = len(f) #Общее колличество принятых файлов
    i = 0
    empty_list = []

    while i != count:
        current_info = info(f[i]).all_data 
        if len(current_info) == 0:
            count -=1
            empty_list.append(f[i].name)
            f.pop(i)
        else:
            i +=1
    
    if len(empty_list) != 0:
        st.write('В данных бланках теплообменники не найдены:')
        st.write(empty_list)


    #Надо оптимизировать
    for file in f: #Проходимся по каждому файлу
        current_info = info(file).all_data #Создаём экземпляр класса (all_data - список)
        BZ = current_info[0][7][1]
        List_BZ.append(BZ)
        name_file = file.name
        for cinfo in current_info: #Пробегаем по полученному списку (список набор кортежей - пар) Здесь мы работаем только с одним файлом
            cvector,vector_scheme, vector_valve, type_size = selection(scheme(cinfo[0][1]), valve(cinfo[0][1], cinfo[2][1]), cinfo[2][1], cinfo[9][1])
            cblank = (blank_auto(selection(scheme(cinfo[0][1]), valve(cinfo[0][1], cinfo[2][1]), cinfo[2][1], cinfo[9][1]), cinfo, st.session_state['developer']))
            cvector,vector_scheme, vector_valve, type_size = selection(scheme(cinfo[0][1], cblank["intermediate coolant"]), valve(cinfo[0][1], cinfo[2][1], cblank["intermediate coolant"]), cinfo[2][1], cinfo[9][1])
            cblank = (blank_auto(cvector, cinfo, st.session_state['developer'])) #Формируем тот самый Вектор, который выводится
            cblank['intermediate coolant'] = cinfo[0][1]
            vector_side = 'П' #По умолчанию сторона справа
            type_scheme =[vector_scheme,vector_valve,type_size] 
            #Попробуем создать список из словарей cblank
            
            cblank['consumption_from_blank'] = cinfo[-1][1]
            Data_table['Название бланка'].append(name_file)
            Data_table['Схема'].append(vector_scheme)
            Data_table['Клапан'].append(vector_valve)
            #Data_table['Типоразмер'].append(type_size)
            Data_table['Сторона'].append(vector_side) 
            Data_table['Теплообменник'].append(cblank['intermediate coolant'])
            Data_table['Расход жидкости'].append(cblank['consumption_from_blank'])
            Data_table['Доля гликоля'].append(cblank['glycol'])
            Data_table['Резерв'].append(0) #По умолчанию для всех резерв = 0

            type_scheme_list.append(type_scheme)
            cvector_list.append(cvector)
            List_cblank.append(cblank)
            nn +=1
            
    #Конец чтения файлов
            
    if nn > 0:
        #Блок отвечающий за подбор в полуавтоматическом режиме
        

        Data_frame = pd.DataFrame(Data_table)

        if vector_podbor: #Если это полуавтоматический режим

            Teplo_inf = Data_table['Теплообменник']

            for i in range(len(Teplo_inf)):
                Teplo_inf[i] = Teplo_inf[i][:3]

            #Множественный выбор
            col = st.columns(5) # Делим на два столбца
            #all = col[0].checkbox(label ='Множественный выбор',key=ind+790)
            col[0].write('Множественный выбор')
            all_col_vector_scheme = col[1].selectbox( 'Схема', ('По умолчанию','1', '2', '3','4','4М','5','5М','6'),key=ind+777)
            all_col_vector_valve = col[2].selectbox( 'Клапан', ('По умолчанию','С', 'Ш'),key=ind+767)
            all_col_vector_side = col[3].selectbox( 'Сторона', ('По умолчанию','П', 'Л'),key=ind+768)
            all_col_rezerve = col[4].selectbox( 'Резерв', ('По умолчанию','0', '1'),key=ind+789)


            if all_col_vector_scheme != 'По умолчанию':
                Data_frame['Схема'] = all_col_vector_scheme
            if all_col_vector_valve != 'По умолчанию':    
                Data_frame['Клапан'] = all_col_vector_valve
            if all_col_vector_side != 'По умолчанию':    
                Data_frame['Сторона'] = all_col_vector_side      
            if all_col_rezerve != 'По умолчанию':    
                Data_frame['Резерв'] =  all_col_rezerve


            #Data_frame['Схема'] = '3'

            edit_frame = st.data_editor(
            Data_frame,
            use_container_width=True,
            hide_index=True,

            column_config={
            "Резерв": st.column_config.SelectboxColumn(
                width="small",
                options=['0', '1'],
                required=True,
            ),
     
            "Схема": st.column_config.SelectboxColumn(
                width="small",
                options=['1', '2', '3','4','4М','5','5М','6'],
                required=True,
            ),

            "Клапан": st.column_config.SelectboxColumn(
                width="small",
                options=['С', 'Ш'],
                required=True,
            ),

            "Сторона": st.column_config.SelectboxColumn(
                width="small",
                options=['П', 'Л'],
                required=True,
            )
                
                },
           disabled=["Название бланка","Теплообменник","Расход жидкости","Доля гликоля"],
           #Попробуем через session state
            )


            ##Проверка на несовместимости!
            #for i in range(len(Teplo_inf)):edit_frame
            #    if ((edit_frame['Схема'].values[i]) != '3') and (Teplo_inf[i] == 'ВОВ'):
            #        st.write((Data_frame['Схема'].values[i]))
            #        st.write(Teplo_inf[i])
            #        edit_frame['Схема'].values[i] = '3'
            #        #st.warning('Внимание:', icon="⚠️❗")

        else: #Если это автоматический режим
            st.dataframe(Data_frame, use_container_width = True, hide_index = True)

        ##################################### Попробую осуществить раздельное создание фрейма и создание бланков
        #Кнопка-сформировать - для полуавтомата, затем загрузить архив
        # загрузить архив - для автомата
        #Временно вынесу все полезные переменные

        error = False
        
        #Проверка на охладители!
        for i in range(len(Teplo_inf)):
            if ((edit_frame['Схема'].values[i]) != '3') and (Teplo_inf[i] == 'ВОВ'):
                st.warning('Внимание, охладители должны быть выполнены по третьей схеме', icon="⚠️")
                break


        #Проверка на клапаны
        for i in range(len(edit_frame['Схема'])):
            if ((edit_frame['Схема'].values[i]) == '1') and (edit_frame['Клапан'].values[i] == 'Ш'):
                st.warning('Внимание, схема - 1 может быть только с седельным клапаном', icon="⚠️")
                break

        for i in range(len(edit_frame['Схема'])):
            if ((edit_frame['Схема'].values[i]) == '3') and (edit_frame['Клапан'].values[i] == 'Ш'):
                st.warning('Внимание, схема - 3 может быть только с седельным клапаном', icon="⚠️")
                break

        for i in range(len(edit_frame['Схема'])):
            if ((edit_frame['Схема'].values[i]) == '4М') and (edit_frame['Клапан'].values[i] == 'Ш'):
                st.warning('Внимание, схема - 4М может быть только с седельным клапаном', icon="⚠️")
                break

        for i in range(len(edit_frame['Схема'])):
            if ((edit_frame['Схема'].values[i]) == '6') and (edit_frame['Клапан'].values[i] == 'С'):
                st.warning('Внимание, схема - 6 может быть только с Шаровым клапаном', icon="⚠️")
                break

        #Проверка на типоразмер
        #st.write(type_scheme_list)
        #for i in range(len(edit_frame['Схема'])):
        #    type_size = type_scheme_list[i][2]
        #    type_klapan = type_scheme_list[i][1]
        #    if ((edit_frame['Схема'].values[i]) == '1') and (type_size in  {10,11}):
        #        st.error(f'Внимание,в строке {i+1} схема - 1, не соответсвует типоразмеру {type_size}', icon="❗")
        #    if ((edit_frame['Схема'].values[i]) == '6') and (type_size in  {1,2,3,9,10,11}):
        #        st.error(f'Внимание,в строке {i+1} схема - 6, не соответсвует типоразмеру {type_size}', icon="❗")
#
        #    if ((edit_frame['Схема'].values[i]) == '2') and (type_size in  {9,10,11}) and (type_klapan == 'Ш'):
        #        st.error(f'Внимание,в строке {i+1} схема  - 2 с шаровым клапаном, не соответсвует типоразмеру {type_size}', icon="❗")
        #    if ((edit_frame['Схема'].values[i]) == '4') and (type_size in  {9,10,11}) and (type_klapan == 'Ш'):
        #        st.error(f'Внимание,в строке {i+1} схема  - 2 с шаровым клапаном, не соответсвует типоразмеру {type_size}', icon="❗")
        #    if ((edit_frame['Схема'].values[i]) == '5') and (type_size in  {9,10,11}) and (type_klapan == 'Ш'):
        #        st.error(f'Внимание,в строке {i+1} схема  - 2 с шаровым клапаном, не соответсвует типоразмеру {type_size}', icon="❗")
        #    if ((edit_frame['Схема'].values[i]) == '5М') and (type_size in  {9,10,11}) and (type_klapan == 'Ш'):
        #        st.error(f'Внимание,в строке {i+1} схема  - 2 с шаровым клапаном, не соответсвует типоразмеру {type_size}', icon="❗")





        if st.button("Сформировать", type="primary",key=ind+999,disabled=error):

            #Повторная проверка на охладители, если что не так, то исправляем!
            one_and_stop = True #Выводим сообщение только один раз

            for i in range(len(Teplo_inf)):
                if ((edit_frame['Схема'].values[i]) != '3') and (Teplo_inf[i] == 'ВОВ'):
                    edit_frame['Схема'].values[i] = '3'
                    if one_and_stop:
                        st.info('Внимание, охладители должны быть выполнены по третьей схеме, при формировании бланков были внесены корректировки', icon="ℹ")
                        one_and_stop = False


            #Проверка на клапаны, если что не так, то исправляем!
            one_and_stop = True
            for i in range(len(edit_frame['Схема'])):
                if ((edit_frame['Схема'].values[i]) == '1') and (edit_frame['Клапан'].values[i] == 'Ш'):
                    edit_frame['Клапан'].values[i] = 'С'
                    if one_and_stop:
                        st.info('Внимание, схема - 1 может быть только с седельным клапаном, при формировании бланков были внесены корректировки', icon="ℹ")
                        one_and_stop = False

            one_and_stop = True
            for i in range(len(edit_frame['Схема'])):
                if ((edit_frame['Схема'].values[i]) == '3') and (edit_frame['Клапан'].values[i] == 'Ш'):
                    edit_frame['Клапан'].values[i] = 'С'
                    if one_and_stop:
                        st.info('Внимание, схема - 3 может быть только с седельным клапаном, при формировании бланков были внесены корректировки', icon="ℹ")
                        one_and_stop = False

            one_and_stop = True
            for i in range(len(edit_frame['Схема'])):
                if ((edit_frame['Схема'].values[i]) == '4М') and (edit_frame['Клапан'].values[i] == 'Ш'):
                    edit_frame['Клапан'].values[i] = 'С'
                    if one_and_stop:
                        st.info('Внимание, схема - 4М может быть только с седельным клапаном, при формировании бланков были внесены корректировки', icon="ℹ")
                        one_and_stop = False

            one_and_stop = True
            for i in range(len(edit_frame['Схема'])):
                if ((edit_frame['Схема'].values[i]) == '6') and (edit_frame['Клапан'].values[i] == 'С'):
                    edit_frame['Клапан'].values[i] = 'Ш'
                    if one_and_stop:
                        st.info('Внимание, схема - 6 может быть только с Шаровым клапаном, при формировании бланков были внесены корректировки', icon="ℹ")
                        one_and_stop = False












            if vector_podbor:
                for i in range(nn):
                    vector_scheme = edit_frame['Схема'].values[i]
                    vector_valve = edit_frame['Клапан'].values[i]
                    vector_side = edit_frame['Сторона'].values[i]
                    rezerve = edit_frame['Резерв'].values[i]
                    rezerve_list.append(rezerve)

                    cblank = List_cblank[i]
                    cvector,vector_scheme, vector_valve, type_size = selection(vector_scheme, vector_valve, cblank["consumption"], vector_side)
                    type_scheme =[vector_scheme,vector_valve,type_size]
                    cblank["vector"] = cvector

                    List_cblank[i] = cblank
                    type_scheme_list[i] = type_scheme
                    cvector_list[i] = cvector

            progress_text = "Пожалуйста, подождите"
            my_bar = st.progress(0, text=progress_text)

            name_archive = ''
            Archive = io.BytesIO()
            #vals = {} #Создаём словарь для передачи в many_bl
            #Фреймы для таблиц цен и векторов
            all_df_cost = {}
            all_df_cost = pd.DataFrame(all_df_cost)
            all_df_verosa = {}
            all_df_verosa = pd.DataFrame(all_df_verosa)


            i = 0
            j = 0
            
            #part = round(90  / nn)
            part = 1 / nn

            for file in f: #Проходимся по каждому файлу
                name_file = file.name
                cvector_list_file = []
                key_cost = []
                rezerve_list_file = []
                vals = {} #Создаём словарь для передачи в many_bl
                current_info = info(file).all_data #Создаём экземпляр класса (all_data - список)
                k = 0 #Для vals
                #st.write(name_file)
                for cinfo in current_info: 
                    #Должно быть в блоке cinfo
                    # Блок формирования бланков docx и добавления в архив (Выполняется после добавления всего в архив)
                    rez_sheme = False
                    if vector_podbor:
                        #st.write(rezerve_list[i])
                        rezerve_list_file.append(rezerve_list[i])
                        rez_sheme = rezerve_list[i]
                    
                    

                    #Вытаскиваем переменные из списков type_scheme_list = [] #[vector_scheme,vector_valve,type_size] 
                    cblank = List_cblank[i]
                    type_scheme = type_scheme_list[i]
                    cvector = cvector_list[i]
                    #cblank["vector"] = cvector


                    bio = io.BytesIO()
                    document = fulfil_temp(cblank,type_scheme,st.session_state['frame_google'] ,rez_sheme,st.session_state['developer'])
                    document.save(bio)
                    
                    name = f"{cblank['order form']}-{cblank['system']}-{cvector}.docx"
                    with ZipFile(Archive, mode='a') as archive:
                        archive.writestr(name, bio.getvalue())
                    name_archive = cblank['orderer']+'.zip'
                    ###################################

                    #Здесь формируются параметры для передачи в vector_verosa (для формирования таблиц)
                    key_cost.append(f'{type_scheme[0]}-{type_scheme[1].upper()}-{type_scheme[2]}')
                    cvector_list_file.append(cvector)
                    #if vector_podbor:
                    #    rezerve_list.append(rezerve)
                    vals[f'scheme{k}'] = scheme(cinfo[0][1], cblank["intermediate coolant"])
                    vals[f'valve{k}'] = valve(cinfo[0][1], cinfo[2][1], cblank["intermediate coolant"])
                    vals[f'side{k}'] = cinfo[9][1]
                    vals[f'rez{k}'] = False
                    k +=1
                    ###################################
                    i+=1 #Просто счётчик
                    my_bar.progress(i*part, text=progress_text)
                    ###################################
                BZ = List_BZ[j]
                #Здесь формируются таблицы (а точнее фреймы) Выполняется во внеШнем цикле file
                df_verosa, df_cost = many_bl(file,vals,name_file,BZ,key_cost,cvector_list_file,vector_podbor,rezerve_list_file) 
                all_df_cost = pd.concat([all_df_cost,df_cost])
                all_df_verosa = pd.concat([all_df_verosa,df_verosa])
                j +=1

            if j == len(f):
                #БЛОК ЗАПОЛНЕНИЯ ТАБЛИЦ (внеШний цикл)
                output = io.BytesIO()
                all_df_verosa.to_excel(output, index=False)
                name = 'ВЕКТОР ВЕРОСА.xlsx'
                with ZipFile(Archive, mode='a') as archive:
                    archive.writestr(name,output.getvalue())




                output = io.BytesIO()
                all_df_cost.to_excel(output, index=False)
                name = 'ВЕКТОР ЦЕНЫ.xlsx' 
                with ZipFile(Archive, mode='a') as archive:
                    archive.writestr(name,output.getvalue())
                my_bar.empty()
                st.download_button('💾Загрузить Архив: ', data=Archive.getvalue(), file_name=name_archive,key=ind+6) # Для каждого файла свой архив (относится к внеШнему циклу)

def vector_form_foo(ind,col,vector_scheme):
    col[0].write(f'Узелрегулирующий ВЕКТОР:')
    col[1].write(' ')
    col[1].write(' ')
    col[1].write(' ')
   

    origin = 0
    match vector_scheme:
        case "1":
            origin = 0
        case "2":
            origin = 1
        case "3": 
            origin = 2
        case "4":
            origin = 3
        case "4М":
            origin = 4
        case "5":
            origin = 5
        case "5М":
            origin = 6
        case "6":
            origin = 7


    vector_scheme = col[0].selectbox( 'Схема', ('1', '2', '3','4','4М','5','5М','6'),key=ind+111,index=origin)
    if vector_scheme in ('1','3','4М'):
        vector_valve = 'С'
        col[0].write('Клапан: С')
    elif vector_scheme in ('6','6М'):
        vector_valve = 'Ш'
        col[0].write('Клапан: Ш')
    else:
        vector_valve = col[0].selectbox( 'Клапан', ('С', 'Ш'),key=ind+222)

    rezerve = col[1].checkbox(label ='Резерв',key=ind+444)
    if rezerve:
        rezerve = 1
    else:
        rezerve = 0
    vector_side = col[1].selectbox( 'Сторона подключения', ('Л', 'П'),key=ind+333,index=1)
    return vector_scheme, vector_valve, vector_side, rezerve

def semi_auto_foo(col,cblank,ind,vector_scheme):
    # Номера бланка | Блок | Схема | Клапан | Сторона | Теплообменник | Расход жидкости | Доля жидкости | Резерв

    origin = 0
    match vector_scheme:
        case "1":
            origin = 0
        case "2":
            origin = 1
        case "3": 
            origin = 2
        case "4":
            origin = 3
        case "4М":
            origin = 4
        case "5":
            origin = 5
        case "5М":
            origin = 6
        case "6":
            origin = 7



    vector_scheme = col[1].selectbox( '', ('1', '2', '3','4','4М','5','5М','6'),key=ind+111,index=origin)
    if vector_scheme in ('1','3','4М'):
        vector_valve = 'С'
        col[2].write('С')
    elif vector_scheme in ('6','6М'):
        vector_valve = 'Ш'
        col[2].write('Ш')
    else:
        vector_valve = col[2].selectbox('', ('С', 'Ш'),key=ind+222)

    rezerve = col[7].checkbox(label = '',key=ind+444)
    if rezerve:
        rezerve = 1
    else:
        rezerve = 0
    vector_side = col[3].selectbox( '', ('Л', 'П'),key=ind+333,index=1)
    return vector_scheme, vector_valve, vector_side, rezerve

#with main_tab[0]: # Для автоматического режима 
#    automate_foo(0,False)

with  main_tab[2]: # Для ручного режима
    col = st.columns(2) # Делим на два столбца

    cblank = {} #Создаём пустой бланк
    Archive = io.BytesIO()
    #Основная юридическая информация
    col[0].write('Основная юридическая информация')
    cblank["orderer"] = col[0].text_input('Заказчик:')
    cblank["object"] = col[0].text_input('Объект:')
    cblank["manager"] = col[0].text_input('Менеджер:')
    cblank["order form"] = col[0].text_input('Номер БЗ:')
    cblank["developer"] = col[0].text_input('Выполнил:')

    #Исходные данные и режимы работы 
    col[1].write('Исходные данные и режимы работы')
    cblank["temperature"] = col[1].number_input('Температура на входе:')
    cblank["consumption"] = col[1].number_input('Расход теплоносителя, т/ч')
    glycol_hidden = False
    cblank["glycol type"] = col[1].radio("Выберите тип теплоносителя",('Этиленгликоль', 'Пропиленгликоль', 'Вода, соответствует требованиям СП 124.13330.2012'))
    if  cblank["glycol type"] == 'Вода, соответствует требованиям СП 124.13330.2012':
        glycol_hidden = True
        cblank["glycol"] = 0
    cblank["glycol"] = col[1].number_input(label ='Cодержание гликоля %',disabled =glycol_hidden)


    #Формирование поправочной цены

    realG = cblank["consumption"] /search_glic_ro(cblank["temperature"], cblank["glycol"]/100)
    cblank['consumption_from_blank'] = cblank["consumption"]
    cblank["consumption"] = realG
    #Формирование вектора

    
    vector_scheme, vector_valve, vector_side,rezerve = vector_form_foo(10,col,1)
    cvector,vector_scheme, vector_valve, type_size = selection(vector_scheme, vector_valve, cblank["consumption"], vector_side)
    #st.write(f'{cvector}')



    cblank["vector"] = cvector
    if rezerve:
        cblank["reserve"] = 1
    else:
        cblank["reserve"] = 0



    type_scheme =[vector_scheme,vector_valve,type_size]
    developer_name = st.session_state['developer']
    table_spec = st.session_state['frame_google']
    document = fulfil_temp(cblank,type_scheme,table_spec,cblank["reserve"],developer_name)
    key_cost = [f'{vector_scheme}-{vector_valve}-{type_size}']
    bio = io.BytesIO()
    document.save(bio)
    name = f"{cvector}.docx"
    with ZipFile(Archive, mode='a') as archive:
        archive.writestr(name, bio.getvalue())



    #Формирование таблиц
    allsis={"Бланк": [cblank["orderer"]],
                "Расход жидкости": [cblank["consumption"]],
                "Температура на входе": [cblank["temperature"]],
                "Доля гликоля": [cblank["glycol"]],
                "Расход поправочный": realG,
                "Вектор": [cblank["vector"]],
                "Номер БЗ": [cblank["order form"]],
                "Резерв": [cblank["reserve"]]}
        
    df_verosa = pd.DataFrame(allsis)  
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_verosa.to_excel(output, index=False)
    name = 'ВЕКТОР ВЕРОСА.xlsx'
    with ZipFile(Archive, mode='a') as archive:
        archive.writestr(name,output.getvalue())
    BZ = [cblank["order form"]]
    rezerve = [str(cblank["reserve"])]
    glycol = [cblank["glycol"]]
    blank = [cblank["orderer"]]
    block = ['-']
    df_cost = table_costs(BZ, key_cost, rezerve,glycol,blank,block)
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_cost.to_excel(output, index=False)
    name = 'ВЕКТОР ЦЕНЫ.xlsx' 
    with ZipFile(Archive, mode='a') as archive:
        archive.writestr(name,output.getvalue())
    name_archive = cblank['order form']+'.zip'
    st.download_button('💾Загрузить Архив: ', data=Archive.getvalue(), file_name=name_archive) # Для каждого файла свой архив (относится к внеШнему циклу)

with  main_tab[0]: #Полуавтоматический режим: ВЕКТОК-схема-клапан-сторона и резерв
    automate_foo(400,True)

with main_tab[1]: #Каналка
    f = st.file_uploader("Перетяните бланки сюда", accept_multiple_files=True) #Виджет ввода бланков()
    many_bl_kanal(f,231023,st.session_state['developer'])
