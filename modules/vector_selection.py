def selection(scheme:str, valve:str, G:float, side:str):
    """Подбор ВЕКТОР по параметрам

    Args:
        scheme (str): Схема
        valve (str): Тип клапана
        G (float): Расход жидкости в т/ч
        side (str): Сторона подключения

    Returns:
        (str): Наименование ВЕКТОР
    """
    valve = valve.lower()
    G = 1000*G
    def TIS(G:float): # Определяем типоразмер для схемы с седельным клапаном
        match G:
            case G if G < 500:
                return "1"
            case G if 500 <= G < 800:
                return "2"
            case G if 800 <= G < 1000:
                return "3"
            case G if 1000 <= G < 2000:
                return "4"
            case G if 2000 <= G < 3500:
                return "5"
            case G if 3500 <= G < 6000:
                return "6"
            case G if 6000 <= G < 10000:
                return "7"
            case G if 10000 <= G < 14000:
                return "8"
            case G if 14000 <= G < 22000:
                return "9"
            case G if 22000 <= G < 40000:
                return "10"
            case G if 40000 <= G < 60000:
                return "11"
    def TIH(G:float): # Определяем типоразмер для схемы с шаровым клапаном
        match G:
            case G if G < 200:
                return "1А"
            case G if 200 <= G < 300:
                return "1Б"
            case G if 300 <= G < 500:
                return "1"
            case G if 500 <= G < 800:
                return "2"
            case G if 800 <= G < 1000:
                return "3"
            case G if 1000 <= G < 2000:
                return "4"
            case G if 2000 <= G < 3500:
                return "5"
            case G if 3500 <= G < 6000:
                return "6"
            case G if 6000 <= G < 10000:
                return "7"
            case G if 10000 <= G < 14000:
                return "8"
    match valve:
        case 'ш':
            TR = TIH(G)
        case 'с':
            TR = TIS(G)
    return f"ВЕКТОР-{scheme}-{valve}-{TR}-{side}-С+".upper() , scheme, valve, TR
