def valve(heat_exchanger:str, G:float, mflagistone=False):
    """Определение типа клапана при полностью автоматическом заполнении

    Args:
        heat_exchanger (str): Теплообменник
        G (float): Расход жидкости в т/ч

    Returns:
        _type_: Тип клапана
    """
    if mflagistone:
        return 'С'
    else:
        G = 1000*G
        heat_exchanger = heat_exchanger[1].lower()
        match heat_exchanger:
            case 'н':
                match G:
                    case G if G<2000:
                        return 'Ш'
                    case _:
                        return 'С'
            case 'о':
                return 'С'
def scheme(heat_exchanger:str, mflagistone=False):
    """Определение схемы

    Args:
        heat_exchanger (str): Теплообменник

    Returns:
        _type_: Схема
    """
    if mflagistone:
        return '5'
    else:
        heat_exchanger = heat_exchanger[1].lower()
        match heat_exchanger:
            case 'о':
                return '3'
            case 'н':
                return '2'
