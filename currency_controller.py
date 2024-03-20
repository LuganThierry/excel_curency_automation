class CurrencyController:

    def __init__(self) -> None:
        pass
    
    @staticmethod
    def convertStringToFloat(currency):

        currency_replaced = currency.replace(',','.')

        currency_float = float(currency_replaced)

        return currency_float

    @staticmethod
    def convertCriptoToFloat(cripto):

        cripto_replaced = cripto.replace('R$ ', '').replace('.', '').replace(',', '.')

        cripto_float = float(cripto_replaced)

        return cripto_float
    
    @staticmethod
    def formatDatabasisDate(dbdate):

        return dbdate.strftime("%d/%m/%Y")    