class NotaServico:
    '''
    Classe que define os parâmetros utilizados na nota de serviço.
    '''
    
    def __init__(self, cnpj, cond_pagto, natureza, osKairos, preco, numero_nota, data):
        self.cnpj = cnpj
        self.cond_pagto = cond_pagto
        self.natureza= natureza
        self.osKairos = osKairos
        self.preco = preco
        self.numero_nota = numero_nota
        self.data = data
    
    def getCNPJ(self):
        return self.cnpj

    def getPAGTO(self):
        return self.cond_pagto
    
    def getNAT(self):
        return self.natureza
    
    def getOS(self):
        return self.osKairos
    
    def getPRECO(self):
        return self.preco
    
    def getNumNOTA(self):
        return self.numero_nota
    
    def getDATA(self):
        return self.data