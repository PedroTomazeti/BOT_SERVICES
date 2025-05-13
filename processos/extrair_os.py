import re

def extrair_os(descricao):
    '''
    Função que utiliza Regex para capturar o número da Ordem de Serviço utilizada na descrição.
    '''
    # Regex para capturar OS mencionadas explicitamente
    padrao_os = re.compile(r'(?:OS KAIRÓS|OS KAIRÓS:|OS KAIROS|OS:|OS|Nº| - )\s+((?:\d+[-\w]*)(?:\s*[&/]\s*\d+[-\w]*)*)')

    # Regex para capturar OS avulsas no formato 4 dígitos + hífen + 3 a 4 caracteres alfanuméricos
    padrao_avulso = re.compile(r'\b\d{2,4}-[A-Z0-9]{2,4}\b')

    resultados = padrao_os.findall(descricao)  # Captura todas as ocorrências de OS

    os_encontradas = set()  # Usamos um set para evitar duplicatas

    # Processa as OS capturadas com o primeiro regex
    for resultado in resultados:
        resultado = resultado.replace("&", "/")
        partes = re.findall(r'(\d{2,4}-[A-Z0-9]{2,4})', resultado)  # Agora segue a nova regra
        os_encontradas.update(partes)

    # Adiciona as OS avulsas caso ainda não tenham sido capturadas
    os_encontradas.update(padrao_avulso.findall(descricao))

    # Caso especial: várias OS com sufixo só no final, ex: 0018/0019/0020/0021-COP
    padrao_intervalo = re.findall(r'((?:\d{2,4}/)+\d{2,4})-([A-Z]{2,4})', descricao)
    for grupo, sufixo in padrao_intervalo:
        numeros = grupo.split('/')
        for numero in numeros:
            os_encontradas.add(f'{numero}-{sufixo}')

    return list(os_encontradas) if os_encontradas else None