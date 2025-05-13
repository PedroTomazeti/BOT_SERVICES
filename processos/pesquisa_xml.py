import os
import sqlite3
import xml.etree.ElementTree as ET
from datetime import datetime
from processos.extrair_os import extrair_os

# Função para ajustar o número da nota com base na unidade
def ajustar_numero_nota(numero_nota, unidade):
    '''
    Função para ajustar o número da nota com base na unidade
    '''
    if unidade == "Filial I":  # São Luís
        return numero_nota.lstrip('0').zfill(6)  # Remove zeros à esquerda
    elif unidade == "Filial II":  # Parauapebas
        return numero_nota[-5:]  # Extrai os últimos 5 dígitos
    else:
        raise ValueError("Unidade desconhecida")

# Função para verificar a existência da NFE correspondente
def verificar_nfe(numero_nota, pasta_nfe, unidade):
    '''
    Função para verificar a existência da NFE correspondente
    '''
    numero_nota_ajustado = ajustar_numero_nota(numero_nota, unidade)

    for arquivo in os.listdir(pasta_nfe):
        if unidade == "Filial I":
            numero_nota_ajustado = numero_nota_ajustado.lstrip('0')
        elif unidade == "Filial II":
            numero_nota_ajustado = numero_nota_ajustado.lstrip('0').zfill(4)

        if arquivo.startswith(f"NFE {numero_nota_ajustado}"):
            if arquivo.endswith('X.pdf'):
                return "Inserido", numero_nota_ajustado  # Retorna o número formatado
            return "Encontrado", numero_nota_ajustado
    return "Não encontrado", numero_nota_ajustado  # Valor padrão se nenhuma condição for atendida

# Função para processar um arquivo XML e salvar as informações no banco de dados
def processar_xml(caminho_arquivo, pasta_nfe, log_queue, unidade):
    '''
    Função para processar um arquivo XML e salvar as informações no banco de dados
    '''
    tree = ET.parse(caminho_arquivo)
    root = tree.getroot()

    # Extrair unidade, ano e mês do caminho do arquivo
    partes_caminho = caminho_arquivo.split(os.sep)

    # Definir namespace com base na unidade
    if unidade == "Filial I":  # São Luís
        ns = {'ns': 'http://www.ctaconsult.com/nfse'}
        cidade = partes_caminho[-5].split('Filial I')[-1].strip()  # Extrair apenas o nome da unidade (após "Filial I")
    elif unidade == "Filial II":  # Parauapebas
        ns = {'ns': 'http://www.abrasf.org.br/nfse.xsd'}
        cidade = partes_caminho[-5].split('Filial II')[-1].strip()  # Extrair apenas o nome da unidade (após "Filial II")
    else:
        raise ValueError("Unidade desconhecida")

    ano = partes_caminho[-4].split()[-1]  # Pegar o último elemento da string do ano (somente "2025")
    mes = partes_caminho[-3].split('-')[1].strip()  # Pegar apenas o nome do mês

    # Conectar ao banco de dados SQLite baseado na unidade
    nome_banco = f"notas_{cidade.lower().replace(' ', '_')}.db"
    con = sqlite3.connect(nome_banco)
    cur = con.cursor()

    # Criar tabela para o mês, se não existir
    tabela_mes = mes.lower()
    cur.execute(f'''
        CREATE TABLE IF NOT EXISTS "{tabela_mes}_{ano}" (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_emissao TEXT,
            numero_nota TEXT UNIQUE,  -- Chave única para evitar duplicatas
            tomador_cnpj TEXT,
            tipo_tributacao TEXT,
            tipo_recolhimento TEXT,
            natureza TEXT,
            ordem_servico TEXT,
            tipo_pagamento TEXT,
            valor_total REAL,
            status_nfe TEXT
        )
    ''')

    # Recuperar dados necessários do XML
    if unidade == "Filial I":  # São Luís
        numero_nota = root.findtext('ns:numeroNota', namespaces=ns)
        dt_emissao_raw = root.findtext('ns:dtEmissao', namespaces=ns)
        dt_emissao = datetime.strptime(dt_emissao_raw.split('T')[0], '%Y-%m-%d').strftime('%d/%m/%Y')

        tomador = root.find('ns:tomador', namespaces=ns)
        tomador_cnpj = tomador.findtext('ns:cnpj', namespaces=ns) if tomador is not None else None

        detalhamento = root.find('ns:detalhamentoNota', namespaces=ns)
        descricao_nota = detalhamento.findtext('ns:descricaoNota', namespaces=ns) if detalhamento is not None else ''
        print(f"\n\n{descricao_nota}")
        valor_total = float(detalhamento.findtext('ns:totais/ns:valotTotalNota', namespaces=ns)) if detalhamento is not None else 0.0

        atividade_executada = root.find('ns:atividadeExecutada', namespaces=ns)
        tipo_recolhimento = atividade_executada.findtext('ns:tipoRecolhimento', namespaces=ns) if atividade_executada is not None else None
        tipo_tributacao = root.findtext('ns:atividadeExecutada/ns:tipoTributacao', namespaces=ns)
        tipo_recolhimento = root.findtext('ns:atividadeExecutada/ns:tipoRecolhimento', namespaces=ns)
        natureza = '30102011' if tipo_recolhimento == 'RETIDO' else '30102002'
        # Ajustar número da nota
        numero_nota_ajustado = ajustar_numero_nota(numero_nota, unidade)

    elif unidade == "Filial II":  # Parauapebas
        inf_nfse = root.find('.//ns:InfNfse', namespaces=ns)
        numero_nota = inf_nfse.findtext('ns:Numero', namespaces=ns)
        dt_emissao_raw = inf_nfse.findtext('ns:DataEmissao', namespaces=ns)
        dt_emissao = datetime.strptime(dt_emissao_raw.split('T')[0], '%Y-%m-%d').strftime('%d/%m/%Y')

        tomador = inf_nfse.find('.//ns:Tomador', namespaces=ns)
        tomador_cnpj = tomador.findtext('.//ns:Cnpj', namespaces=ns) if tomador is not None else None

        servico = inf_nfse.find('.//ns:Servico', namespaces=ns)
        discriminacao = servico.findtext('ns:Discriminacao', namespaces=ns) if servico is not None else ''
        print(f"\n\n{discriminacao}")
        valor_total = float(servico.findtext('ns:Valores/ns:ValorServicos', namespaces=ns)) if servico is not None else 0.0

        tipo_tributacao = "TRIBUTÁVEL"

        iss_retido = servico.findtext('ns:IssRetido', namespaces=ns)
        natureza = '30102012' if iss_retido == '1' else '30102003'
        tipo_recolhimento = "PRÓPRIO" if natureza == '30102003' else 'RETIDO'

        # Ajustar número da nota
        numero_nota_ajustado = ajustar_numero_nota(numero_nota, unidade)
        numero_nota_ajustado = ano + numero_nota_ajustado

    print(f"\nVerificando nota: {numero_nota}")
    cur.execute(f'''
        SELECT COUNT(*) FROM "{tabela_mes}_{ano}" WHERE numero_nota = ?
    ''', (numero_nota_ajustado,))

    count = cur.fetchone()[0]
    print(f"Contagem de notas existentes: {count}")

    if count > 0:
        # Verificar se o status precisa ser atualizado
        cur.execute(f'''
            SELECT status_nfe FROM "{tabela_mes}_{ano}" WHERE numero_nota = ?
        ''', (numero_nota_ajustado,))
        status_atual = cur.fetchone()[0]

        status_nfe, numero_format = verificar_nfe(numero_nota, pasta_nfe, unidade)

        if status_nfe != status_atual:
            cur.execute(f'''
                UPDATE "{tabela_mes}_{ano}" SET status_nfe = ? WHERE numero_nota = ?
            ''', (status_nfe, numero_nota_ajustado))
            con.commit()
            log_queue.put(f"\nStatus da nota {numero_nota} atualizado para '{status_nfe}'.")
            print(f"Status da nota {numero_nota} atualizado para '{status_nfe}'.")
        else:
            log_queue.put(f"\nNota {numero_nota} já existe na tabela '{tabela_mes}_{ano}'. Ignorando.")
            print(f"Nota {numero_nota} já existe na tabela '{tabela_mes}_{ano}'. Ignorando.")
        con.close()
        return

    # Extrair ordem de serviço e tipo de pagamento da descrição
    os_valor, pag_valor = None, None
    if unidade == "Filial I":
        descricao = descricao_nota
    elif unidade == "Filial II":
        descricao = discriminacao

    os_valor = extrair_os(descricao)
    if isinstance(os_valor, list):
        os_valor = ", ".join(map(str, os_valor))  # Converte lista em string separada por vírgula

    if 'PAGAMENTO: FATURAMENTO EM' in descricao:
        pag_valor = descricao.split('FATURAMENTO EM')[1].split()[0].strip()
    elif 'PAGAMENTO:' in descricao:
        pag_valor = descricao.split('PAGAMENTO:')[1].split()[0].strip()
    elif 'PAG.' in descricao:
        pag_valor = descricao.split('PAG.')[1].split()[0].strip()
    elif 'PAGAMENTO EM' in descricao:
        pag_valor = descricao.split('PAGAMENTO EM')[1].split()[0].strip()
    elif 'PAGAMENTO VIA' in descricao:
        pag_valor = descricao.split('PAGAMENTO VIA')[1].split()[0].strip()
    
    # Verificar status da NFE
    status_nfe, numero_format = verificar_nfe(numero_nota, pasta_nfe, unidade)

    # Adicionar log para depuração (opcional)
    log_queue.put(f"\nStatus da NFE para a nota {numero_format}: {status_nfe}")

    # Inserir os dados na tabela
    cur.execute(f'''
        INSERT INTO "{tabela_mes}_{ano}" (
            data_emissao, numero_nota, tomador_cnpj, tipo_tributacao, tipo_recolhimento,
            natureza, ordem_servico, tipo_pagamento, valor_total, status_nfe
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (dt_emissao, numero_nota_ajustado, tomador_cnpj,
        tipo_tributacao, tipo_recolhimento, natureza, os_valor, pag_valor, valor_total, status_nfe))

    # Salvar alterações e fechar a conexão
    con.commit()
    con.close()

    log_queue.put(f"Nota {numero_nota} inserida com sucesso na tabela '{tabela_mes}_{ano}'.")
    print(f"Nota {numero_nota} inserida com sucesso na tabela '{tabela_mes}_{ano}'.")

def main_xml(pasta_xml, pasta_nfe, log_queue, unidade):
    '''
    Função principal que dá início aos processos
    '''
    # Verifica se há arquivos XML na pasta
    arquivos_xml = [arquivo for arquivo in os.listdir(pasta_xml) if arquivo.endswith('.xml')]

    if not arquivos_xml:  # Se a lista de arquivos XML estiver vazia
        print("Nenhuma nota foi adicionada.")
        log_queue.put("\nNenhuma nota foi adicionada.")
    else:
        # Percorrer todos os arquivos XML na pasta e processá-los
        for arquivo in arquivos_xml:
            caminho_completo = os.path.join(pasta_xml, arquivo)
            processar_xml(caminho_completo, pasta_nfe, log_queue, unidade)
    
    log_queue.put("Processamento concluído.")
    print("\nProcessamento concluído.")