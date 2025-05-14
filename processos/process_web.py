from time import sleep
import re
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from path.paths import PATH_UNIDADE

# Variável global para rastrear o número de tentativas
tentativas = 0
limite_tentativas = 5
unidades = ['0102', '0103', '0104'] # Número para pesquisar as unidades dentro do sistema(Seguir seu padrão.)

def normal_input(driver, seletor_container, seletor_field, text, tipo, log_queue):
    container = wait_for_element(driver, By.CSS_SELECTOR, seletor_container)
        
    field = wait_for_click(container, By.CSS_SELECTOR, seletor_field)

    ActionChains(driver).move_to_element(field).perform()
    
    field.clear()

    field.send_keys(text)
    print(f"{tipo} inserido(a) com sucesso.")
    log_queue.put(f"\n{tipo} inserido(a) com sucesso.")

def expand_shadow_element(driver, element):
    """Expande o Shadow DOM de um elemento"""
    return driver.execute_script("return arguments[0].shadowRoot", element)

def shadow_button(driver, shadow_host_selector, botao_selector, log_queue):
    """
    Expande um Shadow DOM e clica em um botão específico dentro dele. (Obrigado Schenaid).
    """
    print(f"Aguardando Shadow Host: {shadow_host_selector}")
    log_queue.put(f"Aguardando Shadow Host: {shadow_host_selector}")
    shadow_host = wait_for_element(driver, By.CSS_SELECTOR, shadow_host_selector)
    shadow_root = expand_shadow_element(driver, shadow_host)

    print(f"Localizando botão: {botao_selector}")
    botao = wait_for_element(shadow_root, By.CSS_SELECTOR, botao_selector)

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao)
    print("Clicando no botão...")
    driver.execute_script("arguments[0].click();", botao)

    sleep(3)

def button(driver, shadow_button, log_queue):
    """Clique em um botão dentro do Shadow DOM do elemento."""
    WebDriverWait(shadow_button, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button'))
    )
    button = shadow_button.find_element(By.CSS_SELECTOR, 'button')
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
    driver.execute_script("arguments[0].click();", button)
    print("Botão encontrado e clicado com sucesso.")
    log_queue.put("Botão encontrado e clicado com sucesso.")

def shadow_input(driver, element, text, log_queue):
    """Acessa um input dentro do Shadow DOM de um elemento."""
    input = wait_for_element(driver, By.CSS_SELECTOR, element)
    shadow = expand_shadow_element(driver, input)
    print("Shadow DOM expandido.")
    log_queue.put("Shadow DOM expandido.")
    inserir = wait_for_click(shadow, By.CSS_SELECTOR, 'input')
    sleep(1.5)

    driver.execute_script("arguments[0].focus();", inserir)
    inserir.clear()
    sleep(1.5)

    ActionChains(driver).move_to_element(inserir).perform()
    input.send_keys(text)

    sleep(1)
    
    print("\nValor inserido.")
    log_queue.put("\nValor inserido.")

    return inserir

def wait_for_element(driver, by, value, timeout=60):
    """Função para esperar um elemento aparecer no DOM"""
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def wait_for_click(host, by, value, timeout=60):
    return WebDriverWait(host, timeout).until(EC.element_to_be_clickable((by, value)))

def click_element(driver, element, timeout=60):
    """Função para esperar e clicar em um elemento"""
    WebDriverWait(driver, timeout).until(EC.element_to_be_clickable(element)).click()

def confirmar_element(driver, by, value, timeout=20):
    """
    Função pra confirmar que o elemento vai aparecer depois de apertar o botão.
    """
    for i in range(3):
        print(f"Tentativa({i+1}) para encontrar o elemento")
        elemento_confirmado = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        if elemento_confirmado:
            print("Elemento encontrado.")
            return True

def confirmar_valor(first_info, valor_desejado):
    """Função para comparar valores."""
    if str(first_info) != str(valor_desejado):
        print(valor_desejado)
        print("Valor não inserido corretamente. Refazendo processo.")
        sleep(3)
        return False
    else:
        print("Valor confirmado. Continuando processo.")
        return True

def confirmando_wa_tgrid(driver, id, posicao, valor_desejado, funcao, tipo, *args, **kwargs):
    """Função para confirmar valores de CNPJ, Forma de Pagamento e Natureza"""
    global tentativas
    tentativas += 1

    # Verificar se o número de tentativas excedeu o limite
    if tentativas > limite_tentativas:
        print("Número máximo de tentativas excedido. Encerrando o processo.")
        driver.quit()

    # Localize o elemento principal (wa-tgrid com id específico)
    grid_element = driver.find_element(By.CSS_SELECTOR, f'wa-tgrid[id="{id}"]')

    # Acesse o shadow root
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", grid_element)

    # Localize o elemento na posição especificada
    elements = shadow_root.find_elements(By.CSS_SELECTOR, '*')
    target_element = elements[posicao]

    if posicao in [41, 29, 15, 16]:
        if target_element:
            if tipo == "CNPJ":
                cnpj = valor_desejado.getCNPJ()
                cnpj_text = target_element.text.strip()
                cnpj_numbers = ''.join(filter(str.isdigit, cnpj_text))
                print("CNPJ:", cnpj_numbers)
                if cnpj_numbers != cnpj:
                    print(cnpj)
                    print("Valor incorreto. Tentando novamente...")
                    sleep(5)
                    funcao(driver, valor_desejado, *args, **kwargs)
                else:
                    print("Valor confirmado. Seguindo com a lógica...")
                    tentativas = 0  # Resetar tentativas ao sucesso
            else:
                all_text = target_element.text.strip().splitlines()
                first_info = all_text[0] if all_text else "Não encontrado"
                print("Primeira Informação:", first_info)
                if tipo == "NATUREZA":
                    natureza = valor_desejado.getNAT()
                    confirma = confirmar_valor(first_info, natureza)
                    if confirma:
                        print("Valor confirmado. Continuando processo.")
                        tentativas = 0  # Resetar tentativas ao sucesso
                    else:
                        print(natureza)
                        print("Valor não inserido corretamente. Refazendo processo.")
                        sleep(3)
                        funcao(driver, valor_desejado, *args, **kwargs)
                elif tipo == "PAGTO":
                    pagto = valor_desejado.getPAGTO()
                    confirma = confirmar_valor(first_info, pagto)
                    if confirma:
                        print("Valor confirmado. Continuando processo.")
                        tentativas = 0  # Resetar tentativas ao sucesso
                    else:
                        print(valor_desejado)
                        print("Valor não inserido corretamente. Refazendo processo.")
                        sleep(3)
                        funcao(driver, valor_desejado, *args, **kwargs)
                else:
                    unidade = unidades[valor_desejado-1]
                    confirma = confirmar_valor(first_info, unidade)
                    if confirma:
                        print("Valor confirmado. Continuando processo.")
                        tentativas = 0  # Resetar tentativas ao sucesso
                    else:
                        sleep(3)
                        funcao(driver, valor_desejado, *args, **kwargs)
        else:
            print(f"Elemento na posição {posicao} não encontrado.")

def confirmando_wa_tmsselbr(driver, id, posicao, valor_desejado, funcao, log_queue, os):
    """Função para confirmar o valor da OS."""
    global tentativas
    tentativas += 1

    # Verificar se o número de tentativas excedeu o limite
    if tentativas > limite_tentativas:
        print("Número máximo de tentativas excedido. Encerrando o processo.")
        log_queue.put("\nNúmero máximo de tentativas excedido. Encerrando o processo.")
        driver.quit()

    # Localize o elemento principal (wa-tgrid com id específico)
    grid_element = driver.find_element(By.CSS_SELECTOR, f'wa-tmsselbr[id="{id}"]')

    # Acesse o shadow root
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", grid_element)

    # Localize o elemento na posição especificada
    elements = shadow_root.find_elements(By.CSS_SELECTOR, '*')
    try:
        target_element = elements[posicao]   

        if target_element:
            all_text = target_element.text.strip().splitlines()
            first_info = all_text[0] if all_text else "Não encontrado"
            print("Primeira Informação:", first_info)
            log_queue.put(f"Primeira Informação:{first_info}.")
            osKairos = os
            if str(first_info) != str(osKairos):
                print(osKairos)
                log_queue.put(f"OS: {osKairos}")
                print("Valor não inserido corretamente. Refazendo processo.")
                log_queue.put("Valor não inserido corretamente. Refazendo processo.")
                sleep(5)
                funcao(driver, valor_desejado, log_queue)
            else:
                print("Valor confirmado. Continuando processo.")
                log_queue.put("Valor confirmado. Continuando processo.")
                tentativas = 0  # Resetar tentativas ao sucesso
        else:
            print(f"Elemento na posição {posicao} não encontrado.")
            log_queue.put(f"Elemento na posição {posicao} não encontrado.")
    
    except IndexError:
        print("Erro: Índice fora do intervalo da lista!")
        log_queue.put("\nErro: Índice fora do intervalo da lista!")
        print("Valor não inserido corretamente. Refazendo processo.")
        log_queue.put("Valor não inserido corretamente. Refazendo processo.")
        sleep(5)
        funcao(driver, valor_desejado, log_queue)

def selecionar_elemento(driver, shadow_root, element, log_queue):
    """Selecionando elemento para construir o corpo da nota."""
    try:
        # Selecionar o elemento-alvo pelo seletor específico
        target_element = shadow_root.find_element(By.CSS_SELECTOR, element)
        if target_element:
            print(f"Elemento localizado: {element}")
            log_queue.put(f"\nElemento localizado: {element}")
            sleep(1)
            # Garantir que o elemento está visível no scroll horizontal
            container = shadow_root.find_element(By.CSS_SELECTOR, 'div.horizontal-scroll')
            driver.execute_script("""
                arguments[0].scrollLeft = arguments[1].getBoundingClientRect().left - arguments[0].getBoundingClientRect().left + arguments[0].scrollLeft;
            """, container, target_element)
            sleep(0.5)  # Pequena pausa para garantir que o scroll foi executado

            # Garantir visibilidade e foco
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_element)
            sleep(1)

            # Simular cliques no elemento
            target_element.click()
            sleep(2)

            # Simular a tecla "Enter" com JavaScript
            driver.execute_script("""
                var event = new KeyboardEvent('keydown', {
                    bubbles: true,
                    cancelable: true,
                    key: 'Enter',
                    code: 'Enter',
                    keyCode: 13
                });
                arguments[0].dispatchEvent(event);
            """, target_element)
            sleep(2)

            print("Interação com o elemento concluída com sucesso.")
            log_queue.put("Interação com o elemento concluída com sucesso.")
        else:
            print(f"Elemento não encontrado: {element}")
            log_queue.put(f"Elemento não encontrado: {element}")
    except Exception as e:
        print(f"Erro ao interagir com o elemento: {e}")

def acessa_container(driver, element, selector, funcao_seguinte, log_queue, *args, **kwargs):
    try:
        # Esperar até que o elemento wa-dialog seja carregado
        dialog_element = wait_for_element(driver, By.CSS_SELECTOR, element)

        # Obter o shadowRoot do elemento
        shadow_root = expand_shadow_element(driver, dialog_element)

        if shadow_root:
            selecionar_elemento(driver, shadow_root, selector, log_queue)
            print("Seleção concluída.")
            log_queue.put("Seleção concluída.")
            funcao_seguinte(driver, log_queue, *args, **kwargs)
        else:
            print('ShadowRoot não encontrado no elemento.')
            log_queue.put("ShadowRoot não encontrado no elemento.")
    except Exception as e:
        print(f'Ocorreu um erro: {e}')

def compara_data(data_x, data_y, log_queue):
    """Comparar datas para dar continuidade na emissão das notas"""
    if data_y == data_x:
        print("Datas de emissão iguais, manter.")
        log_queue.put("Datas de emissão iguais, manter.")
        return True
    elif data_y != data_x:
        print("Mudança na data da nota detectada.")
        log_queue.put("Mudança na data da nota detectada.")
        return False
    
def confirmando_wa_tcbrowse(driver, id, posicao, valor_desejado, log_queue):
    global tentativas
    tentativas += 1

    # Verificar se o número de tentativas excedeu o limite
    if tentativas > limite_tentativas:
        print("Número máximo de tentativas excedido. Encerrando o processo.")
        log_queue.put("\nNúmero máximo de tentativas excedido. Encerrando o processo.")
        driver.quit()

    # Localize o elemento principal (wa-tgrid com id específico)
    grid_element = driver.find_element(By.CSS_SELECTOR, f'wa-tcbrowse[id="{id}"]')

    # Acesse o shadow root
    shadow_root = driver.execute_script("return arguments[0].shadowRoot", grid_element)

    # Localize o elemento na posição especificada
    elements = shadow_root.find_elements(By.CSS_SELECTOR, '*')
    target_element = elements[posicao]   

    if target_element:
        all_text = target_element.text.strip().splitlines()
        first_info = all_text[0] if all_text else "Não encontrado"
        print("Primeira Informação:", first_info)
        log_queue.put(f"Primeira Informação:{first_info}.")
        if str(first_info) != str(valor_desejado):
            print(valor_desejado)
            log_queue.put(f"Número da nota desejada: {valor_desejado}")
            print("Valor não inserido corretamente. Refazendo processo.")
            log_queue.put("Valor não inserido corretamente. Refazendo processo.")
            sleep(5)
            return False
        else:
            print("Valor confirmado. Continuando processo.")
            log_queue.put("Valor confirmado. Continuando processo.")
            tentativas = 0  # Resetar tentativas ao sucesso
            return True
    else:
        print(f"Elemento na posição {posicao} não encontrado.")
        log_queue.put(f"Elemento na posição {posicao} não encontrado.")
        return False

def clicar_elemento_shadow_dom(driver, dialog_id, browse_id, target_selector, log_queue, num_nota):
    try:
        # Localizar o elemento wa-dialog no DOM principal
        dialog_element = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f'wa-dialog[id="{dialog_id}"] > wa-tcbrowse[id="{browse_id}"]'))
        )

        print("Elemento wa-dialog encontrado.")
        log_queue.put("\nElemento wa-dialog encontrado.")

        # Obter o ShadowRoot do elemento
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", dialog_element)

        if not shadow_root:
            print("ShadowRoot não encontrado no elemento wa-dialog.")
            log_queue.put("ShadowRoot não encontrado no elemento wa-dialog.")
            return

        print("ShadowRoot expandido com sucesso.")
        log_queue.put("ShadowRoot expandido com sucesso.")

        # Tentar localizar o elemento dentro do ShadowRoot
        target_element = shadow_root.find_element(By.CSS_SELECTOR, target_selector)

        if not target_element:
            print("Elemento dentro do Shadow DOM não encontrado. Verifique o seletor.")
            log_queue.put("Elemento dentro do Shadow DOM não encontrado. Verifique o seletor.")
            return

        print("Elemento localizado dentro do Shadow DOM:", target_element)
        log_queue.put(f"Elemento localizado dentro do Shadow DOM: {target_element}")

        # Simular um clique no elemento
        ActionChains(driver).move_to_element(target_element).click().perform()

        print("Clique realizado com sucesso!")
        log_queue.put("Clique realizado com sucesso!")

        sleep(2)

        print("\nExpandindo dialog para análise do número da nota.")
        log_queue.put("\nExpandindo dialog para análise do número da nota.")
        dialog_element = wait_for_element(driver, By.CSS_SELECTOR, f'wa-dialog[id="{dialog_id}"] > wa-tcbrowse[id="{browse_id}"]')
        
        shadow_root = expand_shadow_element(driver, dialog_element)

        wa_dialog = wait_for_element(shadow_root, By.CSS_SELECTOR, target_selector)
        sleep(2)
        valor_atual = wa_dialog.text.strip()
        print(f"Valor atual do campo: {valor_atual}")
        log_queue.put(f"Valor atual do campo: {valor_atual}")
        num_nota = num_nota.strip()  # Garante que também esteja sem espaços extras

        if valor_atual == num_nota:
            print("O valor já está correto, nenhuma alteração necessária.")
            log_queue.put("O valor já está correto, nenhuma alteração necessária.")
            sleep(2)
        else:
            print("Interação com o elemento concluída com sucesso.")
            if altera_nota(driver, target_selector, num_nota, log_queue):
                print("Valor alterado com sucesso.")
                log_queue.put("Valor alterado com sucesso.")
                sleep(2)
            else:
                print("Falha ao alterar valor")
                log_queue.put("Falha ao alterar valor")

    except Exception as e:
        print(f"Erro ao interagir com o Shadow DOM: {e}")
        log_queue.put(f"Erro ao interagir com o Shadow DOM: {e}")

def verificar_situacao(driver, log_queue):
    """
    Verifica a situação de um pedido de venda com base na imagem exibida em uma interface web.

    :param driver: Instância do WebDriver do Selenium.
    :param log_queue: Fila para armazenar logs.
    :return: None
    """
    try:
        print("Aguardando tela...")
        wa_dialog = wait_for_element(driver, By.CSS_SELECTOR, 'wa-dialog[id="COMP4500"] > wa-tgrid[id="COMP4513"]')
        
        # Acessa o Shadow Root do wa-tgrid
        shadow_root_tgrid = expand_shadow_element(driver, wa_dialog)
        
        # Seleciona o elemento desejado dentro do Shadow DOM
        image_cell = wait_for_element(shadow_root_tgrid, By.CSS_SELECTOR, 'div.horizontal-scroll > table > tbody > tr#\\30 > td#\\30 > div.image-cell')
        print(f"Image cell: {image_cell}")
        
        # Obtém o atributo style
        estilo_inline = image_cell.get_attribute("style")
        print(f"Estilo inline: {estilo_inline}")

        # Extraindo a parte fixa da URL usando regex
        match = re.search(r'background-image:\s*url\("[^"]+/cache/czls4f_prod/([^"]+)"\)', estilo_inline)
        print(f"Match: {match}")

        if match:
            image_url = match.group(1)  # Obtém apenas o nome do arquivo
            status = "Desconhecido"

            # Verifica qual imagem está sendo usada
            if "br_vermelho_mdi" in image_url:
                status = "Encerrado"
            elif "br_verde_mdi" in image_url:
                status = "Em Aberto"
            elif "br_amarelo_mdi" in image_url:
                status = "Liberado"

            print("Status do Pedido de Venda:", status)
            log_queue.put(f"Status do Pedido de Venda: {status}")

            return status
        else:
            print("Não foi possível extrair a imagem.")
            log_queue.put("Não foi possível extrair a imagem.")
    
    except NoSuchElementException as e:
        print(f"Elemento não encontrado: {e}")
        log_queue.put(f"Erro: Elemento não encontrado - {e}")
    except TimeoutException as e:
        print(f"Tempo de espera excedido: {e}")
        log_queue.put(f"Erro: Tempo de espera excedido - {e}")
    except Exception as e:
        print(f"Erro inesperado: {e}")
        log_queue.put(f"Erro inesperado: {e}")

def clicar_repetidamente(driver, log_queue, id_botao, id_objetivo):
    """
    Clica no botão até que o elemento esperado apareça, com um limite de tentativas.
    """
    max_tentativas = 3
    tentativas = 0

    while tentativas < max_tentativas:
        print(f"Tentativa {tentativas + 1} de {max_tentativas}: Clicando no botão...")
        log_queue.put(f"\nTentativa {tentativas + 1} de {max_tentativas}: Clicando no botão...")

        botao = wait_for_element(driver, By.CSS_SELECTOR, id_botao)
        shadow_botao = expand_shadow_element(driver, botao)
        button(driver, shadow_botao, log_queue)  # Função que realmente clica no botão
        
        sleep(3)
        tentativas += 1
    try:
        # Espera pelo elemento esperado para confirmar que avançou
        botao_objetivo = wait_for_element(driver, By.CSS_SELECTOR, id_objetivo)
        shadow_objetivo = expand_shadow_element(driver, botao_objetivo)
        button(driver, shadow_objetivo, log_queue)  # Função que realmente clica no botão
        print("Elemento esperado encontrado! Saindo do loop.")
        log_queue.put("\nElemento esperado encontrado! Saindo do loop.")
        return True
    except TimeoutException:
        print("Número máximo de tentativas atingido, não foi possível avançar.")
        log_queue.put("\nNúmero máximo de tentativas atingido, não foi possível avançar.")
        return False
    
def definir_nfe(unidade, ano, mes):
    """Definição do caminho para a pasta a ser analisada e alterada. Considere aumentar o número de casos se necessário."""
    match unidade:
        case 1:
            caminho_pdf = f"{PATH_UNIDADE}/Notas {ano}/{mes}/02 - Serviços"
    
    return caminho_pdf

def usar_gatilho(driver, codigo, element, func, log_queue, *args):
    campo_codigo = wait_for_element(driver, By.CSS_SELECTOR, element)
    shadow_input(driver, element, codigo, log_queue)

    valor_atual = acessar_valor(campo_codigo)

    print(f"Valor atual do campo: {valor_atual}")
    log_queue.put(f"Valor atual do campo: {valor_atual}")

    if valor_atual != codigo:
        if tentar_alterar_valor(driver, campo_codigo, codigo, log_queue, element):
            print("Valor alterado com sucesso.")
            log_queue.put("Valor alterado com sucesso.")
            sleep(3)
            func(driver, *args, log_queue)
        else:
            print("Falha ao alterar valor")
            log_queue.put("Falha ao alterar valor")
    else:
        print("O valor já está correto, nenhuma alteração necessária.")
        log_queue.put("O valor já está correto, nenhuma alteração necessária.")

def altera_nota(driver, element, valor_desejado, log_queue):
    """
    Tenta alterar o valor do campo e verifica se a alteração foi bem-sucedida.
    Retorna True se a alteração for confirmada, False se falhar após todas as tentativas.
    """
    limite_tentativas = 3
    tentativas = 0

    while tentativas < limite_tentativas:
        if tentativas > 0:        
            print("\nExpandindo dialog para análise do número da nota.")
            log_queue.put("\nExpandindo dialog para análise do número da nota.")
            dialog_element = wait_for_element(driver, By.CSS_SELECTOR, f'wa-dialog[id="COMP6000"] > wa-tcbrowse[id="COMP6004"]')
                
            shadow_root = expand_shadow_element(driver, dialog_element)

            wa_dialog = wait_for_element(shadow_root, By.CSS_SELECTOR, element)
            sleep(2)
            valor_atual = wa_dialog.text.strip()
            print(f"Valor atual do campo: {valor_atual}")
            log_queue.put(f"Valor atual do campo: {valor_atual}")
            valor_desejado = valor_desejado.strip()  # Garante que também esteja sem espaços extras

            if valor_atual == valor_desejado:
                print(f"Valor alterado com sucesso: {valor_atual}")
                log_queue.put(f"Valor alterado com sucesso: {valor_atual}")
                return True  # Sucesso: o valor foi alterado corretamente
            
            if not shadow_root:
                    print("ShadowRoot não encontrado no elemento wa-dialog.")
                    log_queue.put("ShadowRoot não encontrado no elemento wa-dialog.")
                    return

            print("ShadowRoot expandido com sucesso.")
            log_queue.put("ShadowRoot expandido com sucesso.")

            # Tentar localizar o elemento dentro do ShadowRoot
            target_element = shadow_root.find_element(By.CSS_SELECTOR, element)

            if not target_element:
                print("Elemento dentro do Shadow DOM não encontrado. Verifique o seletor.")
                log_queue.put("Elemento dentro do Shadow DOM não encontrado. Verifique o seletor.")
                return

            print("Elemento localizado dentro do Shadow DOM:", target_element)
            log_queue.put(f"Elemento localizado dentro do Shadow DOM: {target_element}")

            driver.execute_script("""
                var event = new KeyboardEvent('keydown', {
                    bubbles: true,
                    cancelable: true,
                    key: 'Enter',
                    code: 'Enter',
                    keyCode: 13
                });
                arguments[0].dispatchEvent(event);
                """, target_element)
            
            sleep(1)

        campo_rotina = wait_for_element(driver, By.CSS_SELECTOR, 'wa-panel[id="COMP3053"] > wa-text-input[id="COMP3056"]')

        valor_atual = acessar_valor(campo_rotina).strip()
        print(f"Valor atual do campo: {valor_atual}")
        log_queue.put(f"Valor atual do campo: {valor_atual}")

        if valor_atual != valor_desejado:
            print("\nAlterando valor da nota.")
            log_queue.put("\nAlterando valor da nota.")
            
            inserir = shadow_input(driver, 'wa-panel[id="COMP3053"] > wa-text-input[id="COMP3056"]', valor_desejado, log_queue)

            inserir.send_keys(Keys.RETURN)  # Press Enter to submit
            print("Enter pressionado.")
            log_queue.put("Enter pressionado.")
            sleep(0.2)

            tentativas+=1
        else:
            print("O valor já está correto, nenhuma alteração necessária.")
            log_queue.put("O valor já está correto, nenhuma alteração necessária.")
            return True

def tentar_alterar_valor(driver, field, valor_desejado, log_queue, element, max_tentativas=5):
    """
    Tenta alterar o valor do campo e verifica se a alteração foi bem-sucedida.
    Retorna True se a alteração for confirmada, False se falhar após todas as tentativas.
    """
    tentativas = 0

    while tentativas < max_tentativas:
        valor_atual = str(acessar_valor(field)).strip()

        if valor_atual == valor_desejado:
            print(f"Valor alterado com sucesso: {valor_atual}")
            log_queue.put(f"Valor alterado com sucesso: {valor_atual}")
            return True  # Sucesso: o valor foi alterado corretamente

        print(f"Tentativa {tentativas+1}/{max_tentativas}: Alterando valor...")
        log_queue.put(f"Tentativa {tentativas+1}/{max_tentativas}: Alterando valor...")

        shadow_input(driver, element, valor_desejado, log_queue)

        sleep(2)  # Pequena espera para garantir que a alteração seja processada
        tentativas += 1

    print("Erro: Não foi possível confirmar a alteração após múltiplas tentativas.")
    log_queue.put("Erro: Não foi possível confirmar a alteração após múltiplas tentativas.")
    return False  # Falha: o valor não foi alterado corretamente

def acessar_valor(field):
    """
    Obtém o valor atual do campo.
    """
    return field.get_attribute("value")

def gatilho_erro(driver, log_queue):
    print("\nAcessando wa-dialog de erro no gatilho...")
    log_queue.put("\nAcessando wa-dialog de erro no gatilho...")
    shadow_button(driver, 'wa-panel[id="COMP7510"] > wa-button[id="COMP7512"]', 'button', log_queue)

def confirma_valor(driver, valor_atual, codigo, wa_dialog, log_queue, func, *args, **kwargs):
    if valor_atual == codigo:
        print("O valor já está correto, nenhuma alteração necessária.")
        log_queue.put("O valor já está correto, nenhuma alteração necessária.")
    else:
        print("\nValores diferentes.")
        log_queue.put("\nValores diferentes.")
        if tentar_alterar_valor(driver, wa_dialog, codigo, log_queue, 'wa-dialog[id="COMP7500"] > wa-text-input[id="COMP7501"]'):
            print("Valor alterado com sucesso.")
            log_queue.put("Valor alterado com sucesso.")
            func(driver, log_queue, *args, **kwargs)
        else:
            print("Falha ao alterar valor")
            log_queue.put("Falha ao alterar valor")
