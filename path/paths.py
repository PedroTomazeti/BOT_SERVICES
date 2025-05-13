WA_FILIAL = 'wa-panel[id="COMP6002"] > wa-panel[id="COMP6004"] > wa-panel[id="COMP6020"]'
WA_DIALOG = 'wa-dialog[id="COMP7500"] > wa-panel[id="COMP7503"] > wa-panel[id="COMP7504"] '
PATH_UNIDADE = r"path/para/pdf/da/nota/de/servico/da/unidade"

paths = {
    "password_container": ".po-page-login-password-item > po-field-container > div.po-field-container > div.po-field-container-content",
    "filial_container": "pro-branch-lookup > .po-page-login-info-container > po-lookup.po-md-6 > po-field-container > .po-field-container > div.po-field-container-content",
    "enter_unidade": ".po-md-12.po-page-login-info-container.session-settings-buttons > po-button.po-md-6.session-settings-button-enter",
    "produto_item": 'wa-menu[id="COMP3061"] > wa-menu-item[id="COMP3064"] > wa-menu-item[id="COMP3065"]',
    "data_container": ".session-date-container > po-datepicker.po-md-7 > po-field-container > .po-field-container > div > .po-field-container-content-datepicker > .po-field-container-input",
    "ambiente_container": "pro-system-module-lookup > .po-page-login-info-container > po-lookup.po-md-6 > po-field-container > .po-field-container > .po-field-container-content",
    "cnpj_container": 'wa-tab-page[id="COMP6003"] > wa-panel[id="COMP6006"] > wa-panel[id="COMP6009"] > wa-text-input[id="COMP6011"]',
    "pesquisa_cnpj": 'wa-panel[id="COMP7509"] > wa-panel[id="COMP7515"] > wa-panel[id="COMP7531"] > wa-text-input[id="COMP7533"]',
    "confirma_unidade": f'{WA_FILIAL} > wa-text-input[id="COMP6022"]',
    "btn_unidade": f'{WA_FILIAL} > wa-button[id="COMP6023"]',
    "btn_ok_cnpj": f'{WA_DIALOG} > wa-panel[id="COMP7558"] > wa-button[id="COMP7564"]',
    "pesquisa_pagto": 'wa-dialog[id="COMP7500"] > wa-panel[id="COMP7502"] > wa-panel[id="COMP7509"] > wa-panel[id="COMP7511"] > wa-panel[id="COMP7531"] > wa-text-input[id="COMP7533"]',
    "btn_ok_pagto_nat": f'{WA_DIALOG} > wa-panel[id="COMP7561"] > wa-button[id="COMP7567"]'
}