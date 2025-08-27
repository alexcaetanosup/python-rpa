from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
from botcity.web import *
from datetime import datetime
from botcity.plugins.excel import *
import logging
from logging.handlers import RotatingFileHandler

class Bot:
    def bot(self):
        # Sequence: Cadastro de Fornecedores (sequence)

        #  Logger Config Activity
        # Displayname: LoggerConfig
        loggerBot = logging.getLogger("Cadastro de Fornecedores")
        loggerBot.setLevel(logging.DEBUG)
        filelogging = RotatingFileHandler("fileLogging.log", maxBytes = 20000, backupCount = 10)
        filelogging.setLevel(logging.DEBUG)
        formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
        filelogging.setFormatter(formatter)
        loggerBot.addHandler(filelogging)

        # Open Browser Activity
        loggerBot.info("Start: Abrir o cadastro de fornecedore (OpenBrowser)")

        # Displayname: Abrir o cadastro de fornecedore (OpenBrowser)
        webDriverPath = ChromeDriverManager().install()
        webBot = WebBot()
        webBot.driver_path = webDriverPath
        webBot.browser = Browser.CHROME
        webBot.headless = False
        webBotDef_options = default_options()
        webBotDef_options.add_argument("--page-load-strategy=Normal")
        webBot.options = webBotDef_options
        webBot.browse("https://jornadarpa.com.br/alunos/desafios/cadfor25")

        loggerBot.info("End: Abrir o cadastro de fornecedore (OpenBrowser)")

        loggerBot.info("Start: Mapeamento dos elementos da pagina de login (Element_Library)")

        # DisplayName: Mapeamento dos elementos da pagina de login (Element_Library)

        # Sequence: Mapeamento da lista de elementos (Element list)

        # Find Element Activity
        loggerBot.info("Start: Mapear Campo Usuário (find elements)")

        # Displayname: Mapear Campo Usuário (find elements)
        campoEmail = webBot.find_element(selector="usuario", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear Campo Usuário (find elements)")

        # Find Element Activity
        loggerBot.info("Start: Mapear Campo Senha (find elements)")

        # Displayname: Mapear Campo Senha (find elements)
        campoSenha = webBot.find_element(selector="senha", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear Campo Senha (find elements)")

        # Find Element Activity
        loggerBot.info("Start: Mapear Botão Lgpd (find elements)")

        # Displayname: Mapear Botão Lgpd (find elements)
        botalLgpd = webBot.find_element(selector="lgpd", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear Botão Lgpd (find elements)")

        # Find Element Activity
        loggerBot.info("Start: Mapear Botão Login (find elements)")

        # Displayname: Mapear Botão Login (find elements)
        botaoLogin = webBot.find_element(selector="btnLogin", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear Botão Login (find elements)")

        loggerBot.info("End: Mapeamento dos elementos da pagina de login (Element_Library)")

        loggerBot.info("Start: Entrada de Dados de Login (Form_Element_Action_Web)")

        # DisplayName: Entrada de Dados de Login (Form_Element_Action_Web)

        # Sequence: Action list

        # Type Into Activity
        loggerBot.info("Start: Digitação do Campo Email (Type Into campoEmail field)")

        # Displayname: Digitação do Campo Email (Type Into campoEmail field)
        campoEmail.send_keys("participante@desafiosrpa.com.br")

        loggerBot.info("End: Digitação do Campo Email (Type Into campoEmail field)")

        # Type Into Activity
        loggerBot.info("Start: Digitação do Campo Senha (Type Into campoSenha field)")

        # Displayname: Digitação do Campo Senha (Type Into campoSenha field)
        campoSenha.send_keys("evento")

        loggerBot.info("End: Digitação do Campo Senha (Type Into campoSenha field)")

        # Click Activity
        loggerBot.info("Start: Click in botalLgpd element")

        # Displayname: Click in botalLgpd element
        botalLgpd.click()

        loggerBot.info("End: Click in botalLgpd element")

        # Click Activity
        loggerBot.info("Start: Click in botaoLogin element")

        # Displayname: Click in botaoLogin element
        botaoLogin.click()

        loggerBot.info("End: Click in botaoLogin element")

        loggerBot.info("End: Entrada de Dados de Login (Form_Element_Action_Web)")

        # Read Excel Activity
        loggerBot.info("Start: Lear a planilha excel com os dados (Read_Excel)")

        # Displayname: Lear a planilha excel com os dados (Read_Excel)
        excelBot = BotExcelPlugin()
        file_or_path = "C:\\RPA\\Fornecedores\\Lista_exemplo.xlsx"

        listaFornecedores = excelBot.read(file_or_path=file_or_path).as_list(sheet="lista")[1:]
        loggerBot.info("End: Lear a planilha excel com os dados (Read_Excel)")

        loggerBot.info("Start: Element_Library")

        # DisplayName: Element_Library

        # Sequence: Element list

        # Find Element Activity
        loggerBot.info("Start: Mapear botao PF 9Find_Element0")

        # Displayname: Mapear botao PF 9Find_Element0
        botaoPF = webBot.find_element(selector="pf", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear botao PF 9Find_Element0")

        # Find Element Activity
        loggerBot.info("Start: Mapear botão PJ (Find_Element)")

        # Displayname: Mapear botão PJ (Find_Element)
        botaoPJ = webBot.find_element(selector="pj", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear botão PJ (Find_Element)")

        # Find Element Activity
        loggerBot.info("Start: Mapear campo Nome e Razão Social (Find_Element)")

        # Displayname: Mapear campo Nome e Razão Social (Find_Element)
        campoNomeRazao = webBot.find_element(selector="nomeRazao", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear campo Nome e Razão Social (Find_Element)")

        # Find Element Activity
        loggerBot.info("Start: Mapear campo CPF e CNPJ (Find_Element)")

        # Displayname: Mapear campo CPF e CNPJ (Find_Element)
        campoCpfCnpj = webBot.find_element(selector="cpfCnpj", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear campo CPF e CNPJ (Find_Element)")

        # Find Element Activity
        loggerBot.info("Start: Mapear botão Enviar (Find_Element)")

        # Displayname: Mapear botão Enviar (Find_Element)
        botaoEnviar = webBot.find_element(selector="btnEnviar", by=By.ID, waiting_time=1000, ensure_visible=False, ensure_clickable=False)

        loggerBot.info("End: Mapear botão Enviar (Find_Element)")

        loggerBot.info("End: Element_Library")

        # ForEach Activity
        loggerBot.info("Start: ForEach")

        # Displayname: ForEach
        for linha in listaFornecedores:
            # Sequence: Body

            loggerBot.info("Start: Mapear elementos da planilha excel (Form_Element_Action_Web)")

            # DisplayName: Mapear elementos da planilha excel (Form_Element_Action_Web)

            # Sequence: Action list

            # Sequence: Conditional Structure

            # If Activity
            loggerBot.info("Start: If Condition")

            # Displayname: If Condition
            if linha[0] == "PF":
                # Sequence: Body

                # Click Activity
                loggerBot.info("Start: Click in botaoPF element")

                # Displayname: Click in botaoPF element
                botaoPF.click()

                loggerBot.info("End: Click in botaoPF element")


                loggerBot.info("End: If Condition")

            # Else Activity
            # Displayname: Else
            else:
                loggerBot.info("Start: Else")

                # Sequence: Body

                # Click Activity
                loggerBot.info("Start: Click in botaoPJ element")

                # Displayname: Click in botaoPJ element
                botaoPJ.click()

                loggerBot.info("End: Click in botaoPJ element")


                loggerBot.info("End: Else")

            # Type Into Activity
            loggerBot.info("Start: Type Into campoNomeRazao field")

            # Displayname: Type Into campoNomeRazao field
            campoNomeRazao.send_keys(linha[1])

            loggerBot.info("End: Type Into campoNomeRazao field")

            # Type Into Activity
            loggerBot.info("Start: Type Into campoCpfCnpj field")

            # Displayname: Type Into campoCpfCnpj field
            campoCpfCnpj.send_keys(linha[2])

            loggerBot.info("End: Type Into campoCpfCnpj field")

            # Click Activity
            loggerBot.info("Start: Click in botaoEnviar element")

            # Displayname: Click in botaoEnviar element
            botaoEnviar.click()

            loggerBot.info("End: Click in botaoEnviar element")

            loggerBot.info("End: Mapear elementos da planilha excel (Form_Element_Action_Web)")


        loggerBot.info("End: ForEach")

        # Wait Activity
        loggerBot.info("Start: Wait")

        # Displayname: Wait
        webBot.wait(3000)

        loggerBot.info("End: Wait")



        logging.shutdown()

        return
if __name__ == '__main__':
    bot = Bot()
    bot.bot()