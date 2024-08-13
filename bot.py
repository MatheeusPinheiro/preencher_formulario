
# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

#Pandas
import pandas as pd

#Webdriver Manager
from webdriver_manager.chrome import ChromeDriverManager

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False



def preencher_dados(bot, nome, email, telefone, sexo, sobre):

    
    #Nome
    bot.find_element('//*[@id="166517069"]', By.XPATH).send_keys(str(nome))
    bot.wait(1000)

    #email
    bot.find_element('//*[@id="166517072"]', By.XPATH).send_keys(str(email))
    bot.wait(1000)

    #telefone
    bot.find_element('//*[@id="166517070"]', By.XPATH).send_keys(str(telefone))
    bot.wait(1000)

    #sexo
    if str(sexo.lower()) == 'masculino':
        bot.find_element('//*[@id="166517071_1215509812_label"]/span[2]', By.XPATH).click()
    else:
        bot.find_element('//*[@id="166517071_1215509813_label"]/span[1]', By.XPATH).click()
        
    #Sobre
    bot.find_element('//*[@id="166517073"]', By.XPATH).send_keys(str(sobre))
    bot.wait(1000)

    #Btn Enviar Dados
    bot.find_element('//*[@id="patas"]/main/article/section/form/div[2]/button', By.XPATH).click()

    


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = ChromeDriverManager().install()

   
    # Implement here your logic...
    try:
        
        dados_form = pd.read_excel(r'resources/DadosFormulario.xlsx')

        for i in range(len(dados_form)):
            nome = dados_form.iloc[i]['Nome']
            email = dados_form.iloc[i]['Email']
            telefone = dados_form.iloc[i]['Telefone']
            sexo = dados_form.iloc[i]['Sexo']
            sobre = dados_form.iloc[i]['Sobre']

            # Opens the BotCity website.
            bot.browse("https://pt.surveymonkey.com/r/WLXYDX2")

            #Maximize Window
            bot.maximize_window()


            preencher_dados(bot, nome, email,telefone, sexo, sobre)

            bot.wait(4000)

            bot.stop_browser()

    except Exception as er:
        print(f'Erro: {er}')
        bot.save_screenshot('Erro.png')

    finally:
        # Wait 3 seconds before closing
        bot.wait(5000)

        # Finish and clean up the Web Browser
        # You MUST invoke the stop_browser to avoid
        # leaving instances of the webdriver open
       

        # Uncomment to mark this task as finished on BotMaestro
        maestro.finish_task(
            task_id=execution.task_id,
            status=AutomationTaskFinishStatus.SUCCESS,
            message="Tarefa Finalizada."
        )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
