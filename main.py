import subprocess
import getpass
import ctypes
import socket
import sys
import os
import time


def usuarioLogado(computador_destino):
    psexec_path = r'C:\Windows\System32\PsExec.exe'

    # Comando psexec com o usuário, senha, nome do computador e mensagem fornecidos
    comando = f"{psexec_path} \\\\{computador_destino} quser"
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    saida, erro = processo.communicate()
    usuario = str(saida.decode('latin1'))
    lines = usuario.split('\n')

    if len(lines) >= 3:
        columns = lines[6].split()
        if len(columns) >= 2:
            return columns[0]


def interfaceWiFiName(computador_destino):
    psexec_path = r'C:\Windows\System32\PsExec.exe'

    # Comando psexec para descobrir o nome da interface Wi-Fi
    comando = f"{psexec_path} \\\\{computador_destino} netsh interface show interface"
    processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    saida, erro = processo.communicate()
    interface_wifi = str(saida.decode('latin1'))
    lines = interface_wifi.split('\n')

    if len(lines) >= 3:
        columns = lines[8].split()
        if len(columns) >= 2:
            return columns[3]


def sendMessage(computador_destino):
    mensagem = input('Digite a mensagem: ')
    comando = f'\\\\{computador_destino} msg * {mensagem}'

    return comando


def installProgram(computador_destino):
    print('Escolha uma opção: ')
    print('1-Instalar Java.')
    print('2-Instalar VNC.')
    print('3-Instalar Chrome.')
    print('4-Instalar Firefox.')
    print('5-Listar programas instalados.')

    op = input('Escolha uma opção: ')

    usuario_logado_destino = usuarioLogado(computador_destino)
    winget_path = f'C:\\Users\\{usuario_logado_destino}\\AppData\\Local\\Microsoft\\WindowsApps\\winget.exe'

    install_option = {
        '1': f'\\\\{computador_destino} {winget_path} install Oracle.JavaRuntimeEnvironment --silent',
        '2': f'\\\\{computador_destino} {winget_path} install uvncbvba.UltraVnc --silent',
        '3': f'\\\\{computador_destino} {winget_path} install Google.Chrome --silent',
        '4': f'\\\\{computador_destino} {winget_path} install Mozilla.Firefox --silent',
        '5': f'\\\\{computador_destino} {winget_path} list'
    }

    return install_option.get(op, 'Opção inválida.')


def stopService(computador_destino):
    nome_do_servico = input('Digite o nome do serviço: ')
    comando = f'\\\\{computador_destino} sc stop {nome_do_servico}'

    return comando


def startService(computador_destino):
    nome_do_servico = input('Digite o nome do serviço: ')
    comando = f'\\\\{computador_destino} sc start {nome_do_servico}'

    return comando


def gpupdate(computador_destino):
    comando = f'\\\\{computador_destino} gpupdate'

    return comando


def terminateProcess(computador_destino):
    nome_processo = input('Digite o nome do processo: ')
    comando = f'\\\\{computador_destino} taskkill /F /IM {nome_processo}.exe'

    return comando


def copyVncForUserComputer(computador_destino):
    hostname = socket.gethostname()

    finalizaProcesso(computador_destino, 'winvnc')
    comando = f'\\\\{hostname} cmd /c "xcopy "\\\\{hostname}\\c$\\Program Files\\uvnc bvba" "\\\\{computador_destino}\\c$\\Program Files\\uvnc bvba" /E /I /Y"'

    return comando


def restarComputer(computador_destino):
    comando = f'\\\\{computador_destino} shutdown /r /t 0'

    return comando


def ipconfig(computador_destino):
    comando = f'\\\\{computador_destino} ipconfig /all'

    return comando


def computerName(computador_destino):
    comando = f'\\\\{computador_destino} hostname'

    return comando


def copyTeamsForUserComputer(computador_destino):
    hostname = socket.gethostname()
    usuario_logado_origem = getpass.getuser()
    usuario_logado_destino = usuarioLogado(computador_destino)

    finalizaProcesso(computador_destino, 'Teams')
    comando = f'\\\\{hostname} cmd /c "xcopy "\\\\{hostname}\\c$\\Users\\{usuario_logado_origem}\\AppData\\Local\\Microsoft\\Teams" "\\\\{computador_destino}\\c$\\Users\\{usuario_logado_destino}\\AppData\\Local\\Microsoft\\Teams" /E /I /Y"'

    return comando


def disableInterfaceWiFi(computador_destino):
    wifi_name = interfaceWiFiName(computador_destino)

    if wifi_name != 'Wi-Fi':
        wifi_name = 'Wi-Fi'

    comando = f'\\\\{computador_destino} wmic path win32_networkadapter where NetConnectionID="{wifi_name}" call disable'

    return comando


def copy_agent_glpi(computador_destino):
    origem = r"\\10.42.0.38\humap\glpiagent\glpiagentinstall.vbs"
    destino = f"\\\\{computador_destino}\\c$"

    comando = f'xcopy "{origem}" "{destino}" /Y'

    try:
        processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        processo.communicate()

        # Verifica se a cópia foi bem-sucedida
        if processo.returncode == 0:
            print('Cópia concluída com sucesso')
        else:
            print('Erro ao copiar o arquivo')
            print(f'Saída de erro: {processo.stderr.read().decode("latin1")}')
    except Exception as e:
        print(f'Erro ao executar o comando: {e}')


def GlpiInfo(computador_destino):
    copy_agent_glpi(computador_destino)

    time.sleep(3)

    comando = f'\\\\{computador_destino} cscript "C:\\glpiagentinstall.vbs" && del "\\\\{computador_destino}\\c$\\glpiagentinstall.vbs"'

    return comando


def getSerialNumber(computador_destino):
    comando = f'\\\\{computador_destino} wmic bios get serialnumber'

    return comando


def functionPsexec(option):
    computador_destino = input("Digite hostname\ip de destino: ")

    option_menu = {
        '1': sendMessage,
        '2': installProgram,
        '3': stopService,
        '4': startService,
        '5': gpupdate,
        '6': terminateProcess,
        '7': copyVncForUserComputer,
        '8': restarComputer,
        '9': ipconfig,
        '10': computerName,
        '11': copyTeamsForUserComputer,
        '12': disableInterfaceWiFi,
        '13': GlpiInfo,
        '14': getSerialNumber
    }

    return option_menu.get(option, lambda: 'Opção inválida.')(computador_destino)


def finalizaProcesso(computador_destino, name):
    psexec_path = r'C:\Windows\System32\PsExec.exe'

    # Comando psexec com o usuário, senha, nome do computador e mensagem fornecidos
    comando = f"{psexec_path} \\\\{computador_destino} taskkill /F /IM {name}.exe"

    # Executa o comando externo
    subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)


def executaPsexec():
    while True:
        print('Escolha uma opção: ')
        print('0-finalizar execução do programa.')
        print('1-Enviar mensagem.')
        print('2-Instalar um programa da lista.')
        print('3-Parar um serviço.')
        print('4-Iniciar um serviço.')
        print('5-Iniciar gpupdate.')
        print('6-Finalizar um processo.')
        print('7-Copiar pasta do VNC para o computador do usuário.')
        print('8-Reiniciar computador.')
        print('9-ipconfig.')
        print('10-Informa o nome da máquina.')
        print('11-Copiar pasta do Teams para o computador do usuário.')
        print('12-Desabilitar rede wi-fi.')
        print('13-Atualizar Glpi.')
        print('14-Serial Number do computador.')

        opcao = input('\nDigite uma opção: ')

        if opcao == '0':
            raise SystemExit

        else:
            _ = os.system('cls')

            psexec_path = 'C:\\Windows\\System32\\PsExec.exe'
            comando_escolhido = functionPsexec(opcao)

            # Comando psexec com o usuário, senha, nome do computador e mensagem fornecidos
            comando = f"{psexec_path} {comando_escolhido}"

            # Executa o comando externo
            processo = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            saida, erro = processo.communicate()

            # Verifica a saída e o erro
            if processo.returncode == 0:
                print("Comando executado com sucesso.")
                print("Saída:", saida.decode('latin1'))
            else:
                print("Erro ao executar o comando.")
                print("Erro:", erro.decode('latin1'))

            input('Pressione qualquer tecla para continuar...')

            _ = os.system('cls')


def runAsAdmin():
    script = sys.argv[0]
    params = ' '.join(sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, script, params, 1)


if __name__ == '__main__':
    if ctypes.windll.shell32.IsUserAnAdmin():
        executaPsexec()
    else:
        runAsAdmin()
