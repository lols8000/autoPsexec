import subprocess
import ctypes
import socket
import sys
import os


def funcoesPsexec(opcao):

    computador_destino = input("Digite hostname\ip de destino: ")

    if opcao == '1':
        mensagem = input('Digite a mensagem: ')
        comando = f'\\\\{computador_destino} msg * {mensagem}'

    elif opcao == '2':
        print('Escolha uma opção: ')
        print('1-Instalar Java')
        print('2-Instalar VNC')
        op = input('Escolha uma opção: ')

        nome_do_arquivo = input('Digite o nome do arquivo: ')

        if op == '1':
            comando = f'\\\\{computador_destino} -c -f c:\\{nome_do_arquivo}.exe /s'
        if op == '2':
            comando = f'\\\\{computador_destino} -c -f c:\\{nome_do_arquivo}.exe /verysilent'

    elif opcao == '3':
        nome_do_servico = input('Digite o nome do serviço: ')
        comando = f'\\\\{computador_destino} sc stop {nome_do_servico}'

    elif opcao == '4':
        nome_do_servico = input('Digite o nome do serviço: ')
        comando = f'\\\\{computador_destino} sc start {nome_do_servico}'

    elif opcao == '5':
        comando = f'\\\\{computador_destino} gpupdate'

    elif opcao == '6':
        nome_processo = input('Digite o nome do processo: ')
        comando = f'\\\\{computador_destino} taskkill /F /IM {nome_processo}.exe'

    elif opcao == '7':
        hostname = socket.gethostname()
        comando = f'\\\\{hostname} cmd /c "taskkill /F /IM winvnc.exe & xcopy "\\\\{hostname}\\c$\\Program Files\\uvnc bvba" "\\\\{computador_destino}\\c$\\Program Files\\uvnc bvba" /E /I /Y"'

    elif opcao == '8':
        comando = f'\\\\{computador_destino} shutdown /r /t 0'

    else:
        print('comando inválido')

    return comando


def executaPsexec():
    print('Escolha uma opção: ')
    print('1-Enviar mensagem.')
    print('2-Instalar um programa.')
    print('3-Parar um serviço.')
    print('4-Iniciar um serviço.')
    print('5-Iniciar gpupdate.')
    print('6-Finalizar um processo.')
    print('7-Copiar pasta do VNC para o computador do usuário.')
    print('8-Reiniciar computador.')

    opcao = input('\nDigite uma opção: ')

    _ = os.system('cls')

    comandoEscolhido = funcoesPsexec(opcao)

    # Comando psexec com o usuário, senha, nome do computador e mensagem fornecidos
    comando = f"psexec {comandoEscolhido}"

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

def run_as_admin():
    script = sys.argv[0]
    params = ' '.join(sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, script, params, 1)


if __name__ == '__main__':
    if ctypes.windll.shell32.IsUserAnAdmin():
        executaPsexec()
    else:
        run_as_admin()
