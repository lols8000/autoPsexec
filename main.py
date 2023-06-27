import subprocess
import ctypes
import sys


def funcoesPsexec(opcao):
    if opcao == '1':
        mensagem = input('Digite a mensagem: ')
        comando = f'msg * {mensagem}'

    elif opcao == '2':
        print('Escolha uma opção: ')
        print('1-Instalar Java')
        print('2-Instalar VNC')
        op = input('Escolha uma opção: ')

        nome_do_arquivo = input('Digite o nome do arquivo: ')

        if op == '1':
            comando = f'-c -f c:\\{nome_do_arquivo}.exe /s'
        if op == '2':
            comando = f'-c -f c:\\{nome_do_arquivo}.exe /verysilent'


    elif opcao == '3':
        nome_do_servico = input('Digite o nome do serviço: ')
        comando = f'sc stop {nome_do_servico}'

    elif opcao == '4':
        nome_do_servico = input('Digite o nome do serviço: ')
        comando = f'sc start {nome_do_servico}'

    elif opcao == '5':
        comando = 'gpupdate'

    else:
        print('comando inválido')

    return comando


def executaPsexec():
    print('Escolha uma opção: ')
    print('1-Enviar mensagem')
    print('2-Instalar um programa')
    print('3-Parar um serviço')
    print('4-Iniciar um serviço')
    print('5-Iniciar gpupdate')
    opcao = input('Digite uma opção: ')

    comandoEscolhido = funcoesPsexec(opcao)

    computador_destino = input("Digite o nome do computador de destino: ")

    # Comando psexec com o usuário, senha, nome do computador e mensagem fornecidos
    comando = f"psexec \\\\{computador_destino} {comandoEscolhido}"

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


def run_as_admin():
    script = sys.argv[0]
    params = ' '.join(sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, script, params, 1)


if __name__ == '__main__':
    if ctypes.windll.shell32.IsUserAnAdmin():
        executaPsexec()
    else:
        run_as_admin()
