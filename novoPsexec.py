import subprocess


def runIpconfigRemotely(remote_machine):
    # Caminho para o PsExec.exe (ajuste conforme necessário)
    psexec_path = r'C:\Windows\System32\PsExec.exe'

    # Comando a ser executado remotamente (ipconfig neste caso)
    command = 'wmic bios get serialnumber'

    # Monta o comando completo para o PsExec
    full_command = f'{psexec_path} \\\\{remote_machine} {command}'

    try:
        # Executa o comando usando o subprocess
        subprocess.run(full_command, check=True, shell=True)
        print(f"Comando '{command}' executado com sucesso em {remote_machine}.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao executar o comando '{command}' em {remote_machine}. Código de saída: {e.returncode}")


if __name__ == "__main__":
    remote_machine = input("Digite o nome ou IP da máquina remota: ")
    #username = input("Digite o nome de usuário: ")
    #password = input("Digite a senha: ")

    runIpconfigRemotely(remote_machine)
