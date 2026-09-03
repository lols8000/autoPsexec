# Desenvolvimento — Central N2 Workstation v3

## 1. Objetivo

Este guia define práticas para evolução da Central N2 sem degradar arquitetura, segurança ou experiência operacional.

---

## 2. Ambiente de desenvolvimento

Requisitos:

- Windows;
- Python 3.10+;
- Git;
- pytest para desenvolvimento;
- opcionalmente PsExec/Sysinternals para testes manuais controlados.

Instalação de dependência de teste:

```powershell
cd central_n2
python -m pip install pytest
```

O runtime de produção utiliza biblioteca padrão do Python.

---

## 3. Estrutura

```text
central_n2/
├── main.py
├── config/
├── core/
│   ├── executor.py
│   ├── jobs.py
│   ├── logger.py
│   └── result.py
├── modules/
├── ui/
│   ├── console.py
│   └── console_v3.py
├── tests/
└── docs/
```

---

## 4. Regra de dependência

Dependência deve fluir:

```text
UI
 ↓
Módulos
 ↓
Core / Executor
```

Evite:

```text
Core importando UI
Módulo chamando input()
Executor contendo regra de impressora/GLPI/etc.
```

---

## 5. Como criar um novo módulo

Exemplo:

```python
from core.executor import RemoteExecutor
from core.result import CommandResult


class ExampleModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def status(self, host: str) -> CommandResult:
        script = r'''
[pscustomobject]@{
    Example = $true
}
'''
        return self.executor.execute_powershell_json(host, script)
```

### Regras

- receber `RemoteExecutor` por injeção;
- não instanciar executor internamente;
- retornar `CommandResult`;
- preferir JSON estruturado;
- definir timeout quando operação puder ser longa;
- não fazer `input()`/`print()` no módulo;
- escapar entradas interpoladas em PowerShell;
- documentar impacto.

---

## 6. Interface responsiva

Operações potencialmente bloqueantes devem passar pelo método `execute()` da UI v3, que utiliza `ResponsiveJobRunner`.

Padrão:

```python
self.execute(
    "Descrição amigável",
    lambda: self.module.operation(self.host),
    timeout=180,
)
```

Evite:

```python
result = self.module.operation(self.host)
```

quando a operação pode levar tempo perceptível, pois isso congela feedback visual.

---

## 7. Threads

O runner possui pool limitado a quatro workers.

### Quando usar

- WinRM;
- PsExec;
- chamadas de rede;
- PowerShell remoto;
- leitura longa de eventos;
- inventários;
- DISM/SFC.

### Quando não usar automaticamente

- cálculo Python trivial;
- formatação;
- validação local simples;
- múltiplas remediações concorrentes na mesma estação;
- operações destrutivas em massa.

### Regra

Concorrência deve resolver um gargalo mensurável, não ser adicionada por estética técnica.

---

## 8. Timeouts

Toda operação remota deve possuir timeout previsível.

Categorias sugeridas:

```text
consulta simples      30–90 s
inventário pesado     120–300 s
instalação/software   300–900 s
SFC/DISM              até 3600 s
```

Esses valores são referência, não contrato fixo.

### Timeout remoto vs local

O timeout da UI controla espera local. O executor/subprocess também precisa de timeout adequado para impedir worker permanentemente preso.

Quando implementar nova operação longa, verifique os dois níveis.

---

## 9. `CommandResult`

Use o contrato central para manter UI e logging consistentes.

Não retorne tuplas ad-hoc como:

```python
(True, "ok")
```

Prefira:

```python
CommandResult(...)
```

ou métodos do executor que já produzem o resultado.

### `data`

Use `data` para objetos estruturados.

### `stdout`

Use quando a ferramenta nativa só retorna texto ou quando o texto é relevante integralmente.

### `stderr`

Preserve erro suficiente para diagnóstico, evitando segredos.

---

## 10. PowerShell remoto

### Preferir

```powershell
Get-CimInstance
Get-Service
Get-Printer
Get-NetAdapter
Get-PhysicalDisk
```

quando fornecem objetos estruturados.

### Evitar parsing frágil

Não basear lógica crítica em posição fixa de texto quando existe cmdlet estruturado.

### Compatibilidade

Antes de adicionar cmdlet:

- considerar Windows 10/11;
- verificar disponibilidade do módulo;
- prever `try/catch` para recursos opcionais;
- tratar ausência como `null` quando apropriado.

---

## 11. Sanitização de entrada

Entrada usada em PowerShell deve ser escapada.

Padrão mínimo para string single-quoted:

```python
safe = value.replace("'", "''")
```

Melhor ainda: limitar formato quando possível.

Exemplo de porta:

```python
port = int(value)
```

Não crie execução arbitrária de shell baseada diretamente em texto do operador.

---

## 12. Operações destrutivas

Antes de adicionar ação de escrita, responder:

1. ela é necessária para N2?
2. há uma consulta prévia que confirme a causa?
3. existe rollback?
4. usuário pode perder dados?
5. conectividade pode cair?
6. precisa de reboot?
7. deve exigir `SIM`?
8. pode ser executada em lote?
9. precisa de log especial?

Se a resposta for incerta, comece em modo consulta/planejamento.

---

## 13. Testes

Executar sempre:

```powershell
python -m compileall .
python -m pytest -q
```

### Testes mínimos para novo módulo

- validação local pura quando existir;
- geração/escaping de comando;
- comportamento em entrada inválida;
- timeout quando aplicável;
- smoke test de métodos utilizados pela UI.

### Teste remoto

Não deve ser obrigatório no CI público.

Integrações reais devem ocorrer em VM/estação de bancada autorizada.

---

## 14. CI

O GitHub Actions compila o Python e executa pytest em Windows.

Um PR não deve ser promovido se:

- `compileall` falhar;
- pytest falhar;
- documentação estiver divergente;
- houver segredo/configuração interna no diff.

---

## 15. Checklist de Pull Request

```text
[ ] branch criada a partir do master atual
[ ] mudança tem responsabilidade clara
[ ] sem shell=True desnecessário
[ ] entradas sanitizadas
[ ] timeout definido
[ ] operação longa usa runner responsivo
[ ] confirmação em ações de impacto
[ ] CommandResult preservado
[ ] testes adicionados/atualizados
[ ] compileall passa
[ ] pytest passa
[ ] docs atualizadas
[ ] nenhum segredo no diff
[ ] comportamento testado em bancada quando necessário
```

---

## 16. Padrão de commits

Sugestão:

```text
feat: adiciona diagnóstico X
fix: corrige timeout do runner
docs: atualiza manual operacional
test: cobre módulo de impressoras
refactor: separa lógica de rede
```

Commits pequenos e semanticamente claros facilitam rollback e revisão.

---

## 17. Logging

Novas ações administrativas devem ser compatíveis com auditoria.

Evite imprimir segredo no stdout e depois enviá-lo ao logger.

---

## 18. Performance

Antes de otimizar:

1. identificar onde está a espera;
2. distinguir CPU-bound de I/O-bound;
3. medir número de round-trips remotos;
4. agregar consultas em uma execução quando fizer sentido;
5. limitar concorrência;
6. evitar coleta excessiva que sobrecarregue estação do usuário.

### Exemplo

Melhor:

```text
1 chamada PowerShell → CPU + RAM + disco + rede
```

que:

```text
20 chamadas WinRM separadas
```

quando todos os dados podem ser obtidos de forma segura em uma única sessão/comando.

---

## 19. Compatibilidade com UI futura

Módulos não devem depender do console atual.

Isso permite futuramente criar:

```text
Tkinter
Web UI
API local
TUI
```

reutilizando o mesmo executor e módulos.

---

## 20. Dívida técnica conhecida / evolução

Itens recomendados:

- `settings.local.json` com merge seguro;
- sessões WinRM reutilizáveis quando vantajosas;
- job IDs e cancelamento cooperativo;
- correlação automática de sintomas;
- relatório N2 estruturado;
- testes de integração em VM Windows;
- separação de permissões read-only/remediation;
- remoção futura da UI v2 quando não houver mais necessidade histórica.

---

## 21. Definition of Done

Uma funcionalidade não está concluída apenas porque “executa o comando”.

Ela está concluída quando:

```text
funciona
+ falha de forma previsível
+ possui timeout
+ mantém feedback visual
+ retorna evidência
+ respeita segurança
+ possui testes
+ está documentada
```
