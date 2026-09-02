# Execução Automática de BATs

Gerenciador leve para execução de scripts `.BAT` no Windows em intervalos de tempo definidos, utilizando apenas **uma única tarefa** no Agendador de Tarefas do Windows.

Os scripts são executados de forma sequencial (em fila), evitando que múltiplas janelas de terminal sejam abertas simultaneamente e consumam recursos da máquina.

---

## Estrutura do Projeto

```text
vbs/
├── bat/              # Pasta para armazenar os arquivos BAT
│   └── teste.bat
├── logs/             # Histórico de execuções (gerado automaticamente)
│   ├── execucao_2000-01-01_00-00-00.txt
│   └── execucao_2000-01-01_00-00-01_ERRO.txt
├── .gitignore
├── fila.vbs          # Script principal gerenciador da fila
├── config.ini        # Arquivo de configuração dos intervalos
└── estado.ini        # Controle interno de execuções (gerado automaticamente)
```

---

## Arquivos de Configuração

### 1. `config.ini`

Define quais arquivos `.BAT` devem ser executados e o intervalo mínimo (em minutos) entre cada execução.

```
teste1.bat=5
teste2.bat=10
teste3.bat=15
```

* **Lado esquerdo (`=`):** Nome exato do arquivo `.BAT` localizado na pasta `bat/`.
* **Lado direito (`=`):** Intervalo de tempo em **minutos**.

### 2. `estado.ini`

Criado automaticamente na primeira execução do `fila.vbs`. Grava a data e hora do último processamento de cada script para calcular o próximo ciclo. **Não alterar este arquivo manualmente.**

---

## Pastas do Sistema

* **`bat/`**: Guarde aqui todos os scripts **`.bat`** que precisar rodar. Apenas os arquivos cadastrados no `config.ini` serão processados.
* **`logs/`**: Registra o histórico detalhado de todas as execuções (`execucao_YYYY-MM-DD.log`). Caso ocorra uma falha no dia, o arquivo receberá o sufixo `_ERRO.log`.

**Informações gravadas nos logs:**
* Data e hora da execução
* Nome do script executado
* Status da execução
* Código de retorno (exit code)
* Duração da execução



---

## Fluxo de Funcionamento

O Agendador do Windows executa o `fila.vbs` por exemplo: **a cada 1 minuto**. Ao ser acionado, o script lê o `config.ini`, verifica o histórico do `estado.ini` e executa sequencialmente apenas as tarefas que atingiram seu tempo limite.

| Horário | Tarefas Processadas |
| --- | --- |
| **08:00** | `teste1.bat`, `teste2.bat`, `teste3.bat` |
| **08:05** | `teste1.bat` |
| **08:10** | `teste1.bat`, `teste2.bat` |
| **08:15** | `teste1.bat`, `teste3.bat` |

---

## Configuração no Agendador de Tarefas do Windows

É necessário configurar e criar apenas **uma** tarefa no Windows para gerenciar todos os scripts.

1. Pressione `Win + R`, digite **`taskschd.msc`** e pressione **Enter** ou acesse o **Agendador de Tarefas** no menu windows.
2. No menu lateral, clique em **Criar Tarefa**.
3. Configure as abas conforme abaixo:

### Aba General (Geral)

* **Nome:** `Executor de BATs` ou qualquer outro nome a sua escolha.
* Marque: **Executar estando o usuário conectado ou não**
* Marque: **Executar com privilégios mais altos** *(necessário se os scripts exigirem permissão de administrador)*

### Aba Triggers (Disparadores)

* Clique em **Novo...** e configure para iniciar **Diariamente**.
* Em *Configurações avançadas*, marque **Repetir a tarefa a cada:** `1 minuto`.
* Na opção *por um período de:*, selecione **Indefinidamente**.

### Aba Actions (Ações)

* Clique em **Nova...** e selecione a ação **Iniciar um programa**.
* **Programa/script:** `wscript.exe`
* **Adicionar argumentos:** `"C:\CAMINHO\vbs\fila.vbs"` *(substitua pelo caminho real)*
* **Iniciar em:** `C:\CAMINHO\vbs` *(substitua pelo caminho real sem aspas)*

### Aba Settings (Configurações)

* Se a tarefa já estiver em execução, selecione: **Não iniciar uma nova instância**.

---

## Testando a Aplicação

1. Antes de ativar o agendamento, abra a pasta e dê um duplo clique no arquivo `fila.vbs` para realizar um teste manual.
2. Acesse o diretório `logs/` para conferir se os arquivos de log foram gerados corretamente.
3. Após criar a tarefa no Agendador de Tarefas do Windows, clique com o botão direito sobre ela e selecione **Executar**.
