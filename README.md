# 📋 Frequência Automática — Conexão Educação

Este projeto é uma ferramenta de automação desenvolvida em Python para facilitar o registro de frequência de alunos no site **Conexão Educação**. Ele lê os dados de um diário de classe em formato Excel e simula os comandos de teclado necessários para preencher as presenças e faltas automaticamente.

## 🚀 Funcionalidades

- **Leitura Inteligente de Excel:** Suporta arquivos `.xlsx` e `.xlsm`.
- **Filtro Automático de Status:** Ignora alunos com status "Remanejado", "Cancelado" ou "Transferido".
- **Ordenação Alfabética:** Organiza a lista de alunos automaticamente (ignorando acentos).
- **Interface Gráfica (GUI):** Interface intuitiva em modo escuro (Dark Mode) para configuração fácil.
- **Controle de Automação:** Botões para Iniciar, Pausar/Retomar e Parar o processo.
- **Preview de Dados:** Visualize os alunos e suas respectivas frequências antes de iniciar a automação.
- **Configuração de Delay:** Ajuste o tempo entre as ações para se adequar à velocidade de carregamento do site.

## 🛠️ Pré-requisitos

Para executar este script, você precisará ter o Python instalado em sua máquina, além das seguintes bibliotecas:

- `pyautogui`: Para controle do teclado.
- `openpyxl`: Para leitura de arquivos Excel.
- `tkinter`: Para a interface gráfica (geralmente já vem com o Python).

## 🔧 Instalação

1. Clone este repositório ou baixe o arquivo `script_melhorado.py`.
2. Instale as dependências necessárias via terminal:

```bash
pip install pyautogui openpyxl
```

## 📖 Como Usar

1. **Abra o Navegador:** Faça login no site do Conexão Educação e navegue até a tela de lançamento de frequência do aluno desejado.
2. **Execute o Script:**
   ```bash
   python script_melhorado.py
   ```
3. **Configure o Excel:**
   - Informe o nome da **Aba** (ex: "Abr" para Abril).
   - Defina as **Linhas** de início e fim onde estão os nomes dos alunos.
   - Informe a **Coluna** do nome e o intervalo de **Colunas** das presenças.
4. **Carregue o Arquivo:** Clique em "📂 Carregar Excel" e selecione o seu diário.
5. **Verifique o Preview:** Confira na área de texto se os nomes e as frequências (`.` para presente, `f` para falta) foram lidos corretamente.
6. **Inicie a Automação:**
   - Clique em "▶ Iniciar".
   - **Importante:** Você terá **4 segundos** para clicar na janela do navegador (no primeiro campo de frequência do primeiro aluno) antes que a automação comece.
7. **Acompanhe o Progresso:** A barra de progresso indicará qual aluno está sendo processado.

## ⚙️ Detalhes da Configuração

- **Lógica de Teclas:**
  - O site geralmente vem com todos os dias marcados como **PRESENTE**.
  - Se no Excel constar `f` (falta), o script pressiona `Space` (desmarca o checkbox) e `Tab`.
  - Se constar `.` (presente), o script pressiona apenas `Tab` (mantém o checkbox marcado).
  - Ao final dos dias de um aluno, o script aguarda o delay configurado para passar para o próximo.

- **Status Ignorados:** Alunos que possuam termos como "rem. de sala", "cancelado" ou "transf." no nome ou nas células de frequência não serão processados.

## ⚠️ Avisos Importantes

- **Não mexa no computador** enquanto a automação estiver rodando, pois ela assume o controle do teclado.
- Se algo der errado, use o botão **Parar** na interface ou mova o mouse rapidamente para um dos cantos da tela (fail-safe do PyAutoGUI) para interromper a execução.
- Certifique-se de que a ordem dos alunos no Excel é a mesma exibida no site para evitar erros de lançamento.

---
Desenvolvido para otimizar o trabalho administrativo diário. 🍎
