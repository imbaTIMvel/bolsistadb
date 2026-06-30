# BolsistaDB

Repositório oficial do programa BolsistaDB (*Bolsista D*ata*B*ase ou *D*ados *B*ancários), para associação de dados entre planilhas para tratamento de dados bancários dos alunos/bolsistas.

![Logo do BolsistaDB](assets/icons/bolsistadb.ico)

## 1. Requisitos

Para uso adequado do programa, o usuário deve possuir:

- Sistema Operacional: Windows 10 ou 11

Para uso da funcionalidade `Abrir planilha quando estiver pronta`, é recomendado que o usuário possua o *Microsoft Excel* instalado na versão mais recente.

## 2. Guia de Uso

### 2.1 Baixando e Instalando o Programa

Para usar o `BolsistaDB`, primeiro, você deve baixar o arquivo `.exe` disponível [aqui](https://github.com/imbaTIMvel/bolsistadb/releases). Procure pela versão mais recente (*Latest*) e clique no arquivo `.exe` para fazer o download.

> [!Warning]
> Caso você ainda tenha o executável de uma versão antiga do programa, recomenda-se excluí-lo.

Baixado o programa, você pode colocar o arquivo `.exe` onde achar melhor.

### 2.2 Abrindo o Programa

Feito isso, clique no arquivo `.exe` para abrir o programa.

![Abrindo o .exe](assets/tutorial/exe_in_downloads.png)

> [!Warning]
> É possível que o *Windows Defender* acuse o programa como "software perigoso". Neste caso, para executá-lo, você deve clicar em `Mais Informações` e, depois, no botão `Executar assim mesmo`.

![Windows Defender acusando o programa](assets/tutorial/windows_defender_01.png)

![Executar assim mesmo](assets/tutorial/windows_defender_02.png)

### 2.3 Interface do Programa

![Interface do programa - superior](assets/tutorial/ui_01.png)

#### 2.3.1 Campos de Arquivos

O programa possui dois campos para inserção de arquivos (planilhas Excel) de entrada. São eles:

| Campo               | Extensões de arquivo aceitas | Padronização do arquivo                                                 | Aceita mais de um arquivo? |
| ------------------- | ---------------------------- | ----------------------------------------------------------------------- | -------------------------- |
| Lista de Bolsistas  | .xlsx                        | [lista_bolsistas](assets/standard_sheets/lista_bolsistas.xlsx)          | Não                        |
| Formulário Bancário | .xlsx                        | [formulario_bancario](assets/standard_sheets/formulario_bancario.xlsx)  | Não                        |

Para cada um dos campos, há dois botões: `Selecionar` e `Remover`. Ao clicar em `Selecionar`, o programa abre um diálogo do *Explorador de Arquivos*, permitindo que o usuário selecione o arquivo Excel correspondente ao campo.

![Seleção de arquivo](assets/tutorial/file_select_01.png)

![Diálogo de seleção de arquivo](assets/tutorial/file_select_02.png)

Após selecionar o arquivo, o campo de arquivo inserido é atualizado.

![Interface com arquivos selecionados](assets/tutorial/file_select_03.png)

Ao clicar em `Remover`, se houver arquivo(s) selecionado(s) no campo correspondente, o programa remove o arquivo selecionado, deixando o campo vazio.

![Remoção de arquivo](assets/tutorial/file_removal_01.png)

![Interface sem arquivos selecionados](assets/tutorial/file_removal_02.png)

#### 2.3.2 Demais Recursos

![Interface do programa - inferior](assets/tutorial/ui_02.png)

Além disso, o programa possui:
- `Abrir planilha quando estiver pronta`: Um **toggle switch** que permite que o usuário defina se a planilha de saída será aberta ou não após a execução do programa;
- `Gerar Planilha`: O botão que inicia a execução do programa.

### 2.4 Execução

Ao inserir uma Lista de Bolsistas e um Formulário de Dados Bancários, ao clicar em `Gerar Planilha`, o programa deve gerar uma planilha no formato [template](assets/standard_sheets/template.xlsx), com as colunas:

- `No.`: Contagem do item na planilha gerada (em relação ao total);
- `ID`: Número associado ao aluno/bolsista dentro de sua respectiva turma;
- `TURMA`: Turma na qual o aluno/bolsista está inscrito, conforme especificado na Lista de Bolsistas (as turmas são organizadas por ordem alfabética na planilha de saída);
- `NOME COMPLETO`: Nome completo do aluno/bolsista, conforme especificado na Lista de Bolsistas (os alunos são organizados por ordem alfabética dentro de suas turmas);
- `Data Início`: Data em que o aluno/bolsista iniciou no curso/projeto - deve ser inserida manualmente pelo usuário;
- `Data Final`: Data em que o aluno/bolsista saiu do/encerrou o curso - deve ser inserida manualmente pelo usuário;
- `E-MAIL`: E-mail do aluno/bolsista, conforme especificado na Lista de Bolsistas;
- `CPF`: CPF do aluno/bolsista, conforme especificado na Lista de Bolsistas;
- `RG`: RG do aluno/bolsista, conforme especificado na Lista de Bolsistas;
- `DATA NASCIMENTO`: Data de nascimento do aluno/bolsista, conforme especificado na Lista de Bolsistas;
- `ENDEREÇO COMPLETO (Logradouro-Bairro-Cidade)`: Informações de residência do aluno/bolsista, computadas a partir das informações de "ENDEREÇO", "BAIRRO", "CIDADE" e "ESTADO" na Lista de Bolsistas;
- `CEP`: CEP do aluno/bolsista, conforme especificado na Lista de Bolsistas;
- `Nome Responsável`: Nome do responsável pelo aluno/bolsista, caso este seja menor de idade, conforme especificado no Formulário de Dados Bancários;
- `CPF - Respon.`: CPF do responsável pelo aluno/bolsista, caso este seja menor de idade, conforme especificado no Formulário de Dados Bancários;
- `E-mail Respon.`: E-mail do responsável pelo aluno/bolsista, caso este seja menor de idade, conforme especificado no Formulário de Dados Bancários;
- `Contato Respon.`: Contato (número de celular) do responsável pelo aluno/bolsista, caso este seja menor de idade, conforme especificado no Formulário de Dados Bancários;
- `Nome do Banco`: Código e nome do banco associado ao aluno/bolsista, seja conta própria ou pertencente ao responsável, conforme especificado no Formulário de Dados Bancários;
- `Agência`: Agência bancária associada ao aluno/bolsista, seja conta própria ou pertencente ao responsável, conforme especificado no Formulário de Dados Bancários;
- `Díg.Ag`: Dígito da agência associada ao aluno/bolsista, seja conta própria ou pertencente ao responsável, conforme especificado no Formulário de Dados Bancários;
- `Conta`: Código da conta corrente associada ao aluno/bolsista, seja própria ou pertencente ao responsável, conforme especificado no Formulário de Dados Bancários;
- `Díg.C/C`: Dígito da conta corrente associada ao aluno/bolsista, seja própria ou pertencente ao responsável, conforme especificado no Formulário de Dados Bancários;

Caso o aluno seja menor de idade (o algoritmo checa a partir da data inserida em `DATA NASCIMENTO`), o programa acusará caso as colunas `Nome Responsável`, `CPF - Respon.`, `E-mail Respon.` e `Contato Respon.` estejam vazias. Para associar adequadamente as informações entre a Lista de Bolsistas e o Formulário de Dados Bancários, o programa procura por correspondências de **nome completo**, **CPF** e **data de nascimento**. Em caso de múltiplas correspondências ou nenhuma correspondência, o programa deve reportar os erros em um arquivo `.txt` salvo na mesma pasta (a ser escolhida) da planilha de saída.

> [!Note]
> O programa possui um algoritmo próprio de tratamento de dados bancários, os adequando a convenções específicas de cada banco (segundo o banco de dados da FEBRABAN), conforme definido no arquivo [convencoes_bancos](assets/database/convencoes_bancos.pdf). O algoritmo prioriza a identificação do banco pelo **nome** (e não pelo código numérico) e aplica, para cada banco, o padrão de dígitos correspondente à agência e à conta corrente — incluindo o ajuste automático de zeros à esquerda e a remoção de sinais (hífens e pontos) inseridos pelo respondente.
>
> Em caso de dados inconsistentes — como ausência de convenção definida para um banco, ou dados não-computáveis inseridos pelo respondente — o algoritmo reporta as exceções em duas colunas adicionais geradas na planilha de saída: `Status` (nível de validação aplicado) e `Exceções do algoritmo` (descrição do problema identificado). Adicionalmente, as informações originais (inseridas pelos respondentes, antes do tratamento de dados) são preservadas na planilha de saída, nas colunas adicionais `Banco (original)`, `Agência (original)`, `Díg.Ag (original)`, `Conta (original)` e `Díg.C/C (original)`, servindo de fallback para o operador em casos de conferência manual.

> [!Note]
> Com os arquivos de entrada inseridos, ao clicar no botão `Gerar Planilha`, o programa deve fazer as correspondências entre a Lista de Bolsistas e o Formulário de Dados Bancários, gerando a planilha de saída (no formato especificado) e permitindo que o usuário escolha o local de salvamento do arquivo após o processamento.

![Gerando a planilha](assets/tutorial/create_sheet_01.png)

![Caixa de aviso](assets/tutorial/create_sheet_02.png)

![Salvando o arquivo](assets/tutorial/create_sheet_03.png)

> [!Warning]
> Todos os erros de processamento, não correspondência e/ou inconsistência de dados são reportados no arquivo `error_logs.txt`, que é salvo na mesma pasta da planilha de saída.

## 3. Releases

### `v0.1.0` BolsistaDB (*beta release*)

> [!Warning]
> O lançamento beta (*beta release*) foi desenvolvido para testes internos, visando identificar e corrigir bugs antes do lançamento de uma versão estável.

Data de lançamento: `30/06/2026`

Para fazer o download desta versão, clique [aqui](https://github.com/imbaTIMvel/bolsistadb/releases/download/v0.1.0/BolsistaDB.exe).

*Release* inicial do programa de associação de dados entre planilhas para tratamento de dados bancários dos alunos/bolsistas.

**Features:**
- Compatível com planilhas Excel, dos tipos:
  - `Lista de Bolsistas`: Na padronização [lista_bolsistas](assets/standard_sheets/lista_bolsistas.xlsx), no formato .xlsx;
  - `Formulário Bancário`: Na padronização [formulario_bancario](assets/standard_sheets/formulario_bancario.xlsx), no formato .xlsx;
- Associa dados de identificação (nome completo, CPF, data de nascimento, e-mail, RG, endereço e CEP) entre os dois arquivos de entrada, fazendo a correspondência por **nome completo**, **CPF** e **data de nascimento**;
- Identifica automaticamente, a partir da `DATA NASCIMENTO`, se o aluno/bolsista é menor de idade — exigindo, nesse caso, os dados do responsável legal (`Nome Responsável`, `CPF - Respon.`, `E-mail Respon.` e `Contato Respon.`);
- Possui um algoritmo próprio de tratamento de dados bancários, que:
  - Identifica o banco prioritariamente pelo **nome** (e não pelo código numérico), reduzindo erros de digitação no formulário;
  - Aplica, a cada banco, a convenção de dígitos específica (segundo o banco de dados [convencoes_bancos](assets/database/convencoes_bancos.pdf), baseado na FEBRABAN) para agência e conta corrente, incluindo ajuste de zeros à esquerda;
  - Remove sinais (hífens e pontos) informados pelo respondente sem comprometer a integridade do número da conta;
  - Reporta o nível de validação aplicado (`Status`) e quaisquer inconsistências (`Exceções do algoritmo`) em colunas adicionais da planilha de saída;
  - Preserva os dados originais (antes do tratamento) em colunas adicionais, como fallback para conferência manual do operador;
- Organiza a planilha de saída por turma e, dentro de cada turma, por ordem alfabética de aluno/bolsista;
- Reporta erros de processamento, ausência de correspondência e inconsistências de dados em um arquivo `error_logs.txt`, salvo na mesma pasta da planilha de saída;
- Permite que o usuário escolha o diretório de salvamento para a planilha (.xlsx) de saída.

Clique [aqui](https://github.com/imbaTIMvel/bolsistadb/releases) para acessar o **changelog completo**.

## 4. Desenvolvimento

**Autor:**

Timóteo Altoé (*handle*: [imbaTIMvel](github.com/imbaTIMvel))

**Datas:**

`17/06/2026` Início do projeto

`26/06/2026` Lançamento da versão *alfa* - para testes internos

`30/06/2026` Publicação da primeira versão oficial no GitHub

`30/06/2026` Lançamento da versão *beta* - para testes
