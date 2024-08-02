# Automação de Atualização Semanal do Power BI

Os códigos deste repositório têm como objetivo principal simplificar a tarefa de atualização semanal do Power BI dos Programas, automatizando o início do processo de acesso aos documentos, formatação das planilhas e manipulações iniciais das informações no Excel.

## Estrutura do Repositório

O repositório contém três arquivos diferentes em Python:
- **config.py**: Responsável pelo armazenamento do caminho do webdriver do Microsoft Edge e das credenciais do SPM.
- **realizado.py**: Responsável pelo download do arquivo de dados realizados.
- **planejado.py**: Responsável pelo download do arquivo de dados planejados.
- **indicadores.py**: Responsável pelo download do arquivo de indicadores.

Cada arquivo é projetado para realizar operações específicas, dependendo do tipo de dados que estão sendo processados.

## Funcionalidades dos Scripts

1. **Download e Conversão**: Cada script Python realiza o download do arquivo respectivo, transformando-o de .csv para .xlsx. O arquivo config.py armazena informações importantes para o processo.
2. **Formatação**: Após a conversão, os scripts formatam as linhas e colunas, incluindo:
   - Inserção de cores
   - Definição de altura das linhas e largura das colunas
   - Adição de bordas
3. **Salvamento dos Arquivos**: Os arquivos .xlsx formatados são salvos nas seguintes pastas:
   - **Realizado**: `Monitoramento e Avaliação/Relatório de Metas/Mensal/Realizado/2024/Realizado - xx/yy.xlsx` (sendo xx o dia do salvamento e yy o mês do salvamento)
   - **Planejado**: `Monitoramento e Avaliação/Relatório de Metas/Mensal/Planejado/Planejado - xx/yy.xlsx` (sendo xx o dia do salvamento e yy o mês do salvamento)
   - **Indicadores**: `Monitoramento e Avaliação/Relatório de Metas/Mensal/Indicadores/Indicadores - xx/yy.xlsx` (sendo xx o dia do salvamento e yy o mês do salvamento)
  
## Atenção
Como fazemos acesso ao SPM, é importante que suas credenciais (login e senha) estejam disponíveis no código, para que o Python possa acessar o sistema com seu acesso. Depois de clonar o projeto, você deve inseri-las em no arquivo config.py, em USERNAME e PASSWORD. 

## Manipulação na planilha

- Os arquivos "Realizado" e "Planejado" incluem uma nova coluna chamada "AREA2", que é o resultado da concatenação das colunas "AREA" e "ID".
- Os arquivos "Realizado" e "Planejado" incluem uma nova coluna chamada "STATUS2", indicando os status de cada registro, realizando o procv com a planilha anterior de correspondência.
- A coluna "Data_Hora" é separada em outras duas: "Data" e "Hora"

