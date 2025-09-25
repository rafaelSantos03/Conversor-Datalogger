# 📈 Conversor Datalogger
Dataloger é um sistema web desenvolvido para facilitar o gerenciamento e análise de dados de temperatura e umidade coletados em uma planta experimental. Ele permite que os usuários importem, filtrem e gerem relatórios rapidamente, automatizando processos que antes eram manuais e demorados.
__
# 🔧 O que este script faz?
Importação de dados: permite carregar arquivos com informações coletadas nos sensores.
__
-Filtragem e organização: filtra os dados por período, tipo de sensor ou experimento.

Geração de relatórios: cria PDFs com gráficos e tabelas automaticamente.

Interface web intuitiva: usuário interage facilmente com o sistema sem necessidade de programação.

Eficiência: processos que antes levavam horas agora são concluídos em segundos.

O script percorre todos os dados na aba "Minima_Maxima_Datalogger" (da linha 2 até a última linha com dados). Ele verifica os valores de temperatura (coluna B) e umidade (coluna C) para cada data (coluna A) e calcula: - Temperatura Máxima: O maior valor de temperatura para cada data. - Temperatura Mínima: O menor valor de temperatura para cada data. - Umidade Máxima: O maior valor de umidade para cada data. - Umidade Mínima: O menor valor de umidade para cada data.

Armazenamento de Resultados:

Para cada data única, o script armazena os valores máximos e mínimos de temperatura e umidade.
Esses resultados são organizados em uma nova aba, que é criada com o nome "Resultado dd-mm-yyyy" (onde dd-mm-yyyy é a data atual no formato dia-mês-ano).
Criação do Gráfico:

Após gerar os dados na nova aba, o script cria um gráfico de linha para mostrar visualmente as temperaturas e umidades máximas e mínimas.
O gráfico exibe a data no eixo X e os valores no eixo Y.
# 💡 Como funciona na prática:
Primeiro Passo: Você clica no botão "Gerar Resultado" na aba "Minima_Maxima_Datalogger".

Segundo Passo: O script percorre os dados dessa aba e calcula as temperaturas e umidades máximas e mínimas para cada data.
Terceiro Passo: O script cria uma nova aba com o nome "Resultado dd-mm-yyyy" (data atual) e preenche essa aba com os valores calculados.
Quarto Passo: O script cria automaticamente um gráfico de linha para exibir visualmente esses resultados.
__
# 👌 Resultado Esperado:
Nova Aba Criada: Uma nova aba chamada "Resultado dd-mm-yyyy". Tabela de Resultados: A aba contém uma tabela com as colunas: - Data - Temperatura Máxima (°C) - Temperatura Mínima (°C) - Umidade Máxima (%) - Umidade Mínima (%) Gráfico: Um gráfico de linha gerado automaticamente, exibindo a variação das temperaturas e umidades.
__
# 🚀 Benefícios
Redução de erros humanos: a automação garante que os dados sejam processados corretamente.
Aumento da produtividade: analises que levavam horas agora são instantâneas.
Padronização de processos: todas as operações seguem um fluxo definido e confiável.
__
# 🌐 Tecnologias Utilizadas
Frontend: HTML e CSS (separados para facilitar manutenção e personalização)
Backend: Python
Armazenamento: arquivos Excel para persistência de dados
Bibliotecas: ferramentas Python para manipulação de dados e geração de PDFs
__
# ✅ Em Resumo:
O script tem a função de coletar os dados de temperatura e umidade de cada dia, calcular as variações (máximas e mínimas) e gerar um relatório com esses resultados, além de criar um gráfico visualizando essas variações. Isso é útil para análises rápidas sobre as mudanças de temperatura e umidade ao longo do tempo.

O Dataloger foi criado para otimizar o trabalho do time de pesquisa e desenvolvimento, permitindo uma análise de dados mais rápida, segura e eficiente, além de fornecer relatórios padronizados para tomada de decisão.

Projeto veio da necessidade relatada pelos colaboradores da Planta experimental da fazenda experimental 🌱.
__
# 🧑‍💻 Projeto Criado e desenvolvido por.

### Rafael Santos.


<img width="288" height="288" alt="unnamed" src="https://github.com/user-attachments/assets/c99d9e34-0d1d-438d-9b62-aec526224769" />





🔍 linkedin: https://www.linkedin.com/in/rafaelcruzdossantos/


#### Obrigado! 😄
